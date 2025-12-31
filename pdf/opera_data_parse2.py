import pdfplumber
import pandas as pd
import re
import os
from datetime import datetime

class PDFTransactionParser:
    def __init__(self):
        self.transaction_patterns = {
            # 房间相关
            '10000': 'Room Charge',
            '10010': 'Room Charge - Net',
            '10101': 'Room Charge - Net',
            '10107': 'Room Charge - Manual Net',
            '11000': 'Room Service Charge',
            '11001': 'Comp Room Charge Net',
            '110010': 'Comp Room Charge Net',
            '110017': 'Comp Room Charge - Manual Net',
            '111000': 'Comp Room Service Charge',
            '112000': 'Comp Room Tax',
            '12000': 'Room Tax',
            '12007': 'Room Tax - Manual',

            # 餐饮相关
            '20300': 'GLP Lobby Lounge Food Breakfast',
            '20301': 'GLP Lobby Lounge Beverage Breakfast',
            '20308': 'GLP Lobby Lounge Service Charge Breakfast',
            '20309': 'GLP Lobby Lounge Tax Breakfast',
            '20310': 'GLP Lobby Lounge Food Lunch',
            '20311': 'GLP Lobby Lounge Beverage Lunch',
            '20318': 'GLP Lobby Lounge Service Charge Lunch',
            '20319': 'GLP Lobby Lounge Tax Lunch',
            '20320': 'GLP Lobby Lounge Food Afternoon Tea',
            '20321': 'GLP Lobby Lounge Beverage Afternoon Tea',
            '20328': 'GLP Lobby Lounge Service Charge Afternoon Tea',
            '20329': 'GLP Lobby Lounge Tax Afternoon Tea',

            # 八珍坊
            '20610': 'Eight Treasures Food Lunch',
            '20611': 'Eight Treasures Beverage Lunch',
            '20618': 'Eight Treasures Service Charge Lunch',
            '20620': 'Eight Treasures Food Afternoon Tea',
            '20628': 'Eight Treasures Service Charge Afternoon Tea',
            '20630': 'Eight Treasures Food Dinner',
            '20631': 'Eight Treasures Beverage Dinner',
            '20638': 'Eight Treasures Service Charge Dinner',

            # 武藏火锅
            '20830': 'Wulso Hot Pot Food Dinner',
            '20831': 'Wulso Hot Pot Beverage Dinner',
            '20838': 'Wulso Hot Pot Service Charge Dinner',

            # 其他餐厅
            '21130': 'Kulu Kulu Izakaya Food Dinner',
            '21138': 'Kulu Kulu Izakaya Service Charge Dinner',

            # GLP Cafe
            '21200': 'GLP Cafe Food Breakfast',
            '21201': 'GLP Cafe Beverage Breakfast',
            '21208': 'GLP Cafe Service Charge Breakfast',
            '21210': 'GLP Cafe Food Lunch',
            '21218': 'GLP Cafe Service Charge Lunch',
            '21220': 'GLP Cafe Food Afternoon Tea',
            '21228': 'GLP Cafe Service Charge Afternoon Tea',
            '21230': 'GLP Cafe Food Dinner',
            '21231': 'GLP Cafe Beverage Dinner',
            '21238': 'GLP Cafe Service Charge Dinner',

            # 茶楼
            '21310': 'Chalou Food Lunch',
            '21318': 'Chalou Service Charge Lunch',
            '21330': 'Chalou Food Dinner',
            '21338': 'Chalou Service Charge Dinner',

            # 免单交易
            '220300': 'Comp GLP Lobby Lounge Food',
            '220301': 'Comp GLP Lobby Lounge Beverage',
            '220308': 'Comp GLP Lobby Lounge Service Charge',
            '220309': 'Comp GLP Lobby Lounge Tax',
            '220600': 'Comp Eight Treasures Food',
            '220601': 'Comp Eight Treasures Beverage',
            '220608': 'Comp Eight Treasures Service Charge',
            '220800': 'Comp Wulse Hot Pot Food',
            '220801': 'Comp Wulse Hot Pot Beverage',
            '220808': 'Comp Wulse Hot Pot Service Charge',
            '221200': 'Comp GLP Cafe Food',
            '221201': 'Comp GLP Cafe Beverage',
            '221208': 'Comp GLP Cafe Service Charge',
            '221300': 'Comp Chalou Food',
            '221308': 'Comp Chalou Service Charge',
        }

    def parse_pdf_to_dataframe(self, pdf_path):
        """
        解析PDF文件并返回DataFrame
        """
        print(f"开始解析PDF文件: {pdf_path}")

        # 存储所有解析的行
        all_rows = []

        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)

            for page_num, page in enumerate(pdf.pages, 1):
                print(f"正在解析第 {page_num}/{total_pages} 页...")

                # 提取表格或文本
                tables = page.extract_tables()
                text = page.extract_text()

                # 方法1: 尝试提取表格
                if tables and len(tables) > 0:
                    for table in tables:
                        self._process_table(table, all_rows)

                # 方法2: 如果表格提取失败，使用文本解析
                elif text:
                    self._process_text(text, page_num, all_rows)

        # 转换为DataFrame
        df = pd.DataFrame(all_rows, columns=[
            "Date", "Time", "Room No.", "Name",
            "Transaction Code", "Description",
            "Currency", "Debit", "Credit",
            "Cash ID", "User Name", "Supplement Info"
        ])

        return df

    def _process_table(self, table, all_rows):
        """处理表格数据"""
        for row in table:
            # 跳过空行和表头行
            if not row or len(row) < 5:
                continue

            # 检查是否是交易记录行
            if self._is_transaction_row(row):
                parsed_row = self._parse_table_row(row)
                if parsed_row:
                    all_rows.append(parsed_row)

    def _process_text(self, text, page_num, all_rows):
        """处理文本数据"""
        lines = text.split('\n')
        current_record = None

        for line in lines:
            line = line.strip()

            # 跳过不需要的行
            if self._should_skip_line(line):
                continue

            # 检查是否是新的交易记录
            if self._is_new_transaction(line):
                # 保存之前的记录
                if current_record:
                    all_rows.append(current_record)

                # 开始新的记录
                current_record = self._parse_transaction_line(line)

            # 如果是现有记录的补充信息
            elif current_record and self._is_supplement_line(line):
                current_record['Supplement Info'] = self._clean_text(
                    f"{current_record.get('Supplement Info', '')} {line}"
                )

    def _is_transaction_row(self, row):
        """检查是否是交易记录行"""
        # 检查是否有日期时间格式
        date_pattern = r'\d{2}-\d{2}-\d{2}'
        for cell in row[:2]:  # 检查前两个单元格
            if cell and re.search(date_pattern, str(cell)):
                return True
        return False

    def _parse_table_row(self, row):
        """解析表格行"""
        try:
            # 提取基本信息
            date_time = str(row[0]) if len(row) > 0 else ""
            room_no = str(row[1]) if len(row) > 1 else ""
            name = str(row[2]) if len(row) > 2 else ""
            trn_code = str(row[3]) if len(row) > 3 else ""
            description = str(row[4]) if len(row) > 4 else ""
            currency = str(row[5]) if len(row) > 5 else ""
            debit = str(row[6]) if len(row) > 6 else ""
            credit = str(row[7]) if len(row) > 7 else ""
            cash_id_user = str(row[8]) if len(row) > 8 else ""

            # 分割日期和时间
            date_match = re.search(r'(\d{2}-\d{2}-\d{2})', date_time)
            time_match = re.search(r'(\d{1,2}:\d{2})', date_time)

            date = date_match.group(1) if date_match else ""
            time = time_match.group(1) if time_match else ""

            # 转换日期格式
            if date:
                day, month, year = date.split('-')
                date = f"20{year}-{month}-{day}"

            # 分离Cash ID和User Name
            cash_id = ""
            user_name = ""
            if cash_id_user and '-' in cash_id_user:
                parts = cash_id_user.split('-', 1)
                cash_id = parts[0] if parts else ""
                user_name = parts[1] if len(parts) > 1 else ""

            # 清理金额字段
            debit = self._clean_amount(debit)
            credit = self._clean_amount(credit)

            # 获取标准描述
            standard_desc = self.transaction_patterns.get(trn_code, description)

            return {
                "Date": date,
                "Time": time,
                "Room No.": room_no,
                "Name": name,
                "Transaction Code": trn_code,
                "Description": standard_desc,
                "Currency": currency,
                "Debit": debit,
                "Credit": credit,
                "Cash ID": cash_id,
                "User Name": user_name,
                "Supplement Info": description if description != standard_desc else ""
            }

        except Exception as e:
            print(f"解析行时出错: {row}, 错误: {e}")
            return None

    def _is_new_transaction(self, line):
        """检查是否是新交易记录"""
        # 匹配日期时间模式
        date_pattern = r'^\d{2}-\d{2}-\d{2}\s+\d{1,2}:\d{2}'
        return re.match(date_pattern, line) is not None

    def _parse_transaction_line(self, line):
        """解析交易记录行"""
        try:
            # 使用正则表达式提取所有字段
            pattern = (
                r'(?P<date>\d{2}-\d{2}-\d{2})\s+'
                r'(?P<time>\d{1,2}:\d{2})\s+'
                r'(?P<room>\d{4})\s+'
                r'(?P<name>[^0-9]+?)\s+'
                r'(?P<trn_code>\d{3,6})\s+'
                r'(?P<desc>[^MOP]+?)\s+'
                r'MOP\s+'
                r'(?P<debit>[-\d,.]+)\s+'
                r'(?P<credit>[-\d,.]+)\s+'
                r'(?P<cash_user>\S+)'
            )

            match = re.search(pattern, line)
            if match:
                data = match.groupdict()

                # 转换日期格式
                day, month, year = data['date'].split('-')
                date = f"20{year}-{month}-{day}"

                # 清理金额
                debit = self._clean_amount(data['debit'])
                credit = self._clean_amount(data['credit'])

                # 分离Cash ID和User Name
                cash_parts = data['cash_user'].split('-', 1)
                cash_id = cash_parts[0] if cash_parts else ""
                user_name = cash_parts[1] if len(cash_parts) > 1 else ""

                # 获取标准描述
                standard_desc = self.transaction_patterns.get(data['trn_code'], data['desc'].strip())

                return {
                    "Date": date,
                    "Time": data['time'],
                    "Room No.": data['room'],
                    "Name": data['name'].strip(),
                    "Transaction Code": data['trn_code'],
                    "Description": standard_desc,
                    "Currency": "MOP",
                    "Debit": debit,
                    "Credit": credit,
                    "Cash ID": cash_id,
                    "User Name": user_name,
                    "Supplement Info": data['desc'].strip() if data['desc'].strip() != standard_desc else ""
                }

            return None

        except Exception as e:
            print(f"解析行时出错: {line}, 错误: {e}")
            return None

    def _is_supplement_line(self, line):
        """检查是否是补充信息行"""
        supplement_keywords = [
            '[NA Room]', '[Add:', '[10010->', 'Room#', 'CHECK#',
            'Comp transfer', 'Routed From', '=>'
        ]
        return any(keyword in line for keyword in supplement_keywords)

    def _should_skip_line(self, line):
        """检查是否应该跳过该行"""
        skip_keywords = [
            'Grand Lisboa Palace', 'Journal by Cashier', 'Filter',
            'From Date', 'To Date', 'Transactions All', 'Cashier All',
            'Group by', 'Sort Order', 'Room Class All', 'Include Currency',
            'Revenue NET', 'Page', 'of', 'Transaction Code', '====='
        ]

        if not line or len(line.strip()) < 2:
            return True

        return any(keyword in line for keyword in skip_keywords)

    def _clean_amount(self, amount_str):
        """清理金额字符串"""
        if not amount_str:
            return "0.00"

        # 移除逗号、空格，确保格式正确
        cleaned = str(amount_str).replace(',', '').replace(' ', '').strip()

        # 检查是否是负数
        is_negative = cleaned.startswith('-')
        if is_negative:
            cleaned = cleaned[1:]

        # 确保有两位小数
        if '.' not in cleaned:
            cleaned = f"{cleaned}.00"

        # 添加负号
        if is_negative:
            cleaned = f"-{cleaned}"

        return cleaned

    def _clean_text(self, text):
        """清理文本"""
        if not text:
            return ""
        return re.sub(r'\s+', ' ', text).strip()

    def export_to_csv(self, df, csv_path):
        """导出到CSV文件"""
        df.to_csv(csv_path, index=False, encoding='utf-8-sig')
        print(f"数据已导出到: {csv_path}")
        return df.shape[0]

def main():
    # 输入和输出文件路径
    pdf_path = r"C:\Users\Magese\Desktop\data.pdf"
    csv_path = r"C:\Users\Magese\Desktop\output_improved.csv"

    try:
        # 检查文件是否存在
        if not os.path.exists(pdf_path):
            print(f"错误: 找不到文件 {pdf_path}")
            return

        # 创建解析器并解析PDF
        parser = PDFTransactionParser()
        df = parser.parse_pdf_to_dataframe(pdf_path)

        if df.empty:
            print("警告: 没有提取到任何数据")
            return

        # 导出到CSV
        rows = parser.export_to_csv(df, csv_path)

        # 显示统计信息
        print("\n解析统计:")
        print(f"总记录数: {rows}")
        print(f"交易代码种类: {df['Transaction Code'].nunique()}")
        print(f"日期范围: {df['Date'].min()} 到 {df['Date'].max()}")

        # 显示前几行数据
        print("\n前5行数据:")
        print(df.head().to_string())

        # 按交易代码分组统计
        print("\n按交易代码统计:")
        summary = df.groupby(['Transaction Code', 'Description']).agg({
            'Debit': 'sum',
            'Credit': 'sum'
        }).reset_index()
        print(summary.to_string())

    except Exception as e:
        print(f"解析过程中发生错误: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
