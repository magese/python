import camelot
import pandas as pd
import re
import os
from datetime import datetime

def extract_data_from_pdf(pdf_path, csv_path):
    """
    从PDF中提取表格数据并保存为CSV文件
    """
    if not os.path.exists(pdf_path):
        print(f"错误: PDF文件不存在 - {pdf_path}")
        return

    print(f"正在读取PDF文件: {pdf_path}")

    try:
        # 使用更精确的参数
        tables = camelot.read_pdf(
            pdf_path,
            pages='all',
            flavor='lattice',  # 使用格子模式，更适合结构化表格
            line_scale=40,
            copy_text=['h', 'v'],
            split_text=True,
            flag_size=True,
            strip_text='\n',
        )
        print(f"找到 {len(tables)} 个表格区域")
    except Exception as e:
        print(f"读取PDF文件时出错: {e}")
        try:
            # 回退到stream模式
            tables = camelot.read_pdf(
                pdf_path,
                pages='all',
                flavor='stream',
                row_tol=10,
                edge_tol=50,
                strip_text='\n'
            )
            print(f"使用stream模式找到 {len(tables)} 个表格区域")
        except Exception as e2:
            print(f"再次尝试失败: {e2}")
            return

    # 收集所有记录
    all_records = []

    for table_num, table in enumerate(tables):
        print(f"处理表格区域 {table_num+1}/{len(tables)}...")

        df = table.df
        if df.empty:
            continue

        # 处理当前表格
        table_records = process_table(df)
        all_records.extend(table_records)

    print(f"共提取 {len(all_records)} 条记录")

    if all_records:
        final_df = pd.DataFrame(all_records)
        final_df = clean_and_format_data(final_df)
        final_df.to_csv(csv_path, index=False)
        print(f"数据已保存到 {csv_path}")

        # 显示样本
        print("\n前10条记录样本:")
        print(final_df.head(10).to_string(index=False))

        # 统计信息
        print(f"\n数据统计:")
        print(f"总记录数: {len(final_df)}")

        # 检查字段提取情况
        print("\n字段提取统计:")
        for field in ['Date', 'Time', 'Room No.', 'Trn. Code', 'Description',
                      'Check No.Exp. Date', 'Receipt No.', 'Debit', 'Credit', 'User Name']:
            if field in final_df.columns:
                non_empty = final_df[field].apply(lambda x: str(x).strip() != '').sum()
                print(f"  {field}: {non_empty}/{len(final_df)} 条非空")
    else:
        print("没有提取到有效记录")

def process_table(df):
    """
    处理单个表格，提取记录
    """
    records = []
    current_record = {}
    supplement_lines = []

    for idx, row in df.iterrows():
        # 获取行文本
        row_text = ' '.join([str(cell).strip() for cell in row if str(cell).strip()])

        # 跳过标题行和空行
        if not row_text or any(keyword in row_text for keyword in
                               ['Journal by Cashier', 'Transaction Code', 'Date    Time']):
            continue

        # 检查是否是新的交易记录
        is_new_record = False

        # 检查日期模式 (DD-MM-YY HH:MM)
        date_time_match = re.search(r'(\d{2}-\d{2}-\d{2})\s+(\d{2}:\d{2})', row_text)
        if date_time_match:
            is_new_record = True
        else:
            # 检查是否只有日期
            date_only_match = re.search(r'^\s*(\d{2}-\d{2}-\d{2})\s*$', row_text)
            if date_only_match:
                is_new_record = True

        if is_new_record:
            # 保存前一个记录
            if current_record:
                if supplement_lines:
                    current_record['Supplement/Reference/Credit Card No.'] = ' '.join(supplement_lines)
                records.append(current_record)

            # 开始新记录
            current_record = create_new_record()
            supplement_lines = []

            # 提取日期和时间
            if date_time_match:
                current_record['Date'] = date_time_match.group(1)
                current_record['Time'] = date_time_match.group(2)
            elif date_only_match:
                current_record['Date'] = date_only_match.group(1)

        # 处理当前行的内容
        process_row_content(row_text, current_record, supplement_lines)

    # 添加最后一个记录
    if current_record:
        if supplement_lines:
            current_record['Supplement/Reference/Credit Card No.'] = ' '.join(supplement_lines)
        records.append(current_record)

    return records

def create_new_record():
    """创建新的记录模板"""
    return {
        'Date': '',
        'Time': '',
        'Room No.': '',
        'Name': 'China Tennis CTA',
        'Trn. Code': '',
        'Description': '',
        'Check No.Exp. Date': '',
        'Receipt No.': '',
        'Currency': 'MOP',
        'Debit': '0.00',
        'Credit': '0.00',
        'Cash ID': '',
        'User Name': '',
        'Supplement/Reference/Credit Card No.': ''
    }

def process_row_content(row_text, record, supplement_lines):
    """
    处理行内容，提取字段
    """
    # 提取房间号
    if not record['Room No.']:
        room_match = re.search(r'\b(95\d{2})\b', row_text)
        if room_match:
            record['Room No.'] = room_match.group(1)

    # 提取交易代码
    if not record['Trn. Code']:
        # 尝试各种格式
        trn_patterns = [
            r'\b(\d{5})\b',    # 5位数字
            r'\b(\d{6})\b',    # 6位数字
            r'\b(\d{4})\b',    # 4位数字
        ]

        for pattern in trn_patterns:
            trn_match = re.search(pattern, row_text)
            if trn_match:
                code = trn_match.group(1)
                # 排除房间号和金额
                if code != record.get('Room No.', '') and not re.match(r'^\d+\.\d+$', code):
                    record['Trn. Code'] = code
                    break

    # 提取描述
    if not record['Description']:
        # 常见描述关键词
        desc_keywords = [
            'Room Charge', 'Room Service', 'Room Tax',
            'GLP Lobby Lounge', 'Eight Treasures', 'Wulso Hot Pot',
            'Kulu Kulu Izakaya', 'GLP Cafe', 'Chakou',
            'Food', 'Beverage', 'Service Charge', 'Tax',
            'Breakfast', 'Lunch', 'Dinner', 'Afternoon Tea',
            'Comp', 'Manual', 'Nett'
        ]

        for keyword in desc_keywords:
            if keyword in row_text:
                # 提取包含关键词的描述
                start_idx = row_text.find(keyword)
                # 取关键词后的50个字符或到行尾
                end_idx = min(start_idx + len(keyword) + 50, len(row_text))
                description = row_text[start_idx:end_idx]

                # 清理描述
                description = re.sub(r'\s+', ' ', description).strip()
                description = re.sub(r'[^\w\s/-]', '', description)  # 移除特殊字符

                if len(description) > 5:
                    record['Description'] = description
                    break

    # 提取Check No. (完整GLP编号)
    if not record['Check No.Exp. Date']:
        # 匹配完整GLP编号
        glp_patterns = [
            r'(GLP\d{6,20})',  # GLP后跟6-20位数字
            r'(GLP\d+2023\d+)' # GLP后跟数字，包含2023
        ]

        for pattern in glp_patterns:
            glp_match = re.search(pattern, row_text)
            if glp_match:
                record['Check No.Exp. Date'] = glp_match.group(1)
                break

    # 提取Receipt No. (#后跟数字，但不是房间号的一部分)
    if not record['Receipt No.']:
        # 匹配#开头后跟数字，但排除房间号模式
        receipt_match = re.search(r'(#\d{4,6})\b', row_text)
        if receipt_match and receipt_match.group(1) != '#9503':
            record['Receipt No.'] = receipt_match.group(1)

    # 提取金额 - 改进的金额提取
    amount_pattern = r'(-?\s?\d{1,3}(?:,\d{3})*(?:\.\d{2})?)'
    amounts = re.findall(amount_pattern, row_text)

    for amount_str in amounts:
        clean_amount = amount_str.replace(',', '').replace(' ', '')

        # 跳过房间号、交易代码等
        if (clean_amount == record.get('Room No.', '') or
                clean_amount == record.get('Trn. Code', '') or
                clean_amount in ['9503', '9534', '9535']):
            continue

        try:
            amount_float = float(clean_amount)
            if abs(amount_float) > 0.01:  # 有效金额
                # 判断是借方还是贷方
                if '-' in amount_str or amount_float < 0:
                    # 负数通常是贷方
                    record['Credit'] = str(abs(amount_float))
                    record['Debit'] = '0.00'
                else:
                    # 正数通常是借方
                    record['Debit'] = clean_amount
                    record['Credit'] = '0.00'
                break
        except:
            continue

    # 提取User Name
    if not record['User Name']:
        user_patterns = [
            r'(\d{6}-[A-Z]+@[A-Z]?)',      # 163540-PEGGYLYU@
            r'(\d{3}-[A-Z]+@[A-Z]?)',      # 990-YANNISLEONG@
            r'(998-OPERA-IFC)',            # 998-OPERA-IFC
            r'(\d{6}-[A-Z]+)',             # 162558-BRYANTLEON
            r'(\d{3}-[A-Z]+@G)',           # 990-JAMESLIANG@G
            r'(\d{3}-[A-Z]+@@)'            # 990-YANNISLEONG@@
        ]

        for pattern in user_patterns:
            user_match = re.search(pattern, row_text)
            if user_match:
                record['User Name'] = user_match.group(1)
                break

    # 判断是否是补充信息行
    supplement_indicators = [
        'Room#', 'CHECK#', '=>', 'Routed', 'From',
        'Comp transfer', '[', ']', 'Add:', '10010->110010',
        'NA Room', 'QT Room'
    ]

    is_supplement = any(indicator in row_text for indicator in supplement_indicators)

    # 检查是否是描述行（已处理）
    is_description = record['Description'] and record['Description'] in row_text

    # 检查是否是金额行
    amount_in_row = bool(re.search(r'\d+\.\d{2}', row_text))

    # 如果是补充信息且不是主要字段，添加到补充信息
    if is_supplement and not is_description and not amount_in_row:
        # 清理补充信息
        clean_supplement = row_text.strip()
        # 移除已提取的主要字段
        for field in ['Date', 'Time', 'Room No.', 'Trn. Code', 'Description']:
            if record.get(field):
                clean_supplement = clean_supplement.replace(str(record[field]), '')

        clean_supplement = re.sub(r'\s+', ' ', clean_supplement).strip()
        if clean_supplement and len(clean_supplement) > 3:
            supplement_lines.append(clean_supplement)

def clean_and_format_data(df):
    """
    清理和格式化提取的数据
    """
    if df.empty:
        return df

    # 1. 格式化日期
    def format_date(date_str):
        try:
            if isinstance(date_str, str) and date_str.strip():
                date_str = date_str.strip()
                # 尝试解析dd-mm-yy格式
                try:
                    dt = datetime.strptime(date_str, '%d-%m-%y')
                    return dt.strftime('%Y-%m-%d')
                except:
                    return date_str
        except:
            pass
        return date_str

    df['Date'] = df['Date'].apply(format_date)

    # 2. 清理金额字段
    def clean_amount(amount, is_debit=True):
        if isinstance(amount, str):
            amount = amount.strip()
            if not amount or amount in ['0', '0.0', '0.00']:
                return '0.00'

            # 移除逗号和空格
            amount = amount.replace(',', '').replace(' ', '')

            # 处理负数
            is_negative = amount.startswith('-')
            if is_negative:
                amount = amount[1:]

            try:
                # 转换为浮点数并格式化
                float_amount = float(amount)
                if is_negative and not is_debit:
                    float_amount = -float_amount
                return f"{abs(float_amount):.2f}"
            except:
                # 尝试提取数字
                match = re.search(r'[-]?\d+\.?\d*', amount)
                if match:
                    try:
                        return f"{abs(float(match.group())):.2f}"
                    except:
                        return '0.00'
        elif isinstance(amount, (int, float)):
            return f"{abs(float(amount)):.2f}"

        return '0.00'

    df['Debit'] = df['Debit'].apply(lambda x: clean_amount(x, True))
    df['Credit'] = df['Credit'].apply(lambda x: clean_amount(x, False))

    # 3. 清理交易代码
    df['Trn. Code'] = df['Trn. Code'].apply(
        lambda x: re.search(r'\d{4,6}', str(x)).group() if re.search(r'\d{4,6}', str(x)) else ''
    )

    # 4. 清理描述字段
    def clean_description(desc):
        if isinstance(desc, str):
            desc = desc.strip()
            # 移除多余空格
            desc = re.sub(r'\s+', ' ', desc)
            # 截断过长的描述
            if len(desc) > 100:
                desc = desc[:100] + '...'
        return desc

    df['Description'] = df['Description'].apply(clean_description)

    # 5. 清理Check No.字段 - 保留完整GLP编号
    def clean_check_no(check_no):
        if isinstance(check_no, str):
            check_no = check_no.strip()
            # 提取完整GLP编号
            glp_match = re.search(r'(GLP\d{6,})', check_no)
            if glp_match:
                return glp_match.group(1)
        return check_no

    df['Check No.Exp. Date'] = df['Check No.Exp. Date'].apply(clean_check_no)

    # 6. 清理Receipt No.字段
    def clean_receipt_no(receipt_no):
        if isinstance(receipt_no, str):
            receipt_no = receipt_no.strip()
            # 确保以#开头且后跟数字
            if receipt_no and not receipt_no.startswith('#'):
                # 检查是否是数字
                if receipt_no.isdigit() and len(receipt_no) >= 4:
                    return '#' + receipt_no
            elif receipt_no.startswith('#'):
                # 确保#后是数字
                if len(receipt_no) > 1 and receipt_no[1:].isdigit():
                    return receipt_no
        return receipt_no

    df['Receipt No.'] = df['Receipt No.'].apply(clean_receipt_no)

    # 7. 清理User Name字段
    df['User Name'] = df['User Name'].apply(
        lambda x: str(x).strip() if pd.notna(x) else ''
    )

    # 8. 清理所有文本字段
    text_fields = ['Time', 'Room No.', 'Name', 'Description',
                   'Check No.Exp. Date', 'Receipt No.', 'Currency',
                   'User Name', 'Supplement/Reference/Credit Card No.']

    for field in text_fields:
        if field in df.columns:
            df[field] = df[field].apply(lambda x: str(x).strip() if pd.notna(x) else '')

    # 9. 确保Cash ID为空
    df['Cash ID'] = ''

    # 10. 移除没有日期的记录
    df = df[df['Date'].notna() & (df['Date'] != '')]

    # 11. 重新排列列顺序
    column_order = [
        'Date', 'Time', 'Room No.', 'Name',
        'Trn. Code', 'Description', 'Check No.Exp. Date',
        'Receipt No.', 'Currency', 'Debit', 'Credit',
        'Cash ID', 'User Name', 'Supplement/Reference/Credit Card No.'
    ]

    # 确保所有列都存在
    for col in column_order:
        if col not in df.columns:
            df[col] = ''

    # 只保留需要的列
    df = df[column_order]

    # 12. 去重（基于关键字段）
    key_columns = ['Date', 'Time', 'Room No.', 'Trn. Code', 'Debit', 'Credit']
    df = df.drop_duplicates(subset=key_columns, keep='first')

    return df

def main():
    """
    主函数：从PDF提取数据并保存为CSV
    """
    # 设置输入和输出文件的绝对路径
    pdf_file_path = r"C:\Users\Magese\Desktop\data2.pdf"
    csv_file_path = r"C:\Users\Magese\Desktop\output10.csv"

    # 如果路径不存在，使用当前目录
    if not os.path.exists(pdf_file_path):
        current_dir = os.getcwd()
        pdf_files = [f for f in os.listdir(current_dir) if f.lower().endswith('.pdf')]
        if pdf_files:
            pdf_file_path = os.path.join(current_dir, pdf_files[0])
            print(f"使用找到的PDF文件: {pdf_file_path}")
        else:
            print("错误: 请指定有效的PDF文件路径")
            return

    # 提取数据
    extract_data_from_pdf(pdf_file_path, csv_file_path)

if __name__ == "__main__":
    main()
