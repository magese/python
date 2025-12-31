import pdfplumber
import re
import pandas as pd
from collections import defaultdict

def extract_data_with_multiple_lines(pdf_path):
    """
    从PDF中提取表格数据 - 处理同一收据号的多行记录
    """
    all_data = []
    columns = [
        "Date", "Time", "Room No.", "Name", "Trn. Code", "Description",
        "Check No. Exp. Date", "Receipt No.", "Currency", "Debit",
        "Credit", "Cash Id User Name", "Supplement/Reference/Credit Card No."
    ]

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            print(f"处理第 {page_num + 1} 页...")

            text = page.extract_text()
            if not text:
                continue

            # 按行分割文本
            lines = text.split('\n')

            i = 0
            while i < len(lines):
                line = lines[i].strip()

                # 检查是否是交易行（以日期开头）
                if re.match(r'^\d{2}-\d{2}-\d{2}\s+\d{2}:\d{2}\s+\d+', line):
                    # 解析交易行
                    transactions, skip_lines = parse_transaction_line_with_multiple_records(line, lines, i)

                    if transactions:
                        all_data.extend(transactions)
                        i += skip_lines
                    else:
                        i += 1
                else:
                    i += 1

    return pd.DataFrame(all_data, columns=columns)

def parse_transaction_line_with_multiple_records(line, lines, current_index):
    """
    解析交易行，处理同一收据号的多行记录
    """
    try:
        transactions = []

        # 首先解析当前行
        transaction_data, supplement_info, skip_count = parse_single_transaction_line(line, lines, current_index)
        if transaction_data:
            transactions.append(transaction_data)

            # 检查后续行是否属于同一收据号的记录
            # 后续行可能没有完整的日期时间信息，但可能有交易代码和金额

            next_index = current_index + 1
            supplement_lines_count = 0

            while next_index < len(lines):
                next_line = lines[next_index].strip()

                # 如果是新交易记录或空行，停止
                if re.match(r'^\d{2}-\d{2}-\d{2}\s+\d{2}:\d{2}', next_line) or not next_line:
                    break

                # 检查是否可能是同一收据号的另一行记录
                # 可能包含交易代码和金额信息
                if re.search(r'\d{5,6}\s+[A-Za-z\s\-]+?\s+(?:MOP|HKD|USD|CNY)\s+[\d,]+\.\d{2}\s+[\d,]+\.\d{2}', next_line):
                    # 尝试解析为附加交易记录
                    additional_transaction = parse_additional_transaction_record(
                        next_line, lines, next_index,
                        transaction_data[0],  # date
                        transaction_data[1],  # time
                        transaction_data[2],  # room_no
                        transaction_data[3],  # name
                        transaction_data[7]   # receipt_no
                    )

                    if additional_transaction:
                        transactions.append(additional_transaction)
                        supplement_lines_count += 1
                        next_index += 1
                        continue

                # 如果不是交易记录，可能是补充信息
                next_index += 1

            # 计算总共跳过的行数
            total_skip = skip_count + supplement_lines_count

            return transactions, total_skip

    except Exception as e:
        print(f"解析多行交易时出错: {line[:100]}...")
        print(f"错误: {e}")

    return [], 1

def parse_single_transaction_line(line, lines, current_index):
    """
    解析单行交易记录
    """
    try:
        # 分割行
        parts = line.split()
        if len(parts) < 10:
            return None, "", 1

        # 基础字段
        date = parts[0]
        time = parts[1]
        room_no = parts[2]

        # 查找交易代码（5-6位数字）
        trn_code_index = -1
        for i, part in enumerate(parts):
            if re.match(r'^\d{5,6}$', part) and i > 2:
                trn_code_index = i
                break

        if trn_code_index == -1:
            return None, "", 1

        # 名称是从索引3到交易代码索引-1的部分
        name_parts = parts[3:trn_code_index]
        name = " ".join(name_parts)
        trn_code = parts[trn_code_index]

        # 提取描述和收据号
        description_parts = []
        receipt_no = ""

        # 从交易代码后开始扫描
        j = trn_code_index + 1
        while j < len(parts):
            part = parts[j]

            # 检查是否是收据号（GLP开头且后面有足够长的数字）
            if part.startswith("GLP") and len(part) > 10:
                # 检查是否真的是收据号（包含足够的数字）
                if re.search(r'GLP\d{12,}', part):
                    receipt_no = part
                    j += 1
                    break

            # 检查是否是货币（如果是货币，说明没有收据号）
            if part in ["MOP", "HKD", "USD", "CNY"]:
                break

            # 添加到描述
            description_parts.append(part)
            j += 1

        description = " ".join(description_parts)

        # 继续查找货币、借方、贷方、现金ID
        currency = ""
        debit = ""
        credit = ""
        cash_id = ""

        # 从当前位置开始查找
        while j < len(parts):
            part = parts[j]

            # 货币
            if part in ["MOP", "HKD", "USD", "CNY"] and not currency:
                currency = part
            # 借方金额
            elif re.match(r'^[\d,]+\.\d{2}$', part) and not debit:
                debit = part.replace(',', '')
            # 贷方金额
            elif re.match(r'^[\d,]+\.\d{2}$', part) and not credit:
                credit = part.replace(',', '')
            # 现金ID（包含连字符，且不是金额）
            elif "-" in part and not re.match(r'^[\d,]+\.\d{2}$', part) and not cash_id:
                cash_id = part

            j += 1

        # 如果没有找到货币，尝试从后面找
        if not currency:
            for part in parts:
                if part in ["MOP", "HKD", "USD", "CNY"]:
                    currency = part
                    break

        # 清理描述
        description = clean_description_text(description)

        # Check No. Exp. Date 字段为空
        check_no = ""

        # 提取补充信息
        supplement_info = extract_supplement_info_for_line(lines, current_index)

        # 计算跳过的行数
        skip_lines = 1 + count_supplement_lines_for_line(lines, current_index)

        return [date, time, room_no, name.strip(), trn_code, description.strip(),
                check_no, receipt_no, currency, debit, credit, cash_id.strip()], supplement_info, skip_lines

    except Exception as e:
        print(f"解析单行交易时出错: {line[:100]}...")
        print(f"错误: {e}")
        return None, "", 1

def parse_additional_transaction_record(line, lines, current_index, date, time, room_no, name, receipt_no):
    """
    解析同一收据号的附加交易记录
    """
    try:
        # 分割行
        parts = line.split()

        # 查找交易代码（5-6位数字）
        trn_code_index = -1
        for i, part in enumerate(parts):
            if re.match(r'^\d{5,6}$', part):
                trn_code_index = i
                break

        if trn_code_index == -1:
            return None

        # 从交易代码前可能有一些描述性文字
        description_parts = parts[:trn_code_index]
        trn_code = parts[trn_code_index]

        # 从交易代码后开始提取描述和金额信息
        desc_start = trn_code_index + 1
        description_parts2 = []
        currency = ""
        debit = ""
        credit = ""

        j = desc_start
        while j < len(parts):
            part = parts[j]

            # 检查是否是货币
            if part in ["MOP", "HKD", "USD", "CNY"] and not currency:
                currency = part
                j += 1
                continue

            # 检查是否是金额
            if re.match(r'^[\d,]+\.\d{2}$', part):
                if not debit:
                    debit = part.replace(',', '')
                elif not credit:
                    credit = part.replace(',', '')
                else:
                    break
                j += 1
                continue

            # 添加到描述
            description_parts2.append(part)
            j += 1

        # 合并描述
        description = " ".join(description_parts + description_parts2)
        description = clean_description_text(description)

        # 如果没有找到货币，使用默认值
        if not currency:
            currency = "MOP"

        # 如果没有找到贷方，设为0.00
        if not credit:
            credit = "0.00"

        # Check No. Exp. Date 字段为空
        check_no = ""

        # 现金ID可能为空或使用默认值
        cash_id = ""

        # 提取补充信息
        supplement_info = extract_supplement_info_for_line(lines, current_index)

        return [date, time, room_no, name.strip(), trn_code, description.strip(),
                check_no, receipt_no, currency, debit, credit, cash_id, supplement_info]

    except Exception as e:
        print(f"解析附加交易记录时出错: {line[:100]}...")
        print(f"错误: {e}")
        return None

def clean_description_text(description):
    """
    清理描述文本
    """
    if not description:
        return ""

    # 移除补充信息中常见的内容
    patterns_to_remove = [
        r'Comp transfer from Window \d+ to \d+.*',
        r'Room#.*',
        r'CHECK#.*',
        r'\[.*?\]',
        r'=>.*',
        r'#\d+',
    ]

    cleaned = description
    for pattern in patterns_to_remove:
        cleaned = re.sub(pattern, '', cleaned)

    # 清理多余的空白
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()

    return cleaned

def extract_supplement_info_for_line(lines, current_index):
    """
    提取补充信息
    """
    supplement_lines = []

    # 检查当前行之后的行是否为补充信息
    next_index = current_index + 1
    while next_index < len(lines):
        next_line = lines[next_index].strip()

        # 如果下一行是新的交易记录，停止收集
        if re.match(r'^\d{2}-\d{2}-\d{2}\s+\d{2}:\d{2}', next_line):
            break

        # 如果下一行是空白行，停止收集
        if not next_line:
            break

        # 如果是交易代码标题行，停止收集
        if "Transaction Code" in next_line:
            break

        # 检查是否是另一条交易记录（包含交易代码和金额）
        if re.search(r'\d{5,6}\s+[A-Za-z\s\-]+?\s+(?:MOP|HKD|USD|CNY)\s+[\d,]+\.\d{2}\s+[\d,]+\.\d{2}', next_line):
            break

        # 添加补充信息行
        supplement_lines.append(next_line)
        next_index += 1

    # 合并补充信息行
    if supplement_lines:
        return " | ".join(supplement_lines)

    return ""

def count_supplement_lines_for_line(lines, current_index):
    """
    计算补充信息行数
    """
    count = 0
    next_index = current_index + 1

    while next_index < len(lines):
        next_line = lines[next_index].strip()

        if re.match(r'^\d{2}-\d{2}-\d{2}\s+\d{2}:\d{2}', next_line):
            break

        if not next_line or "Transaction Code" in next_line:
            break

        # 检查是否是另一条交易记录
        if re.search(r'\d{5,6}\s+[A-Za-z\s\-]+?\s+(?:MOP|HKD|USD|CNY)\s+[\d,]+\.\d{2}\s+[\d,]+\.\d{2}', next_line):
            break

        count += 1
        next_index += 1

    return count

def process_pdf_complete(pdf_path, csv_path):
    """
    完整处理PDF文件
    """
    print(f"开始处理PDF文件: {pdf_path}")

    # 使用改进的方法提取数据
    print("使用多行记录解析方法提取数据...")
    df = extract_data_with_multiple_lines(pdf_path)

    if df.empty or len(df) < 100:  # 如果提取的记录太少
        print("提取的记录较少，尝试备用方法...")
        df = extract_with_pattern_matching(pdf_path)

    # 数据清洗和验证
    df = clean_and_validate_complete(df)

    # 保存为CSV
    df.to_csv(csv_path, index=False, encoding='utf-8-sig')
    print(f"数据已保存到: {csv_path}")
    print(f"总记录数: {len(df)}")

    # 检查特定收据号的数据
    check_specific_receipts(df)

    return df

def extract_with_pattern_matching(pdf_path):
    """
    使用模式匹配提取数据
    """
    all_data = []
    columns = [
        "Date", "Time", "Room No.", "Name", "Trn. Code", "Description",
        "Check No. Exp. Date", "Receipt No.", "Currency", "Debit",
        "Credit", "Cash Id User Name", "Supplement/Reference/Credit Card No."
    ]

    # 编译正则表达式模式
    # 模式：匹配完整交易记录
    pattern = re.compile(
        r'(\d{2}-\d{2}-\d{2})\s+'  # 日期
        r'(\d{2}:\d{2})\s+'  # 时间
        r'(\d+)\s+'  # 房间号
        r'(.+?)\s+'  # 名称（非贪婪）
        r'(\d{5,6})\s+'  # 交易代码（5-6位）
        r'(.+?)\s+'  # 描述（非贪婪）
        r'(?:(GLP\d{12,})\s+)?'  # 可选的收据号
        r'(MOP|HKD|USD|CNY)\s+'  # 货币
        r'([\d,]+\.\d{2})\s+'  # 借方
        r'([\d,]+\.\d{2})\s+'  # 贷方
        r'(.+)'  # 现金ID和其他
    )

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            print(f"模式匹配处理第 {page_num + 1} 页...")

            text = page.extract_text()
            if not text:
                continue

            lines = text.split('\n')

            i = 0
            while i < len(lines):
                line = lines[i].strip()

                # 跳过非交易行
                if not line or "Transaction Code" in line or "Filter" in line:
                    i += 1
                    continue

                # 尝试匹配模式
                match = pattern.search(line)
                if match:
                    groups = match.groups()
                    date, time, room_no, name, trn_code, description, receipt_no, currency, debit, credit, rest = groups

                    # 清理数据
                    description = clean_description_text(description)
                    debit = debit.replace(',', '') if debit else "0.00"
                    credit = credit.replace(',', '') if credit else "0.00"

                    # 从rest中提取现金ID
                    cash_id = ""
                    if rest:
                        # 查找包含连字符的部分作为现金ID
                        cash_match = re.search(r'([\w-]+-\w+)', rest)
                        if cash_match:
                            cash_id = cash_match.group(1)

                    # 提取补充信息
                    supplement_info = extract_supplement_info_for_line(lines, i)

                    all_data.append([
                        date, time, room_no, name.strip(), trn_code, description,
                        "", receipt_no if receipt_no else "", currency, debit, credit,
                        cash_id, supplement_info
                    ])

                    # 跳过补充信息行
                    skip_lines = 1 + count_supplement_lines_for_line(lines, i)
                    i += skip_lines
                else:
                    i += 1

    return pd.DataFrame(all_data, columns=columns)

def clean_and_validate_complete(df):
    """
    完整的数据清洗和验证
    """
    if df.empty:
        return df

    # 确保所有字段都是字符串
    for col in df.columns:
        df[col] = df[col].astype(str).fillna('')

    # 清理空白
    for col in df.columns:
        df[col] = df[col].str.strip()

    # 统一现金ID格式
    df['Cash Id User Name'] = df['Cash Id User Name'].apply(
        lambda x: x if x and x != 'nan' else ''
    )

    # 检查并修复收据号
    def fix_receipt_no(receipt):
        if not receipt or receipt == 'nan':
            return ""

        # 提取GLP开头的完整收据号
        match = re.search(r'(GLP\d{12,})', receipt)
        if match:
            return match.group(1)

        return receipt

    df['Receipt No.'] = df['Receipt No.'].apply(fix_receipt_no)

    # 清理描述字段
    def clean_desc_final(desc):
        if not desc or desc == 'nan':
            return ""

        # 移除常见错误内容
        patterns_to_remove = [
            r'Comp transfer from Window.*',
            r'Room#.*',
            r'CHECK#.*',
            r'\[.*?\]',
            r'=>.*',
            r'#\d+',
            r'GLP\d{12,}.*',
        ]

        cleaned = desc
        for pattern in patterns_to_remove:
            cleaned = re.sub(pattern, '', cleaned)

        # 清理空白
        cleaned = re.sub(r'\s+', ' ', cleaned).strip()

        # 如果描述太短，可能是提取错误
        if len(cleaned) < 3:
            return ""

        return cleaned

    df['Description'] = df['Description'].apply(clean_desc_final)

    return df

def check_specific_receipts(df):
    """
    检查特定收据号的数据
    """
    # 检查收据号 GLP37740720231209211300
    target_receipt = "GLP37740720231209211300"
    matching_rows = df[df['Receipt No.'] == target_receipt]

    print(f"\n收据号 {target_receipt} 的记录数: {len(matching_rows)}")

    if len(matching_rows) > 0:
        print("找到的记录:")
        for idx, row in matching_rows.iterrows():
            print(f"  交易代码: {row['Trn. Code']}, 描述: {row['Description']}, "
                  f"借方: {row['Debit']}, 贷方: {row['Credit']}")

    # 检查其他常见收据号
    common_receipts = df['Receipt No.'].value_counts().head(10)
    print(f"\n最常见的收据号:")
    for receipt, count in common_receipts.items():
        if receipt and receipt != '':
            print(f"  {receipt}: {count} 条记录")

def main():
    # 设置文件路径
    pdf_file = r"C:\Users\Magese\Desktop\data2.pdf"
    csv_file = r"C:\Users\Magese\Desktop\output14.csv"

    try:
        # 处理PDF并保存为CSV
        df = process_pdf_complete(pdf_file, csv_file)

        # 显示统计信息
        print("\n数据统计:")
        print(f"总交易数: {len(df)}")
        if not df.empty:
            print(f"日期范围: {df['Date'].min()} 到 {df['Date'].max()}")
            print(f"交易代码种类: {df['Trn. Code'].nunique()}")
            print(f"货币种类: {df['Currency'].nunique()}")

            # 显示有收据号和无收据号的记录数
            with_receipt = df[df['Receipt No.'] != '']
            print(f"有收据号的记录: {len(with_receipt)}")
            print(f"无收据号的记录: {len(df) - len(with_receipt)}")

            # 检查重复的收据号
            receipt_counts = df['Receipt No.'].value_counts()
            multi_receipts = receipt_counts[receipt_counts > 1]
            print(f"多次出现的收据号数量: {len(multi_receipts)}")

    except FileNotFoundError:
        print(f"错误: 找不到文件 {pdf_file}")
    except Exception as e:
        print(f"处理过程中发生错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
