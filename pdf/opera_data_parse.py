import pdfplumber
import csv
import re
import os
from datetime import datetime

def parse_pdf_to_csv(pdf_path, csv_path):
    """
    解析PDF文件并转换为CSV格式
    """
    print(f"开始解析PDF文件: {pdf_path}")

    # 存储所有解析的行
    all_rows = []
    total_pages = 0
    total_rows = 0

    # 定义表头
    header = [
        "Date", "Time", "Room No.", "Name",
        "Supplement/Reference/Credit Card No.", "Trn. Code",
        "Description", "Check No.Exp. Date", "Receipt No.",
        "Currency", "Debit", "Credit", "Cash ID User Name"
    ]

    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)

        for page_num, page in enumerate(pdf.pages, 1):
            print(f"正在解析第 {page_num}/{total_pages} 页...")

            # 提取页面文本
            text = page.extract_text()
            if not text:
                continue

            # 按行分割文本
            lines = text.split('\n')

            # 跳过表头和不需要的行
            skip_keywords = [
                "Grand Lisboa Palace", "Journal by Cashier and Transaction Code",
                "Date    Time    Room No.", "Filter", "From Date", "To Date",
                "Transactions All", "Cashier All", "Group by Transaction Code",
                "Sort Order Chronological", "Room ", "Room Class All",
                "Include Currency Exchange", "Revenue NET Amount",
                "firjrmbytrans", "finjrnlbytrans", "=====", "Transaction Code",
                "Total", "Page", "of", "Filter", "Debit", "Credit", "Transaction Code Total"
            ]

            i = 0
            while i < len(lines):
                line = lines[i].strip()

                # 跳过不需要的行
                if not line or any(keyword in line for keyword in skip_keywords):
                    i += 1
                    continue

                # 尝试匹配日期时间模式 (如 01-11-23  23:21 或 01-11-23 04:24)
                date_match = re.search(r'(\d{2}-\d{2}-\d{2})\s+(\d{1,2}:\d{2})', line)

                if date_match:
                    # 提取基本信息
                    date_str = date_match.group(1)  # 格式: 日-月-年
                    current_time = date_match.group(2)

                    # 将日期从 "日-月-年" 转换为 "年-月-日"
                    try:
                        # 分割日期部分
                        day, month, year = date_str.split('-')
                        # 假设20XX年
                        full_year = f"20{year}"
                        current_date = f"{full_year}-{month}-{day}"
                    except:
                        current_date = date_str  # 如果转换失败，保持原格式

                    # 提取房间号 (通常是4位数字)
                    room_match = re.search(r'\b(\d{4})\b', line)
                    current_room = room_match.group(1) if room_match else ""

                    # 提取名称
                    current_name = ""
                    if "China Tennis CTA" in line:
                        current_name = "China Tennis CTA"

                    # 提取交易代码 - 现在支持多种格式
                    # 交易代码可能是3-6位数字，如: 10010, 110010, 990200, 20300, 20610等
                    trn_code_patterns = [
                        r'\b(\d{5})\b',      # 5位数字，如10010
                        r'\b(\d{6})\b',      # 6位数字，如110010
                        r'\b(\d{3,4})\b',    # 3-4位数字，如990, 2030等
                    ]

                    current_trn_code = ""
                    for pattern in trn_code_patterns:
                        trn_match = re.search(pattern, line)
                        if trn_match:
                            # 确保不是房间号或日期的一部分
                            candidate = trn_match.group(1)
                            if candidate != current_room and candidate != year and candidate != month and candidate != day:
                                # 确保不是Cash ID的一部分 (如990-后面的990)
                                if f"{candidate}-" not in line and f"-{candidate}" not in line:
                                    current_trn_code = candidate
                                    break

                    # 如果还没找到交易代码，尝试从已知的交易代码列表中查找
                    if not current_trn_code:
                        known_trn_codes = [
                            "10000", "10010", "10101", "10107", "11000", "11001", "110010",
                            "110017", "11007", "111000", "112000", "12000", "12007", "20300",
                            "20301", "20308", "20309", "20310", "20311", "20318", "20319",
                            "20320", "20321", "20328", "20329", "20610", "20611", "20618",
                            "20620", "20628", "20630", "20631", "20638", "20830", "20831",
                            "20838", "20930", "20931", "20938", "21130", "21138", "21200",
                            "21201", "21208", "21210", "21218", "21220", "21228", "21230",
                            "21231", "21238", "21310", "21318", "21330", "21338", "220300",
                            "220301", "220308", "220309", "220600", "220601", "220608",
                            "220609", "220801", "220808", "221200", "221201", "221208",
                            "221300", "221308", "990200"
                        ]

                        for code in known_trn_codes:
                            if f" {code} " in f" {line} ":
                                current_trn_code = code
                                break

                    # 初始化其他字段
                    supplement = ""
                    description = ""
                    check_no_exp_date = ""
                    receipt_no = ""
                    currency = ""
                    debit = ""
                    credit = ""
                    cash_id_user = ""

                    # 先提取Check No.Exp. Date - GLP编号
                    # GLP编号模式: GLP后跟6-7位数字，然后是年份月份日期和时间
                    check_pattern = r'(GLP\d{6,7}\d{12,14})'
                    check_matches = re.findall(check_pattern, line)

                    # 如果找到GLP编号，从行中移除它们，避免混入description
                    clean_line = line
                    if check_matches:
                        for check_match in check_matches:
                            clean_line = clean_line.replace(check_match, "")
                        check_no_exp_date = " ".join(check_matches)

                    # 提取Cash ID User Name - 需要匹配完整的ID
                    # 特别注意：需要匹配完整的格式，包括结尾的字符
                    cash_user_patterns = [
                        r'(\d{3,6}-[A-Z]+-[A-Z]+)',           # 如 998-OPERA-IFC
                        r'(\d{3,6}-[A-Z]+[A-Z0-9@\.]+[A-Z0-9@])',  # 如 990-JAMESLIANG@G
                        r'(\d{3,6}-[A-Z]+[A-Z0-9@]+@[A-Z])',  # 如 990-YANNISLEONG@
                        r'(\d{6}-[A-Z]+@)',                   # 如 163540-PEGGYLYU@
                        r'(\d{6}-[A-Z]+)',                    # 如 162558-BRYANTLEON
                    ]

                    # 尝试从行尾开始匹配，因为Cash ID通常在行尾
                    line_for_cash_id = line
                    for pattern in cash_user_patterns:
                        cash_match = re.search(pattern, line_for_cash_id)
                        if cash_match:
                            cash_id_user = cash_match.group(1)
                            # 从clean_line中移除Cash ID，避免混入其他字段
                            clean_line = clean_line.replace(cash_id_user, "")
                            break

                    # 如果还没有找到，尝试从行中直接搜索已知模式
                    if not cash_id_user:
                        known_patterns = [
                            '998-OPERA-IFC', '990-JAMESLIANG@G', '990-YANNISLEONG@',
                            '990-ANNEWONG@G', '990-STELLAWONG@G', '163540-PEGGYLYU@',
                            '162558-BRYANTLEON', '990-JJMEISLJANG@G', '990-JAMESLIANQ@G',
                            '162558-BRYANTLEDI', '990-YANNISLEONG@', '990-JAMESLIANG@G'
                        ]
                        for pattern in known_patterns:
                            if pattern in line:
                                cash_id_user = pattern
                                clean_line = clean_line.replace(cash_id_user, "")
                                break

                    # 提取描述字段 - 从交易代码到金额之间的部分
                    if current_trn_code and current_trn_code in clean_line:
                        # 找到交易代码的位置
                        trn_index = clean_line.find(current_trn_code)

                        # 查找MOP出现的位置
                        mop_index = clean_line.find("MOP")

                        if trn_index >= 0 and mop_index > trn_index:
                            # 提取交易代码到MOP之间的内容
                            desc_part = clean_line[trn_index + len(current_trn_code):mop_index].strip()

                            # 移除可能的金额数字，确保描述纯净
                            # 移除任何看起来像金额的模式
                            desc_part = re.sub(r'[-]?\s*\d{1,3}(?:,\d{3})*\.\d{2}', '', desc_part)
                            desc_part = re.sub(r'\s+', ' ', desc_part).strip()

                            # 如果描述以标点符号开头，清理掉
                            if desc_part and desc_part[0] in [',', '.', ';', ':']:
                                desc_part = desc_part[1:].strip()

                            description = desc_part

                    # 如果描述为空，使用一个默认值
                    if not description or len(description) < 2:
                        # 尝试从原始行中提取描述，但排除已知模式
                        if current_trn_code:
                            # 查找交易代码后面的部分，直到行尾或MOP
                            trn_index = line.find(current_trn_code)
                            if trn_index >= 0:
                                # 排除GLP编号和Cash ID
                                remaining = line[trn_index + len(current_trn_code):]

                                # 移除GLP编号
                                remaining = re.sub(r'GLP\d{6,7}\d{12,14}', '', remaining)

                                # 移除Cash ID
                                if cash_id_user:
                                    remaining = remaining.replace(cash_id_user, "")

                                # 移除金额模式
                                remaining = re.sub(r'[-]?\s*\d{1,3}(?:,\d{3})*\.\d{2}', '', remaining)
                                remaining = re.sub(r'MOP', '', remaining)

                                remaining = re.sub(r'\s+', ' ', remaining).strip()

                                # 清理标点符号
                                if remaining and remaining[0] in [',', '.', ';', ':']:
                                    remaining = remaining[1:].strip()

                                if remaining and len(remaining) > 1:
                                    description = remaining

                    # 如果描述仍然为空或过短，使用一个通用描述
                    if not description or len(description) < 2:
                        # 根据交易代码判断可能的描述
                        if current_trn_code:
                            if current_trn_code.startswith('100'):
                                description = "Room Charge"
                            elif current_trn_code.startswith('110'):
                                description = "Room Service"
                            elif current_trn_code.startswith('120'):
                                description = "Room Tax"
                            elif current_trn_code.startswith('203'):
                                description = "GLP Lobby Lounge"
                            elif current_trn_code.startswith('212'):
                                description = "GLP Cafe"
                            elif current_trn_code.startswith('213'):
                                description = "Chalou"
                            elif current_trn_code.startswith('206'):
                                description = "Eight Treasures"
                            elif current_trn_code.startswith('209'):
                                description = "Wulso Hot Pot"
                            elif current_trn_code.startswith('211'):
                                description = "Kulu Kulu Izakaya"
                            elif current_trn_code.startswith('22'):
                                description = "Comp Transaction"
                            elif current_trn_code == "990200":
                                description = "Comp Settlement"
                            else:
                                description = "Transaction"

                    # 提取金额信息
                    # 查找MOP和金额模式
                    mop_match = re.search(r'(MOP)\s+([-]?\s*\d{1,3}(?:,\d{3})*\.\d{2})\s+([-]?\s*\d{1,3}(?:,\d{3})*\.\d{2})', line)

                    if mop_match:
                        currency = mop_match.group(1)
                        # 提取Debit和Credit，去除逗号
                        debit_raw = mop_match.group(2).replace(" ", "")
                        credit_raw = mop_match.group(3).replace(" ", "")

                        # 去除逗号
                        debit = debit_raw.replace(",", "") if debit_raw else ""
                        credit = credit_raw.replace(",", "") if credit_raw else ""

                    # 提取补充信息 - 通常包含在方括号[]中
                    supplement_parts = []
                    bracket_pattern = r'\[(.*?)\]'
                    bracket_matches = re.findall(bracket_pattern, line)
                    if bracket_matches:
                        supplement_parts.extend(bracket_matches)

                    # 检查后续行是否有补充信息
                    supplement_lines = []
                    j = i + 1
                    max_supplement_lines = 3  # 最多检查接下来的3行
                    lines_checked = 0

                    while j < len(lines) and lines_checked < max_supplement_lines:
                        next_line = lines[j].strip()

                        # 如果下一行是空行或包含跳过关键词，跳过
                        if not next_line or any(keyword in next_line for keyword in skip_keywords):
                            j += 1
                            lines_checked += 1
                            continue

                        # 如果下一行以方括号开始或包含补充信息关键词
                        if (next_line.startswith('[') or
                                '[NA Room]' in next_line or
                                '[Add:' in next_line or
                                '[10010->' in next_line or
                                'Room#' in next_line or
                                'CHECK#' in next_line):
                            supplement_lines.append(next_line)
                            j += 1
                            lines_checked += 1
                        else:
                            # 检查是否可能是续行（没有日期时间模式）
                            next_date_match = re.search(r'(\d{2}-\d{2}-\d{2})\s+(\d{1,2}:\d{2})', next_line)
                            if not next_date_match:
                                # 可能是上一行的续行
                                supplement_lines.append(next_line)
                                j += 1
                                lines_checked += 1
                            else:
                                break

                    # 合并补充信息行
                    if supplement_lines:
                        supplement_parts.extend(supplement_lines)

                    # 合并补充信息
                    if supplement_parts:
                        supplement = " ".join(supplement_parts)

                    # 清理补充信息中的多余空格
                    supplement = re.sub(r'\s+', ' ', supplement).strip() if supplement else ""

                    # 清理description中的多余空格
                    description = re.sub(r'\s+', ' ', description).strip() if description else ""

                    # 创建一行数据
                    row = [
                        current_date,  # 已经是年-月-日格式
                        current_time,
                        current_room,
                        current_name,
                        supplement,
                        current_trn_code,
                        description,
                        check_no_exp_date,
                        receipt_no,
                        currency,
                        debit,
                        credit,
                        cash_id_user
                    ]

                    all_rows.append(row)
                    total_rows += 1

                    # 跳过已处理的补充信息行
                    i = j
                else:
                    i += 1

    print(f"解析完成，共处理 {total_pages} 页，提取 {total_rows} 条记录")

    # 写入CSV文件
    with open(csv_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(header)
        writer.writerows(all_rows)

    print(f"CSV文件已保存到: {csv_path}")

    return total_pages, total_rows

def main():
    # 输入和输出文件路径
    pdf_path = r"C:\Users\Magese\Desktop\data2.pdf"  # 替换为你的PDF文件路径
    csv_path = r"C:\Users\Magese\Desktop\output9.csv"  # 输出CSV文件路径

    try:
        # 检查PDF文件是否存在
        if not os.path.exists(pdf_path):
            print(f"错误: 找不到文件 {pdf_path}")
            return

        # 解析PDF并保存为CSV
        pages, rows = parse_pdf_to_csv(pdf_path, csv_path)

        print("\n解析统计:")
        print(f"总页数: {pages}")
        print(f"总记录数: {rows}")
        print(f"输出文件: {csv_path}")

        # 显示前几行示例
        if rows > 0:
            print("\n前5行示例:")
            with open(csv_path, 'r', encoding='utf-8-sig') as f:
                for i, line in enumerate(f):
                    if i < 6:  # 包括表头
                        print(line.strip())

    except Exception as e:
        print(f"解析过程中发生错误: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
