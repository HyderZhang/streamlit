import streamlit as st
import pandas as pd
import math
import xlsxwriter
from io import BytesIO

def num_to_chinese(num):
    """
    将正整数转换为中文数字（支持1-10），例如 1 -> "一"，2 -> "二"，……
    """
    mapping = {
        1: "一",
        2: "二",
        3: "三",
        4: "四",
        5: "五",
        6: "六",
        7: "七",
        8: "八",
        9: "九",
        10: "十"
    }
    if num in mapping:
        return mapping[num]
    else:
        tens = num // 10
        ones = num % 10
        tens_part = mapping[tens] + "十" if tens > 1 else "十"
        ones_part = mapping[ones] if ones != 0 else ""
        return tens_part + ones_part

def compute_seating_pattern(seats_per_row):
    """
    根据“中间靠左优先，左右交替”规则计算排内座位填充顺序，
    返回一个列表，表示每个参会者应填入的座位自然索引（0为最左侧）。
    """
    pattern = []
    if seats_per_row % 2 == 0:
        best_index = seats_per_row // 2 - 1
        left = best_index
        right = best_index + 1
        toggle = True
        while left >= 0 or right < seats_per_row:
            if toggle and left >= 0:
                pattern.append(left)
                left -= 1
            elif not toggle and right < seats_per_row:
                pattern.append(right)
                right += 1
            toggle = not toggle
    else:
        best_index = seats_per_row // 2
        pattern.append(best_index)
        left = best_index - 1
        right = best_index + 1
        toggle = True
        while left >= 0 or right < seats_per_row:
            if toggle and left >= 0:
                pattern.append(left)
                left -= 1
            elif not toggle and right < seats_per_row:
                pattern.append(right)
                right += 1
            toggle = not toggle
    return pattern

def generate_column_labels(seats_per_row):
    """
    根据规则生成座位列标签：左侧为降序奇数、右侧为升序偶数
    如 seats_per_row = 10 时生成：["09", "07", "05", "03", "01", "02", "04", "06", "08", "10"]
    """
    if seats_per_row % 2 == 0:
        left_count = seats_per_row // 2
        right_count = seats_per_row // 2
    else:
        left_count = seats_per_row // 2 + 1
        right_count = seats_per_row // 2
    labels = []
    # 左侧：降序奇数
    for i in range(left_count):
        num = 2 * left_count - (2 * i + 1)
        labels.append(f"{num:02d}")
    # 右侧：升序偶数
    for i in range(right_count):
        num = 2 * (i + 1)
        labels.append(f"{num:02d}")
    return labels

def generate_seating_chart(file_path, seats_per_row):
    """
    读取 Excel 文件（人员名单包含 PERSONID 和 NAME 字段），
    按 PERSONID 升序排序后，按照每排 seats_per_row 个座位规则分配座位。
    返回 DataFrame（seating_df）和说明文字（explanation_text），其中说明文字
    按 PERSONID 升序排列，并且如果 NAME 为空或 NaN 则替换为“预留空位”。
    """
    # 读取 Excel 文件，并按 PERSONID 升序排序
    df = pd.read_excel(file_path)
    df_sorted = df.sort_values(by='PERSONID', ascending=True).reset_index(drop=True)
    
    total_people = len(df_sorted)
    num_rows = math.ceil(total_people / seats_per_row)
    seating_pattern = compute_seating_pattern(seats_per_row)
    col_labels = generate_column_labels(seats_per_row)
    
    seating_chart = []
    # assignments 用于记录每个参会者的座位分配信息，供生成说明文字使用
    assignments = []
    
    # 按顺序分配座位（这里使用原始顺序，即不进行翻转）
    for row in range(num_rows):
        row_seats = ["" for _ in range(seats_per_row)]
        start_idx = row * seats_per_row
        end_idx = min(start_idx + seats_per_row, total_people)
        persons = df_sorted.iloc[start_idx:end_idx]
        for i, (_, person) in enumerate(persons.iterrows()):
            if i < len(seating_pattern):
                seat_index = seating_pattern[i]
                row_seats[seat_index] = person['NAME']
                # 使用原始行号直接生成排号
                row_label = f"第{num_to_chinese(row + 1)}排"
                seat_label = col_labels[seat_index]
                assignments.append({
                    "PERSONID": person["PERSONID"],
                    "NAME": person["NAME"],
                    "ROW": row_label,
                    "SEAT": seat_label
                })
        seating_chart.append(row_seats)
    
    # 构造最终表格（翻转行顺序，使得第一排在表格最下方）
    seating_chart_final = seating_chart[::-1]
    row_labels = [f"第{num_to_chinese(i+1)}排" for i in range(num_rows)]
    row_labels_final = row_labels[::-1]
    
    final_table = []
    for label, row in zip(row_labels_final, seating_chart_final):
        final_table.append([label] + row)
    columns = ["排号"] + col_labels
    seating_df = pd.DataFrame(final_table, columns=columns)
    
    # 生成说明文字：按 PERSONID 升序排列，如果 NAME 为空或 NaN，则替换为“预留空位”
    assignments.sort(key=lambda x: x["PERSONID"])
    explanation_lines = []
    for item in assignments:
        name_val = item['NAME']
        if pd.isna(name_val) or str(name_val).strip().lower() == 'nan' or not str(name_val).strip():
            name_val = "预留空位"
        explanation_lines.append(f"{name_val} 在{item['ROW']}第{item['SEAT']}座位。")
    explanation_text = "\n".join(explanation_lines)
    
    return seating_df, explanation_text

def write_to_excel(seating_df, explanation_text):
    """
    使用 xlsxwriter 将 seating_df、主席台信息和说明文字写入一个 Excel 文件，
    将结果保存到 BytesIO 对象中返回。
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        seating_df.to_excel(writer, sheet_name="Sheet1", startrow=0, startcol=0, index=False)
        workbook  = writer.book
        worksheet = writer.sheets["Sheet1"]
        
        num_table_rows = seating_df.shape[0] + 1  # 包括表头
        num_table_cols = seating_df.shape[1]
        
        title_row = num_table_rows
        first_col = 0
        last_col = num_table_cols - 1
        title_text = "\u200B=========主席台========="
        merge_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 14,
            'bold': True,
            'num_format': '@'
        })
        worksheet.merge_range(title_row, first_col, title_row, last_col, title_text, merge_format)
        
        explanation_start_row = title_row + 2
        explanation_lines = explanation_text.split("\n")
        explanation_text_full = "\n".join(explanation_lines)
        explanation_format = workbook.add_format({
            'align': 'left',
            'valign': 'top',
            'text_wrap': True,
            'font_size': 10
        })
        explanation_end_row = explanation_start_row + len(explanation_lines) - 1
        worksheet.merge_range(explanation_start_row, first_col, explanation_end_row, last_col,
                              explanation_text_full, explanation_format)
        for r in range(explanation_start_row, explanation_end_row + 1):
            worksheet.set_row(r, 20)
    output.seek(0)
    return output

# ---------------- Streamlit 界面 ----------------

st.title("会议座位表生成器")

st.write("上传包含 PERSONID 和 NAME 字段的 Excel 文件（例如 meetperson.xlsx）")

uploaded_file = st.file_uploader("选择 Excel 文件", type=["xlsx"])

seats_per_row = st.number_input("每排座位数", min_value=1, max_value=30, value=10, step=1)

if uploaded_file is not None:
    if st.button("生成座位表"):
        try:
            seating_df, explanation_text = generate_seating_chart(uploaded_file, seats_per_row)
            excel_data = write_to_excel(seating_df, explanation_text)
            
            st.success("生成成功！")
            st.download_button(
                label="下载座位表 Excel 文件",
                data=excel_data,
                file_name="seatingmap.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.write("### 座位说明")
            st.text(explanation_text)
        except Exception as e:
            st.error(f"生成失败：{e}")
