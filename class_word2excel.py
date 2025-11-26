import os
import re
import pandas as pd
from docx import Document

# 将教务处的Word文档课表，转换为可直接导入智慧校园系统的Excel文档

def extract_tables_from_docx(docx_path):
    """从Word文档提取表格并验证结构一致性（增强鲁棒性）"""
    doc = Document(docx_path)
    all_data = []
    header = None
    table_count = 0
    first_table_cols = None

    for table in doc.tables:
        table_count += 1
        table_data = []

        # 提取表格内容
        for row in table.rows:
            row_data = [cell.text.strip() for cell in row.cells]
            if all(cell == '' for cell in row_data):
                continue
            table_data.append(row_data)

        if not table_data:
            continue

        # 验证列数一致性（关键修复）
        if table_count == 1:
            first_table_cols = len(table_data[0])
            header = table_data[0]  # 保存首表表头
        else:
            if len(table_data[0]) != first_table_cols:
                raise ValueError(
                    f"表格{table_count}列数({len(table_data[0])})与首表({first_table_cols})不一致！"
                    "请确保所有表格表头格式相同"
                )

        # 处理表头行（仅首表使用表头）
        if table_count == 1:
            rows = table_data[1:]
        else:
            rows = table_data if table_data[0] != header else table_data[1:]

        all_data.extend(rows)

    if not all_data:
        raise ValueError("未提取到有效表格数据！请检查Word文档格式")

    # 增强：创建包含原始表头的DataFrame（保留原始表头用于后续映射）
    df = pd.DataFrame(all_data, columns=header)
    # 保留原始表头信息（用于鲁棒性处理）
    df.attrs['_original_header'] = header
    return df


def clean_and_standardize(df):
    """核心修复：日期时间逻辑重构 + 课程名称时间优先级处理 + 表头鲁棒性处理"""
    required_cols = ['日期', '时间', '内容', '主讲人', '职务职称', '上课地点']

    # === 修复：表头鲁棒性处理（关键增强）===
    # 1. 清洗表头列名（移除空格/多余字符）
    cleaned_header = [re.sub(r'\s+', '', str(col)) for col in df.columns]

    # 2. 定义表头映射规则（支持各种变体）
    header_mapping = {
        '内容': ['内容', '教学内容', '课程内容', '授课内容'],
        '主讲人': ['主讲人', '教师', '教 师', '教  师', '讲师', '授课教师'],
        '职务职称': ['职务职称', '职称', '职务', '身份'],
        '日期': ['日期', '日期时间', '日 期'],
        '时间': ['时间', '时段', '上课时间'],
        '上课地点': ['上课地点', '地点', '授课地点']
    }

    # 3. 映射到标准列名
    new_columns = []
    for col in df.columns:
        found = False
        for std_col, variants in header_mapping.items():
            # 检查清洗后的表头是否匹配
            if re.sub(r'\s+', '', str(col)) in [re.sub(r'\s+', '', v) for v in variants]:
                new_columns.append(std_col)
                found = True
                break
        if not found:
            new_columns.append(col)  # 保留原始列名（但后续会报错）

    # 重命名列
    df.columns = new_columns

    # === 修复：主讲人列值空格清理（关键增强）===
    if '主讲人' in df.columns:
        df['主讲人'] = df['主讲人'].apply(
            lambda x: str(x).replace(' ', '').replace('　', '').strip()
            if pd.notna(x) else x
        )
    # === 修复：上课地点列换行符处理 ===
    if '上课地点' in df.columns:
        df['上课地点'] = df['上课地点'].apply(
            lambda x: str(x).replace('\n', '').replace('\r', '').strip()
            if pd.notna(x) else x
        )
    # === 修复：检查必要列（使用标准列名）===
    for col in required_cols:
        if col not in df.columns:
            # 提供更友好的错误提示（包含可能的变体）
            possible_vars = []
            for std_col, variants in header_mapping.items():
                if std_col == col:
                    possible_vars = [v for v in variants if v != col]
                    break

            error_msg = f"缺少必要列: {col}（请检查Word表头，可能的变体: {', '.join(possible_vars)}）"
            raise ValueError(error_msg)

    # === 修复1：月份提取与补全（关键改进）===
    months = []
    current_month = None
    month_pattern = r'(\d+)月'

    # 遍历日期列提取月份
    for date_str in df['日期']:
        date_str = str(date_str).strip()
        if '月' in date_str:
            match = re.search(month_pattern, date_str)
            if match:
                current_month = match.group(1).zfill(2)  # 补零为两位
        months.append(current_month)

    # 验证是否找到月份（避免默认值）
    if all(m is None for m in months):
        raise ValueError("文档中未找到任何月份信息！请确保日期列包含'X月'格式（如'5月'）")

    # === 修复2：日期标准化（补零处理）===
    date_strings = []
    for idx, (date_str, month) in enumerate(zip(df['日期'], months)):
        date_str = str(date_str).strip()

        # 提取日（支持"25日"、"25"、"5"等格式）
        day_match = re.search(r'(\d+)(日|号)?$', date_str)
        day = day_match.group(1).zfill(2) if day_match else '01'

        # 组装完整日期（2025/05/25格式）
        date_strings.append(f"2025/{month}/{day}")

    # === 修复3：时间处理（双重优先级）===
    def get_time_range(time_desc):
        """优先级：课程名称时间 > 日期列时间描述"""
        # 先尝试从课程名称提取时间
        name_time = re.search(r'(\d{1,2}[:：]\d{2})[—\-～至](\d{1,2}[:：]\d{2})', str(df['内容'].iloc[idx]))
        if name_time:
            start, end = name_time.groups()
            return (start.replace('：', ':'), end.replace('：', ':'))

        # 再用日期列时间描述
        time_desc = str(time_desc).lower().strip()
        if '上午' in time_desc or '早' in time_desc:
            return ("9:00", "11:00")
        elif '下午' in time_desc or '午' in time_desc:
            return ("14:00", "16:00")
        elif '晚' in time_desc or '夜' in time_desc:
            return ("18:00", "20:00")
        return ("9:00", "11:00")  # 默认上午

    # 生成开始/结束时间
    start_times = []
    end_times = []
    for idx in df.index:
        start_time, end_time = get_time_range(df['时间'].iloc[idx])
        start_times.append(f"{date_strings[idx]} {start_time}:00")
        end_times.append(f"{date_strings[idx]} {end_time}:00")

    # === 修复4：课程名称清理（移除提取的时间）===
    course_names = []
    for idx, name in enumerate(df['内容']):
        cleaned = re.sub(r'\d{1,2}[:：]\d{2}[—\-～至]\d{1,2}[:：]\d{2}', '', str(name)).strip()
        course_names.append(cleaned if cleaned else "未命名课程")

    # 构建最终DataFrame
    df_clean = pd.DataFrame({
        '课程名称': course_names,
        '教学形式': '无',
        '授课教师': df['主讲人'].fillna('').str.strip(),
        '教师身份': df['职务职称'].apply(
            lambda x: '校内' if pd.notna(x) and x.strip() != '' else '校外'
        ),
        '开始时间': start_times,
        '结束时间': end_times,
        '上课地点': df['上课地点'].fillna('').str.strip(),
        '教职工是否听课': ''
    })

    # 修复：授课教师为空时清空教师身份
    df_clean.loc[df_clean['授课教师'] == '', '教师身份'] = ''

    return df_clean


def main(docx_file):
    output_excel = "课表导入.xlsx"
    output_excel = f"课表_{os.path.splitext(docx_file)[0]}.xlsx"

    try:
        df_raw = extract_tables_from_docx(docx_file)

        df_final = clean_and_standardize(df_raw)

        df_final.to_excel(output_excel, index=False, sheet_name="课表")
        print(f"✅ 转换完成: {output_excel} \n(Excel文件可直接导入智慧校园系统)\n")
    except Exception as e:
        print(f"❌ 处理失败: {str(e)}")


if __name__ == "__main__":
    # 自动查找当前目录下的.docx文件
    docx_files = [f for f in os.listdir('.') if f.endswith('.docx')]
    if not docx_files:
        raise FileNotFoundError(
            "当前目录无.docx文件！仅支持.docx格式的Word文件\n【可用Office或WPS打开文件后，另存为.docx格式文件】")
    else:
        for file in docx_files:
            print(f"🔍 正在处理: {file} (共{len(docx_files)}个文件)")
            main(file)
        print("✅ 所有文件处理完成！")
    input("可关闭此窗口，或按回车键退出...")