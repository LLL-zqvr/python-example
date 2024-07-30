import pandas as pd
import os
import openpyxl
def find_header_row(df, target_columns, max_rows=10):
    """
    在DataFrame的前max_rows行中查找包含所有target_columns的行。
    这通常用于Excel文件，其中列名可能不在第一行。
    如果找到，返回该行的索引（从0开始）；否则返回None。
    """
    for i in range(max_rows):
        # 将当前行的非空值转换为字符串集合，并与目标列名集合进行比较
        if set(df.iloc[i].dropna().astype(str)) >= set(target_columns):
            return i
    return None


#识别xlsx后缀的内容
def extract_data_from_file_xlsx(file_path, target_columns):
    """
    从单个Excel文件中提取特定列的数据。
    如果找到有效的列名行并成功提取数据，则返回DataFrame；否则返回None。
    如果表格标题不符合，则返回错误信息。

    参数:
    file_path (str): Excel文件的路径。
    target_columns (list of str): 需要提取的列名列表。

    返回:
    pandas.DataFrame 或 None: 如果找到有效的列名行并成功提取数据，则返回DataFrame；否则返回None。
    """

    try:
        # 尝试读取Excel文件的前几行以查找列名
        # 这里假设列名可能不在第一行，所以先不设置header参数
        df = pd.read_excel(file_path, nrows=min(10, len(target_columns) + 2), header=None)

        # 查找包含目标列名的行
        header_row = find_header_row(df, target_columns)

        if header_row is not None:
            # 如果找到有效的列名行，则重新读取整个文件，并设置正确的header
            df = pd.read_excel(file_path, header=header_row, engine= 'openpyxl')

            # 检查是否所有目标列都在DataFrame中
            if not set(target_columns).issubset(df.columns):
                raise ValueError(f"File {file_path} does not contain all required columns.")

            # 只保留包含目标列名的列
            df = df[target_columns]

            return df

        else:
            # 如果没有找到有效的列名行，则返回错误信息
            raise ValueError(f"文件 {file_path} 没有包含所有目标列的有效表头行。")

    except Exception as e:
        # 如果在读取或处理文件时发生错误，打印错误消息并返回None
        print(f"Error processing file {file_path}: {e}")
        return None

# 识别xsl后缀的内容
def extract_data_from_file_xls(file_path, target_columns):
    """
    从单个Excel文件中提取特定列的数据。
    如果找到有效的列名行并成功提取数据，则返回DataFrame；否则返回None。
    如果表格标题不符合，则返回错误信息。

    参数:
    file_path (str): Excel文件的路径。
    target_columns (list of str): 需要提取的列名列表。

    返回:
    pandas.DataFrame 或 None: 如果找到有效的列名行并成功提取数据，则返回DataFrame；否则返回None。
    """

    try:
        # 尝试读取Excel文件的前几行以查找列名
        # 这里假设列名可能不在第一行，所以先不设置header参数
        df = pd.read_excel(file_path, nrows=min(10, len(target_columns) + 2), header=None)

        # 查找包含目标列名的行
        header_row = find_header_row(df, target_columns)

        if header_row is not None:
            # 如果找到有效的列名行，则重新读取整个文件，并设置正确的header
            df = pd.read_excel(file_path, header=header_row, engine= 'xlrd')

            # 检查是否所有目标列都在DataFrame中
            if not set(target_columns).issubset(df.columns):
                raise ValueError(f"File {file_path} does not contain all required columns.")

            # 只保留包含目标列名的列
            df = df[target_columns]

            return df

        else:
            # 如果没有找到有效的列名行，则返回错误信息
            raise ValueError(f"文件 {file_path} 没有包含所有目标列的有效表头行。")

    except Exception as e:
        # 如果在读取或处理文件时发生错误，打印错误消息并返回None
        print(f"Error processing file {file_path}: {e}")
        return None

def check(folder_path):
    frozen = []
    for filename in os.listdir(folder_path):
        file = folder_path + "/" + filename
        excel = openpyxl.load_workbook(file)
        sheet = excel['Sheet1']
        if sheet.freeze_panes is not None:
            frozen.append(filename)
    if frozen is not None:
        for filename in frozen:
            print(filename,"中有冻结窗格!")

def process_excel_folder(folder_path, target_columns, output_file):
    """
    遍历指定文件夹中的所有Excel文件，提取数据，并将这些数据合并为一个单一的DataFrame，最后保存到新的Excel文件。

    参数:
    folder_path (str): 包含Excel文件的文件夹路径。
    target_columns (list of str): 需要从每个文件中提取的列名列表。
    output_file (str): 合并后的数据保存的文件名（包括路径）。
    """

    dataframes = []  # 创建一个空列表，用于存储从每个文件中提取的DataFrame

    # 遍历文件夹中的所有文件
    for filename in os.listdir(folder_path):
        # 检查文件是否为Excel文件
        if filename.endswith('.xlsx'):
            file_path = os.path.join(folder_path, filename)  # 构建完整的文件路径
            # 提取数据
            df = extract_data_from_file_xlsx(file_path, target_columns)
            if df is not None:  # 如果成功提取数据，则将其添加到列表中
                dataframes.append(df)
        elif filename.endswith('.xls'):
            file_path = os.path.join(folder_path, filename)
            df = extract_data_from_file_xls(file_path, target_columns)
            if df is not None:  # 如果成功提取数据，则将其添加到列表中
                dataframes.append(df)

    if dataframes:
        combined_df = pd.concat(dataframes, ignore_index=True)  # 合并DataFrame，忽略原始索引

        # 将合并后的DataFrame保存到新的Excel文件
        combined_df.to_excel(output_file, index=False)  # 不保存索引
        print(f"Data saved to {output_file}")  # 打印成功消息

    else:
        # 如果没有找到任何有效的数据，则打印错误消息
        print("No valid data found in any files.")


# 示例用法
target_columns = ['姓名', '身份证号码', '联系方式', '民族', '人员类型', '所在街道', '所在社区', '所属网格', '住址', '填报人', '填报时间', '修改类型']
output_file = r'F:\人口数据\重核（开江——总）.xlsx'
process_excel_folder(r'C:\Users\888888\Desktop\开江', target_columns, output_file)
check(r'C:\Users\888888\Desktop\test')