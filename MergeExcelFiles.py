import os
import pandas as pd
import datetime


# 提取BG name的辅助函数
def extract_string(input_string, delimiter, index):
    parts = input_string.split(delimiter)
    if index < len(parts):
        return parts[index]
    else:
        return None

# 使用方法说明：将4个BG的数据导出，并且按BG命名，放到脚本一个目录，然后用python运行这个脚本，即可实现自动将文件按照过滤条件自动合并成一个。

# 设置要合并的文件扩展名和过滤条件
file_extension = '.xlsx'  # 替换为实际的文件扩展名
filter_field = '创建时间'  # 替换为要过滤的字段名
filter_values = ('2023-10', '2023-09', '2023-08')  # 替换为要过滤的特殊文本
#filter_values = ('2023-10',)  # 替换为要过滤的特殊文本

# 获取当前时间
current_time = datetime.datetime.now()
output_files = './tmp/OR_' + current_time.strftime("%Y%m%d_%H%M%S") + '.xlsx'

# 获取当前目录下所有扩展名为xlsx的文件
file_names = [file for file in os.listdir('./') if (file.endswith(file_extension))]

# 创建零时目录：
if not os.path.exists('./tmp'):
    os.mkdir('./tmp')

# 创建一个空的数据框来存储合并的数据
merged_data = pd.DataFrame()

try:
    # 循环读取和合并每个表格
    for file_name in file_names:
        print("Start merge files: " + str(file_name))

        # 读取表格数据
        df = pd.read_excel(os.getcwd() + ".\\" + file_name)

        bg_name = extract_string(file_name, ".", 0)

        for item in filter_values:
            #print('filter ++'+item+'++')
            # 过滤数据
            df_filtered = df[df[filter_field].str.contains(item)]

            # print("BG Name: "+bg_name)
            df_filtered.insert(0, 'BG', bg_name)

            # 将过滤后的数据合并到整体数据框
            merged_data = pd.concat([merged_data, df_filtered], ignore_index=True)

    print("Save result to : " + str(output_files))
    # 将合并后的数据保存为一个新的Excel表格
    merged_data.to_excel(output_files, index=False)

except Exception as e:
    print("文件处理失败:", str(e))
