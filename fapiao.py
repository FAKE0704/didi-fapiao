import os
import PyPDF2
from tqdm import tqdm
from time import sleep
import pandas as pd
from pdfminer.high_level import extract_text
import re


def patch_read_filename(filePath):
    # 输入文件夹地址,格式: r"C:\Users\..."
    nameList = os.listdir(filePath)
    names= globals()
    count=0

    # 初始化一个空列表来存储文件名
    file_list = []

    # 遍历文件夹
    for filename in os.listdir(filePath):
        file_list.append(filename)

    # 输出文件名列表
    return file_list
    



def format_date(full_date):
    """格式化日期字符串
    将日期字符串格式从"2024-05-16 19:26至2024-08-13 21:33"转化为"2021-01-01至2021-01-02"
    """
    try:
        # 分割出两个日期部分
        parts = full_date.split("至")
        start_date = parts[0].split(" ")[0]
        end_date = parts[1].split(" ")[0]
        return f"{start_date}至{end_date}"
    except Exception as e:
        print(f"日期格式化失败: {e}")
        return full_date

def process_pdf(a):
    try:
        # 1. 取PDF文件名的第二个字符到符号"-"前的所有字符作为k1
        filename = os.path.basename(a)
        name_part = filename.split('.')[0]  # 去掉扩展名
        if '-' not in name_part:
            raise ValueError("PDF文件名中没有找到'-'字符")
        second_char_index = 1
        k1 = name_part[second_char_index : name_part.index('-')]
        print(f"k1:{k1}")

        # 2. 扫描PDF的文字，取"行程时间："后该行的所有字符作为date
        text = extract_text(a)
        lines = text.split('\n')
        
        # 定义列名和数据列表
        # columns = ["序号", "服务商", "车型", "上车时间", "城市", "起点", "终点", "金额"]
        columns = [ "服务商", "车型", "上车时间", "城市", "起点", "终点", "金额"]
        header_index = None
        data = []
        current_record = []
        collecting = False
        in_table = False

        date = None
        for line in lines:
            line = line.strip() # 去掉头尾的空格
            # 跳过空行和说明页码信息
            if not line or line.startswith("说明：") or line.startswith("页码："):
                continue
            
            # print(line) # 测试用
            if "行程时间：" in line:
                # print('找到了:行程时间：') # 测试用
                date = line.split("行程时间：")[1].strip()
                date = format_date(date)

            # 判断是否开始收集数据
            # print(line)
            # print(line == k1)
            if k1 in line:
                # print('开始收集数据')
                collecting = True

            if collecting:
                # 收集字段，直到字段数量等于列名数量
                if line in ["1", "2", "3","4","5","6","7","8","9","10","11","12"]:
                    # 序号行作为记录的开头
                    current_record.insert(0, line)
                else:
                    # print(line) # 测试用
                    # print("进入正则匹配")

                    # 使用正则表达式匹配车型和时间
                    # match = re.match(r'\s*(\w+\s+\w+) (\d{4}-\d{2}-\d{2} \d{2}:\d{2})$', line)
                    match = re.match(r'^\s*(\w+\s*\w+) (\d{4}-\d{2}-\d{2} \d{2}:\d{2})\s*$', line)
                    if match:
                        car_type = (match.group(1).strip().split())
                        if len(car_type) == 1:
                            current_record.append(car_type)
                        elif len(car_type) == 2:
                            current_record.append(car_type[0])
                            current_record.append(car_type[1])
                        else:
                            print(f"脚本未考虑该情况:{car_type}")
                        
                        departure_time = match.group(2)
                        # print(match.group(1).strip())
                        # print(f"车型:{car_type}, 时间:{departure_time}")
                        
                        current_record.append(departure_time)
                    else:
                        # print(f"正则表达式未匹配的行: {line}") # 测试用
                        current_record.append(line)
                

                # 检查是否完成一条记录
                
                # print(len(current_record))
                # # print(len(columns))
                # if (len(current_record) == 6) or (len(current_record) ==8):
                #     print(current_record)

                if len(current_record) >= len(columns):
                    # 处理金额字段，去除"元"字
                    current_record_tmp = current_record[:len(columns)]
                    current_record_tmp[-1] = current_record_tmp[-1].replace("元", "")
                    data.append(current_record_tmp)
                    # 删除前 n 个元素
                    del current_record[:len(columns)]
                    collecting = False

        # 3. k = k1 + date
        k = f"{k1}_{date}"
        
        

        # 4. 创建和输出DataFrame到directory
        # 5. 
        directory, _ = os.path.split(a)
        if data:
            df = pd.DataFrame(data, columns=columns)
            
            try:
                df.loc[:, "金额"] = df.loc[:, "金额"].astype(float)
            except ValueError as e:
                print(f"脚本错误：{e}")
                print("以下为相关信息：")
                print(df.loc[:, "金额"])
            df[["日期", "时间"]] = df["上车时间"].str.extract(r"^(\d{4}-\d{2}-\d{2})\s(\d{2}:\d{2})$")
            total_amount = df["金额"].sum()
            df.to_excel(directory+"/"+k+f"总金额{total_amount}.xlsx")
            print(f"已输出到路径: {directory}/{k}总金额{total_amount}.xlsx")
            
        else:
            print("未找到有效数据")
            return None
        
        print(f"k: {k}")
        print(f"k1: {k1}")
        print(f"date: {date}")

    except Exception as e:
        print(f"发生错误: {e}")
        


# 示例使用
# 假设你的pdf文件在"C:\Users\xxx"路径下

file_path = r"C:\Users\xxx" # 请用自己的地址替换
testlist =patch_read_filename(file_path)

for pdffile in testlist:
    if pdffile.endswith('电子行程单.pdf'):
        process_pdf(file_path+"/"+pdffile)