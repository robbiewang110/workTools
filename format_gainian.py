import re

def process_strings():
    try:
        # 读取输入文件
        with open('gainian.txt', 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 使用正则表达式分割字符串，支持中文和英文分号
        # 正则表达式模式：中文分号"；"或英文分号";"
        strings = re.split(r'[;；]', content)
        
        # 去除每个字符串两端的空白字符，并过滤空字符串
        cleaned_strings = [s.strip() for s in strings if s.strip()]
        
        # 去重并保持原始顺序
        seen = set()
        unique_strings = []
        for s in cleaned_strings:
            if s not in seen:
                seen.add(s)
                unique_strings.append(s)
        
        # 写入输出文件，使用Linux换行方式(\n)
        with open('output.txt', 'w', encoding='utf-8', newline='\n') as f:
            for s in unique_strings:
                f.write(s + '\n')
        
        print(f"处理完成！共找到 {len(strings)} 个字符串片段，去重后剩下 {len(unique_strings)} 个。")
        print(f"结果已保存到 output.txt")
        
    except FileNotFoundError:
        print("错误：找不到 gainian.txt 文件")
    except Exception as e:
        print(f"处理过程中发生错误：{e}")

if __name__ == "__main__":
    process_strings()