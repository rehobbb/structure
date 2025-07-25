import re

# 初始化存储0度和90度风振系数的列表
beta0_list = []
beta90_list = []

# 读取文件内容
with open('yjkwindforce.txt', 'r', encoding='gbk', errors='ignore') as file:
    content = file.read()

# 提取1-30层的数据
for layer in range(1, 31):
    # 匹配每层的章节内容
    layer_pattern = re.compile(rf'\[层{layer}\](.*?)�����F\(kN\):', re.DOTALL)
    layer_match = layer_pattern.search(content)
    
    if layer_match:
        layer_content = layer_match.group(1)
        
        # 提取0度方向的βz值
        beta0_pattern = re.compile(r'����0��:(.*?)βz\s+([\d.]+)', re.DOTALL)
        beta0_match = beta0_pattern.search(layer_content)
        if beta0_match:
            beta0 = float(beta0_match.group(2))
            beta0_list.append(beta0)
        
        # 提取90度方向的βz值
        beta90_pattern = re.compile(r'����90��:(.*?)βz\s+([\d.]+)', re.DOTALL)
        beta90_match = beta90_pattern.search(layer_content)
        if beta90_match:
            beta90 = float(beta90_match.group(2))
            beta90_list.append(beta90)

# 将结果写入文件
with open('wind_results.txt', 'w', encoding='utf-8') as f:
    f.write("第1~30层0度风振系数βz列表:\n")
    f.write(str(beta0_list) + "\n\n")
    f.write("第1~30层90度风振系数βz列表:\n")
    f.write(str(beta90_list) + "\n")

# 打印完成信息
print("结果已保存到 wind_results.txt 文件")