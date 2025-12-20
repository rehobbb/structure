import re
# 文件路径：YJK上部结构工程量文本
direct = 'D:/03-学习/07-PYTHON/YJK-V01/上部结构工程量.txt'
# 按自然层分割文本
def split_text(direct):
    with open(direct, 'r', encoding='utf-16') as f:
        text = f.read()
    # 编译正则：匹配“>第 数字 自然层:”，并保留分隔符
    pattern= re.compile(r'(>第\s*\d+自然层:.*?)(?=>第\s*\d+自然层:|>全楼统计:)',re.DOTALL)
    matches = pattern.findall(text)
    matches = [m.strip() for m in matches]
    return matches
# 调用函数获取结果
result = split_text(direct)
data_conc = {}
data_steel = {}
list_conc = ['楼层','楼板','悬挑板','梁','柱','墙(总计)']
list_steel = ['楼层','梁','柱']
for l_c in list_conc:
    data_conc.setdefault(l_c,[])
for l_s in list_steel:
    data_steel.setdefault(l_s,[])
for chunk in result:
    dict_conc = {}
    dict_steel = {}
    floor_match = re.search('>第\s*(\d+)自然层:',chunk)
    data_conc['楼层'].append(floor_match.group(1))
    data_steel['楼层'].append(floor_match.group(1))
    pattern_conc = r'\s*砼等级(.*?)(钢等级|$)'
    pattern_steel = r'\s*钢等级(.*?)$'
    match_conc = re.search(pattern_conc,chunk,re.DOTALL)
    match_steel = re.search(pattern_steel,chunk,re.DOTALL)
    if match_conc:
        chunk_conc_list = match_conc.group(1).split('\n')
    if match_steel:
        chunk_steel_list = match_steel.group(1).split('\n')
    for list in chunk_conc_list:
        for l_c in list_conc[1:]:
            if l_c in list:
                dict_conc[l_c] = list
    for l_c in list_conc[1:]:
        if l_c not in dict_conc:
            data_conc[l_c].append(0)
        else:
            list_c = dict_conc[l_c].split()
            conc = [float(l) for l in list_c[2:] if l.replace('.','',1).isdigit()]
            data_conc[l_c].append(round(sum(conc),0))
    for list in chunk_steel_list:
        for l_s in list_steel[1:]:
            if l_s in list:
                dict_steel[l_s] = list
    for l_s in list_steel[1:]:
        if l_s not in dict_steel:
            data_steel[l_s].append(0)
        else:
            list_s = dict_steel[l_s].split()
            steel = [float(l) for l in list_s[2:] if l.replace('.','',1).isdigit()]
            data_steel[l_s].append(round(sum(steel),1))
print(data_conc.keys())        
print(data_steel.keys())



