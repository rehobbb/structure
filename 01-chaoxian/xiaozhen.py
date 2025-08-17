import re,csv,pandas as pd
def findchunks(indicators,endflag,allfile):
    chunks = []
    for indicator in indicators.values():
        i=0
        findTarget = False
        chunk = []
        while i<len(allfile):
            if indicator in allfile[i]:
                findTarget = True
                chunk.append(allfile[i])
                i += 1
                continue
            if findTarget:
                if endflag in allfile[i]:
                    break
                else :
                    chunk.append(allfile[i])
                    i += 1
                    continue
            i += 1
        chunks.append(chunk)
    return chunks

def is_contained(target,string_list):
    return any(target in s for s in string_list)

def get_disp_data(onechunk,index,data_a,indi_key,f_ratio):
    i = 0
    while i<len(onechunk):
        parts1 = onechunk[i].split()
        if parts1[0].isdigit() and int(parts1[0])<100:
            floor = (int(parts1[0])) 
            if is_contained(f_ratio,onechunk):
                parts2 = onechunk[i+1].split()
                value = max(float(parts1[index+2]),float(parts2[index]))
            else:
                 parts2 = re.split(r'\s{2,}',onechunk[i+1].strip())
                 if 'w' in indi_key:
                    value =parts2[index+1].replace('/ ','/')
                 else:
                    value =parts2[index].replace('/ ','/')
            exist_dict = next((item for item in data_a if item['floor'] == floor), None)
            if exist_dict:
                exist_dict[indi_key] = value
            else:
                dict = {'floor':floor,indi_key:value}
                data_a.append(dict)
            i += 1
            continue
        i += 1

direct1 = input('pleas input the directory:')
direct = direct1 + '\\wdisp.out'
d_index = 3
f_ratio = '规定'
with open(direct,'r') as file:
    allfile = file.readlines()
# 01. 找到对应字符串行范围
indicators = {
    'ex_disp':'X 方向地震作用下的楼层最大位移',
    'ey_disp':'Y 方向地震作用下的楼层最大位移' ,
    'wx_disp':'+X 方向风荷载作用下的楼层最大位移',
    'wy_disp':'+Y 方向风荷载作用下的楼层最大位移',
    'ex+_dratio':'X+ 偶然偏心规定水平力作用下的楼层最大位移',
    'ex-_dratio':'X- 偶然偏心规定水平力作用下的楼层最大位移',
    'ey+_dratio':'Y+ 偶然偏心规定水平力作用下的楼层最大位移',
    'ey-_dratio':'Y- 偶然偏心规定水平力作用下的楼层最大位移',
           }
endflag = '=='
datachunks = findchunks(indicators,endflag,allfile)
filterdatachunks = [ [s for s in tmpchunk if s.strip()]
    for tmpchunk in datachunks
]
# with open(direct1+'\\data.txt','w') as file2:
#     for chunk in filterdatachunks:
#         file2.writelines(chunk)
#         file2.write('\n')
# 02. 处理字符串
data_a = []
indi_keys = list(indicators.keys())
indi_values = list(indicators.values())
for onechunk,indi_key in zip(filterdatachunks ,indi_keys):
    get_disp_data(onechunk,d_index,data_a,indi_key,f_ratio)
with open(direct1+'\\data3.csv','w',newline='') as file3:
    headers = data_a[0].keys()
    writer = csv.DictWriter(file3,fieldnames=headers)
    writer.writeheader()
    writer.writerows(data_a)
# 04. 存入到表格
df = pd.DataFrame(data_a)
df.to_excel(direct1+'\\data3.xlsx',index=False)

