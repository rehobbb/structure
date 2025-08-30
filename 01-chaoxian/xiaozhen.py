import re,pandas as pd
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
#  ***************************       
def get_eforce_data(onechunk,e_index,data_a,indi_key):
     i = 0
     v_pattern = r'(\d+\.?\d*)\(\s*(\d+\.?\d*)%\)'
     while i<len(onechunk):
        parts = re.split(r'\s{2,}',onechunk[i].strip())
        if parts[0].isdigit():
            floor = int(parts[0])
            value_v = float(re.search(v_pattern,parts[e_index]).group(1))
            value_vmr = float(re.search(v_pattern,parts[e_index]).group(2))/100
            value_m = float(parts[e_index+1])
            exist_dict = next((item for item in data_a if item['floor'] == floor),None)
            if exist_dict:
                exist_dict[indi_key+'_v'] = round(value_v,0)
                exist_dict[indi_key+'_vmr'] = round(value_vmr,4)
                exist_dict[indi_key+'_m'] = round(value_m,0)
            else:
                dict = {'floor':floor,indi_key+'_v':round(value_v,0),
                        indi_key+'_vmr':round(value_vmr,4),indi_key+'_m':round(value_m,0)}
                data_a.append(dict)
            i += 1
            continue
        i += 1


def get_wforce_data(onechunk,w_index,data_a):
    i = 0
    while i <len(onechunk):
        parts = onechunk[i].split()
        if parts[0].isdigit():
            parts1 = onechunk[i+1].split()
            floor = int(parts[0])
            value_vx = float(parts[w_index])
            value_mx = float(parts[w_index+1])
            value_vy = float(parts1[w_index-2])
            value_my = float(parts1[w_index-1])
            exist_dict = next((item for item in data_a if item['floor'] == floor),None)
            if exist_dict:
                exist_dict['wx_v'] = round(value_vx,0)
                exist_dict['wx_m'] = round(value_mx,0)
                exist_dict['wy_v'] = round(value_vy,0)
                exist_dict['wy_m'] = round(value_my,0)
            else:
                dict = {'floor':floor,'wx_v':round(value_vx,0),'wx_m':round(value_mx,0),
                'wy_v':round(value_vy,0),'wy_m':round(value_my,0)}
                data_a.append(dict)
            i += 1
            continue
        i += 1

#  ***************************   
direct1 = input('pleas input the directory:')
direct_disp = direct1 + '\\wdisp.out'
direct_force_e = direct1 + '\\wzq.out'
direct_force_w = direct1 + '\\wmass.out'
d_index = 3
e_index = 3
w_index = 4
f_ratio = '规定'
with open(direct_disp,'r') as file_0:
    allfile_disp = file_0.readlines()
with open(direct_force_e,'r') as file_1:
    allfile_force_e = file_1.readlines()
with open(direct_force_w,'r') as file_2:
    allfile_force_w = file_2.readlines()
# 01. 找到对应字符串行范围
indicators_disp = {
    'ex_disp':'X 方向地震作用下的楼层最大位移',
    'ey_disp':'Y 方向地震作用下的楼层最大位移' ,
    'wx_disp':'+X 方向风荷载作用下的楼层最大位移',
    'wy_disp':'+Y 方向风荷载作用下的楼层最大位移',
    'ex+_dratio':'X+ 偶然偏心规定水平力作用下的楼层最大位移',
    'ex-_dratio':'X- 偶然偏心规定水平力作用下的楼层最大位移',
    'ey+_dratio':'Y+ 偶然偏心规定水平力作用下的楼层最大位移',
    'ey-_dratio':'Y- 偶然偏心规定水平力作用下的楼层最大位移',
           }
indicators_force_e = {
    'ex':'各层 X 方向的作用力(CQC)',
    'ey':'各层 Y 方向的作用力(CQC)',
}
indicators_force_w = {'w':'                           风荷载信息'}
endflag_disp = '=='
endflag_force_e = '='
endflag_force_w = '各楼层等效尺寸'
datachunks_disp = findchunks(indicators_disp,endflag_disp,allfile_disp)
datachunks_e_force = findchunks(indicators_force_e,endflag_force_e,allfile_force_e)
datachunks_w_force = findchunks(indicators_force_w,endflag_force_w,allfile_force_w) 
f_datachunks_disp = [ [s for s in tmpchunk if s.strip()]
    for tmpchunk in datachunks_disp
]
f_datachunks_force_e = [[s for s in tmpchunk if s.strip()]
    for tmpchunk in datachunks_e_force
]
f_datachunks_force_w = [[s for s in tmpchunk if s.strip()]
    for tmpchunk in datachunks_w_force
]


# 02. 处理字符串
data_a = []
indi_keys_disp = list(indicators_disp.keys())
indi_keys_force_e = list(indicators_force_e.keys())
for onechunk,indi_key in zip(f_datachunks_disp ,indi_keys_disp):
    get_disp_data(onechunk,d_index,data_a,indi_key,f_ratio)
for onechunk,indi_key in zip(f_datachunks_force_e ,indi_keys_force_e):
    get_eforce_data(onechunk,e_index,data_a,indi_key)
get_wforce_data(f_datachunks_force_w[0],w_index,data_a)
# with open(direct1+'\\data3.csv','w',newline='') as file3:
#     headers = data_a[0].keys()
#     writer = csv.DictWriter(file3,fieldnames=headers)
#     writer.writeheader()
#     writer.writerows(data_a)
# 04. 存入到表格
df = pd.DataFrame(data_a)
df.to_excel(direct1+'\\data3.xlsx',index=False)

