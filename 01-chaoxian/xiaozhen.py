import re,pandas as pd
import xlsxwriter

#  ***************************    
def find_chunk(indicator,endflag,list):
    chunk = []
    i = 0 
    find_target = False
    end_count = 0
    while i<len(list):
        if indicator in list[i]:
            find_target = True
            chunk.append(list[i])
            i += 1
            continue
        if find_target:
            if endflag == '**':
                if endflag in list[i]:
                    end_count += 1
                if end_count == 2:
                    break
                else :
                    chunk.append(list[i])
                    i += 1
                    continue
            else:
                if endflag in list[i]:
                    break
                else :
                    chunk.append(list[i])
                    i += 1
                    continue
        i += 1
    return chunk
#  ***************************    
def find_chunks(indicators,endflag,file_list):
    chunks = []
    for indicator in indicators.values():
        chunk = find_chunk(indicator,endflag,file_list)
        chunks.append(chunk)
    return chunks
#  ***************************    
def is_contained(target,string_list):
    return any(target in s for s in string_list)
#  ***************************    
def get_disp_data(chunk,index,data,key,ratio):
    i = 0
    while i<len(chunk):
        parts1 = chunk[i].split()
        if parts1[0].isdigit() and int(parts1[0])<100:
            floor = (int(parts1[0])) 
            if is_contained(ratio,chunk):
                parts2 = chunk[i+1].split()
                value = max(float(parts1[index+2]),float(parts2[index]))
            else:
                 parts2 = re.split(r'\s{2,}',chunk[i+1].strip())
                 if 'w' in key:
                    value =parts2[index+1].replace('/ ','/')
                 else:
                    value =parts2[index].replace('/ ','/')
            exist_dict = next((item for item in data if item['fl'] == floor), None)
            if exist_dict:
                exist_dict[key] = value
            else:
                dict = {'fl':floor,key:value}
                data.append(dict)
            i += 1
            continue
        i += 1
#  ***************************       
def get_eforce_data(chunk,index,data,key):
     i = 0
     v_pattern = r'(\d+\.?\d*)\(\s*(\d+\.?\d*)%\)'
     while i<len(chunk):
        parts = re.split(r'\s{2,}',chunk[i].strip())
        if parts[0].isdigit():
            floor = int(parts[0])
            value_v = float(re.search(v_pattern,parts[index]).group(1))
            value_vmr = float(re.search(v_pattern,parts[index]).group(2))/100
            value_m = float(parts[index+1])
            exist_dict = next((item for item in data if item['fl'] == floor),None)
            if exist_dict:
                exist_dict[key+'_v'] = round(value_v,0)
                exist_dict[key+'_vmr'] = round(value_vmr,4)
                exist_dict[key+'_m'] = round(value_m,0)
            else:
                dict = {'fl':floor,key+'_v':round(value_v,0),
                        key+'_vmr':round(value_vmr,4),key+'_m':round(value_m,0)}
                data.append(dict)
            i += 1
            continue
        i += 1
#  ***************************    
def get_wforce_data(chunk,index,data):
    i = 0
    while i <len(chunk):
        parts = chunk[i].split()
        if parts[0].isdigit():
            parts1 = chunk[i+1].split()
            floor = int(parts[0])
            value_vx = float(parts[index])
            value_mx = float(parts[index+1])
            value_vy = float(parts1[index-2])
            value_my = float(parts1[index-1])
            exist_dict = next((item for item in data if item['fl'] == floor),None)
            if exist_dict:
                exist_dict['wx_v'] = round(value_vx,0)
                exist_dict['wx_m'] = round(value_mx,0)
                exist_dict['wy_v'] = round(value_vy,0)
                exist_dict['wy_m'] = round(value_my,0)
            else:
                dict = {'fl':floor,'wx_v':round(value_vx,0),'wx_m':round(value_mx,0),
                'wy_v':round(value_vy,0),'wy_m':round(value_my,0)}
                data.append(dict)
            i += 1
            continue
        i += 1
#  ***************************  
def get_ratiovc_data(chunk,index,data):
    i = 0
    while i <len(chunk):
        parts = chunk[i].split()
        if parts[0].isdigit():
            floor = int(parts[0])
            ratio_vx = float(parts[index])
            ratio_vy = float(parts[index+1])
            exist_dict = next((item for item in data if item['fl'] == floor),None)
            if exist_dict:
                exist_dict['r_vcx'] = round(ratio_vx,2)
                exist_dict['r_vcy'] = round(ratio_vy,2)
            else:
                dict = {'fl':floor,'r_vcx':round(ratio_vx,2),'r_vcy':round(ratio_vy,2)}
                data.append(dict)
            i += 1
            continue
        i += 1
# ********************* 
def get_ratios_data(chunk,stru,data):
    if '剪' in stru:
        rs_x = 'Ratx2'
        rs_y = 'Raty2'
    else:
        rs_x = 'Ratx1'
        rs_y = 'Raty2'
    floor = None
    for list in chunk:
        floor_match = re.search(r'Floor No\.\s*(\d+)',list)
        if floor_match:
            floor = int(floor_match.group(1))
            continue
        rs_match = re.search(re.escape(rs_x)+r'=\s*([\d.]+)\s*' + re.escape(rs_y) + r'=\s*([\d.]+)',list)
        if rs_match and floor is not None:
            ratio_x = float(rs_match.group(1))
            ratio_y = float(rs_match.group(2))
            exist_dict = next((item for item in data if item['fl'] == floor),None)
            if exist_dict:
                exist_dict['r_sx'] = round(ratio_x,2)
                exist_dict['r_sy'] = round(ratio_y,2)
            else:  
                dict = {'fl':int(floor),'r_sx':round(ratio_x,2),'r_sy':round(ratio_y,2)}
                data.append(dict)
#***********************OPTION 2***************************
# def get_ratios_data(onechunk,s_structure,data_a):
# if '剪' in s_structure:
#     rs_x = 'Ratx2'
#     rs_y = 'Raty2'
# else:
#     rs_x = 'Ratx1'
#     rs_y = 'Raty2'  
#     i = 0
# floor_chunks = []
# pattern = r'\d+\.\d+'
# while i < len(onechunk):
#     if 'fl' in onechunk[i]:
#         result = re.search(r'\d+',onechunk[i])
#         if result:
#             floor = int(result.group())
#             floor_chunk = find_chunk('fl','--',onechunk[i:])
#             floor_chunks.append(floor_chunk)
#         i += 1
#         continue
#     i += 1
# for floor_chunk in floor_chunks:  
#     j = 0
#     floor = re.search(r'\d+',floor_chunk[0]).group()
#     while j <len(floor_chunk):           
#         if rs_x in floor_chunk[j]:
#             ratio_x,ratio_y = map(float,re.findall(pattern,floor_chunk[j]))
#             exist_dict = next((item for item in data_a if item['fl'] == int(floor)),None)
#             if exist_dict:
#                 exist_dict['ratio_sx'] = round(ratio_x,2)
#                 exist_dict['ratio_sy'] = round(ratio_y,2)
#             else:  
#                 dict = {'fl':int(floor),'ratio_sx':round(ratio_x,2),'ratio_y':round(ratio_sy,2)}
#                 data_a.append(dict)
#             j += 1
#             continue
#         j += 1
#  ***************************    
def get_ratiom_data(chunk,index,data):
    i = 0 
    while i <len(chunk):
        parts = chunk[i].split()
        if parts[0].isdigit() :
            floor = int(parts[0])
            if parts[index-1] == 'X':
                value_mx = float(parts[index].strip('%'))/100
                exist_dict = next((item for item in data if item['fl'] == floor),None)
                if exist_dict:
                    exist_dict['r_mx'] = round(value_mx,2)
                else:
                     dict = {'fl':floor,'r_mx':round(value_mx,2)}
                     data.append(dict)
                i += 1
                continue
            else :
                value_my = float(parts[index].strip('%'))/100
                exist_dict = next((item for item in data if item['fl'] == floor),None)
                if exist_dict:
                    exist_dict['r_my'] = round(value_my,2)
                else:
                     dict = {'fl':floor,'r_my':round(value_my,2)}
                     data.append(dict)
                i += 1
                continue
        i += 1
# ********************* 
def get_ratiov0_data(chunk,index,data):
    i = 0 
    while i <len(chunk):
        parts = chunk[i].split()
        if parts[0].isdigit() :
            floor = int(parts[0])
            if parts[index-5] == 'X':
                value_vx = float(parts[index].strip('%'))/100
                exist_dict = next((item for item in data if item['fl'] == floor),None)
                if exist_dict:
                    exist_dict['r_vx'] = round(value_vx,2)
                else:
                     dict = {'fl':floor,'r_vx':round(value_vx,2)}
                     data.append(dict)
                i += 1
                continue
            else :
                value_vy = float(parts[index].strip('%'))/100
                exist_dict = next((item for item in data if item['fl'] == floor),None)
                if exist_dict:
                    exist_dict['r_vy'] = round(value_vy,2)
                else:
                     dict = {'fl':floor,'r_vy':round(value_vy,2)}
                     data.append(dict)
                i += 1
                continue
        i += 1
def fraction_to_float(fraction):
    if isinstance(fraction,str) and '/' in fraction:
        numerator,denominator = fraction.split('/')
        try:
            return float(numerator) / float(denominator)
        except ZeroDivisionError:
            return 0.0
    return float(fraction)
# ********************* 
direct = input('pleas input the directory:')
direct_wdisp = direct + '\\wdisp.out'
direct_wzq = direct + '\\wzq.out'
direct_wmass = direct + '\\wmass.out'
direct_wv02q = direct + '\\wv02q.out'
s_structure = ''
d_index = 3
e_index = 3
w_index = 4
ratiov_index = 4
m_index = 3
v0_index = 7
f_ratio = '规定'
indicators_disp = {
    'ex_df':'X 方向地震作用下的楼层最大位移',
    'ey_df':'Y 方向地震作用下的楼层最大位移' ,
    'wx_df':'+X 方向风荷载作用下的楼层最大位移',
    'wy_df':'+Y 方向风荷载作用下的楼层最大位移',
    'ex+_dr':'X+ 偶然偏心规定水平力作用下的楼层最大位移',
    'ex-_dr':'X- 偶然偏心规定水平力作用下的楼层最大位移',
    'ey+_dr':'Y+ 偶然偏心规定水平力作用下的楼层最大位移',
    'ey-_dr':'Y- 偶然偏心规定水平力作用下的楼层最大位移',
           }
indicators_force_e = {
    'ex':'各层 X 方向的作用力(CQC)',
    'ey':'各层 Y 方向的作用力(CQC)',
}
indicator_force_w = '                           风荷载信息'
indicator_ratio_vc = '楼层抗剪承载力验算'
indicator_ratio_s = '各层刚心、偏心率、相邻层侧移刚度比等计算信息'
indicator_ratio_m = '规定水平力下框架柱、短肢墙地震倾覆力矩百分比'
indicator_ratio_v0 = '框架柱地震剪力及百分比'
endflag_disp = '=='
endflag_force_e = '='
endflag_force_w = '各楼层等效尺寸'
endflag_ratio_v = '**'
endflag_ratio_s = '**'
endflag_ratio_m = '**'
endflag_ratio_v0 = '**'
data_a = []
indi_keys_disp = list(indicators_disp.keys())
indi_keys_force_e = list(indicators_force_e.keys())
with open(direct_wdisp,'r') as file_d:
    allfile_wdisp = file_d.readlines()
    datachunks_disp = find_chunks(indicators_disp,endflag_disp,allfile_wdisp)
    f_datachunks_disp = [ [s for s in tmpchunk if s.strip()]
    for tmpchunk in datachunks_disp]
    for onechunk,indi_key in zip(f_datachunks_disp ,indi_keys_disp):
        get_disp_data(onechunk,d_index,data_a,indi_key,f_ratio)
with open(direct_wzq,'r') as file_e:
    allfile_wzq = file_e.readlines()
    datachunks_e_force = find_chunks(indicators_force_e,endflag_force_e,allfile_wzq)
    f_datachunks_e_force = [[s for s in tmpchunk if s.strip()]
    for tmpchunk in datachunks_e_force]
    for onechunk,indi_key in zip(f_datachunks_e_force ,indi_keys_force_e):
         get_eforce_data(onechunk,e_index,data_a,indi_key)
with open(direct_wmass,'r') as file_m:
    allfile_wmass = file_m.readlines()
    for line in allfile_wmass:
        if '结构总体信息' in line:
            s_tmp = allfile_wmass[allfile_wmass.index(line)+1]
            s_structure = re.search(r'\s*\w+:\s*(\w+)',s_tmp).group(1)
            break
    datachunk_w_force = find_chunk(indicator_force_w,endflag_force_w,allfile_wmass) 
    f_datachunk_w_force = [s for s in datachunk_w_force if s.strip()]
    get_wforce_data(f_datachunk_w_force,w_index,data_a)
    datachunk_ratio_v = find_chunk(indicator_ratio_vc,endflag_ratio_v,allfile_wmass) 
    f_datachunk_ratio_v = [s for s in datachunk_ratio_v if s.strip()]
    get_ratiovc_data(f_datachunk_ratio_v,ratiov_index,data_a)
    datachunk_ratio_s = find_chunk(indicator_ratio_s,endflag_ratio_s,allfile_wmass) 
    f_datachunk_ratio_s = [s for s in datachunk_ratio_s if s.strip()]
    get_ratios_data(f_datachunk_ratio_s,s_structure,data_a)
with open(direct_wv02q,'r') as file_v:
    allfile_wv02q = file_v.readlines()
    datachunk_ratio_m = find_chunk(indicator_ratio_m,endflag_ratio_m,allfile_wv02q)
    f_datachunk_ratio_m = [s for s in datachunk_ratio_m if s.strip()]
    get_ratiom_data(f_datachunk_ratio_m,m_index,data_a)
    datachunk_ratio_v0 = find_chunk(indicator_ratio_v0,endflag_ratio_v0,allfile_wv02q)
    f_datachunk_ratio_v0 = [s for s in datachunk_ratio_v0 if s.strip()]
    get_ratiov0_data(f_datachunk_ratio_v0,v0_index,data_a)

with pd.ExcelWriter(direct + '\\data3.xlsx',engine = 'xlsxwriter') as writer:
    df = pd.DataFrame(data_a)
    head = df.columns.tolist()
    for col in head:
        if 'df' in col:
            df[col] = df[col].apply(fraction_to_float)
    df.to_excel(writer,index = False,sheet_name='YJK') 
    workbook = writer.book
    worksheet = writer.sheets['YJK']
    for i , col in enumerate(head[1:]):
        chart = workbook.add_chart(
            {
                'type':'scatter',
                'subtype':'smooth'
            }
        )
        chart.add_series(
            {
                'name':col,
                'categories':['YJK',1,df.columns.get_loc(col),len(df),df.columns.get_loc(col)],
                'values':['YJK',1,0,len(df),0],
                'line':{'color':'blue','width':2.25}
            }
        )
        chart.set_title({
            'name':col,
            'name_font':{
                'name':'Arial',
                'size':12,
                'bold':True,
                'color':'black'
            },
        })
        chart.set_x_axis({
            'name':col,
            'num_font':{'size':10},
            'name_font':{'size':10},
            'major_gridlines':{
                'visible':True,
                'line':{'color':'gray','width':0.5,'dash_type':'long_dash'}
            },
        })
        chart.set_y_axis({
            'name':'楼层',
            'num_font':{'size':10},
            'name_font':{'size':10},
            'major_gridlines':{
                'visible':True,
                'line':{'color':'gray','width':0.5,'dash_type':'long_dash'}
            },
        })
        cm_to_px_width = 36.5
        cm_to_px_height = 41
        chart.set_size({
            'width':int(7 * cm_to_px_width),
            'height':int(9.5* cm_to_px_height)
        })
        chart.set_legend({
            'position':'right',
            'font':{
                'name':'Arial',
                'size':10,
                'bold':False,
                'color':'black',
            },
            'border':{
                'color':'black',
                'width':0.7,
                'dash_type':'solid',
            },
            'layout':{
                'x':0.6,
                'y':0.4,
                'width':0.25,
                'height':0.1,
            },
        })
        chart.set_plotarea({
            'border':{'none':True},
            'layout':{
                'x':0.2,
                'y':0.15,
                'width':0.65,
                'height':0.7,
            },
        })
        chart_col = i * (chart.width // 60)
        worksheet.insert_chart(len(df)+4,chart_col,chart,{
            'x_offset':i*5,
            'y_offset':0,
        })








