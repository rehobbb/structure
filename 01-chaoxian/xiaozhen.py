import re,pandas as pd
from pandas.core.api import NamedAgg
import xlsxwriter
import numpy as np
from xiaozhen_define import *

#*****************************
#结构参数
stories = '3-25'
limit_dfe = '1/500'
limit_dfw = '1/500'
limit_df = {
    'limit_dfe':fraction_to_float(limit_dfe),
    'limit_dfw':fraction_to_float(limit_dfw),
}
limit_id = {
    'limit_drd':1.2,
    'limit_dru':1.5,
    'limit_vmr':0.016,
    'limit_vc':0.80,
    'limit_stf':1.0,
    'limit_md':0.10,
    'limit_mu':0.50,
    'limit_v0d':0.10,
    'limit_v0u':0.20,
}
#*****************************
#路径及搜索参数
direct = input('please input the directory:')
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
    'ex_ds':'X 方向地震作用下的楼层最大位移',
    'ey_ds':'Y 方向地震作用下的楼层最大位移',
    'wx_ds':'+X 方向风荷载作用下的楼层最大位移',
    'wy_ds':'+Y 方向风荷载作用下的楼层最大位移',
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
#*****************************
#输出excel参数
head_df = {
    'ex_v':'层剪力-EX',
    'ey_v':'层剪力-EY',
    'ex_m':'层弯矩-EX',
    'ey_m':'层弯矩-EY',
    'wx_v':'层剪力-WX',
    'wy_v':'层剪力-WY',
    'wx_m':'层弯矩-WX',
    'wy_m':'层弯矩-WY',
    'ex_df':'位移角-EX',
    'ey_df':'位移角-EY',
    'ex_ds':'位移-EX',
    'ey_ds':'位移-EY',
    'wx_df':'位移角-WX',
    'wy_df':'位移角-WY',
    'wx_ds':'位移-WX',
    'wy_ds':'位移-WY',
}
head_id = {
    'ex+_dr':'位移比-X+',
    'ex-_dr':'位移比-X-',
    'ey+_dr':'位移比-Y+',
    'ey-_dr':'位移比-Y-',
    'ex_vmr':'剪重比-X',
    'ey_vmr':'剪重比-Y',
    'r_vcx':'抗剪承载力比-X',
    'r_vcy':'抗剪承载力比-Y',
    'r_sx':'侧向刚度比-X',
    'r_sy':'侧向刚度比-Y',
    'r_mx':'框架倾覆力矩比-X',
    'r_my':'框架倾覆力矩比-Y',
    'r_vx':'框架剪力比-X',
    'r_vy':'框架剪力比-Y',
}
plot_df = {
    'ex_v':['楼层剪力(kN)',''],
    'ey_v':['楼层剪力(kN)',''],
    'ex_m':['倾覆弯矩(kN.m)',''],
    'ey_m':['倾覆弯矩(kN.m)',''],
    'wx_v':['楼层剪力(kN)',''],
    'wy_v':['楼层剪力(kN)',''],
    'wx_m':['倾覆弯矩(kN.m)',''],
    'wy_m':['倾覆弯矩(kN.m)',''],
    'ex_df':['位移角','limit_dfe'],
    'ey_df':['位移角','limit_dfw'],
    'ex_ds':['位移(mm)',''],
    'ey_ds':['位移(mm)',''],
    'wx_df':['位移角','limit_dfw'],
    'wy_df':['位移角','limit_dfw'],
    'wx_ds':['位移(mm)',''],
    'wy_ds':['位移(mm)',''],
}
plot_id = {
    'ex+_dr':['扭转位移比','limit_dr'],
    'ex-_dr':['扭转位移比','limit_dr'],
    'ey+_dr':['扭转位移比','limit_dr'],
    'ey-_dr':['扭转位移比','limit_dr'],
    'ex_vmr':['剪重比','limit_vmr'],
    'ey_vmr':['剪重比','limit_vmr'],
    'r_vcx':['抗剪承载力比','limit_vc'],
    'r_vcy':['抗剪承载力比','limit_vc'],
    'r_sx':['侧向刚度比','limit_stf'],
    'r_sy':['侧向刚度比','limit_stf'],
    'r_mx':['框架倾覆力矩比','limit_m'],
    'r_my':['框架倾覆力矩比','limit_m'],
    'r_vx':['框架剪力比','limit_v0'],
    'r_vy':['框架剪力比','limit_v0'],
}

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
                if 'ds' in key:
                    parts2 = chunk[i].split()
                    value =round(float(parts2[3]),2)
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

# ********************* 
def process_df(df,df_head,id_head):
        df_df = pd.merge(df[['fl']],df[df_head],left_index=True,right_index=True)
        df_id = pd.merge(df[['fl']],df[id_head],left_index=True,right_index=True)
        for key,value in limit_df.items():
            df_df[key] = value
        for key,value in limit_id.items():
            df_id[key] = value
        return df_df,df_id
def output_df(df,writer,sheet_name,head):
    old_df = df.copy()
    columns_head = list(head.keys())
    columns = df.columns.tolist()
    for col in columns:
        if 'df' in col:
            df[col] = df[col].apply(fraction_to_float)
    df.to_excel(writer,index = False,sheet_name=sheet_name,startrow=2,header=False) 
    ws = writer.sheets[sheet_name]
    for col_idx,col_name in enumerate(columns): 
        if col_name == 'fl':
            ws.write(4+len(df),col_idx,'最大值')
            ws.write(5+len(df),col_idx,'所在楼层')
        if col_name in columns_head:
            ws.write(1,col_idx,'YJK')
            ws.write(4+len(df),col_idx,old_df[df[col_name]==df[col_name].max()][col_name].values[0])
            ws.write(5+len(df),col_idx,old_df[df[col_name] == df[col_name].max()]['fl'].values[0])
        elif 'limit' in col_name:
            ws.write(1,col_idx,col_name)
    ws.write(0,0,'楼层')
    start_cols = [i for i,col in enumerate(columns) if col in columns_head]
    for i,name in enumerate(columns_head):
        start_col = start_cols[i]
        end_col = start_cols[i+1]-1 if i<len(columns_head)-1 else len(columns_head)
        if end_col-start_col == 0 :
            ws.write(0,start_col,head[name])
            ws.write(3+len(df),start_col,head[name])
        else:
             ws.merge_range(0,start_col,0,end_col,head[name]) 
             ws.merge_range(3+len(df),start_col,3+len(df),end_col,head[name])      
# ********************* 
def plot_scatter_chart(df,wb,ws,head,plot,limit):
    columns_head = list(head.keys())
    format_data = wb.add_format(
        {
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
        }
    )
    ws.set_column(0,30,10,format_data)
    j,k = 0,0
    for  col in columns_head:
        chart = wb.add_chart(
            {
                'type':'scatter',
                'subtype':'smooth'
            }
        )
        chart.add_series(
            {
                'name':col,
                'categories':[ws.name,2,df.columns.get_loc(col),len(df)+1,df.columns.get_loc(col)],
                'values':[ws.name,2,0,len(df)+1,0],
                'line':{'color':'blue','width':2.25}
            }
        )
        if plot[col][1] != '':
            limit_col = plot[col][1]
            if limit_col in limit.keys():
                chart.add_series(
                    {
                        'name':'限值',
                        'categories':[ws.name,2,df.columns.get_loc(limit_col),len(df)+1,df.columns.get_loc(limit_col)],
                        'values':[ws.name,2,0,len(df)+1,0],
                        'line':{'color':'red','width':2.25,'dash_type':'dash'}
                    }
                )
            else :
                limit_col_d = limit_col + 'd'
                limit_col_u = limit_col + 'u'
                chart.add_series(
                    {
                        'name':'限值_d',
                        'categories':[ws.name,2,df.columns.get_loc(limit_col_d),len(df)+1,df.columns.get_loc(limit_col_d)],
                        'values':[ws.name,2,0,len(df)+1,0],
                        'line':{'color':'red','width':2.25,'dash_type':'dash'}
                    }
                )
                chart.add_series(
                    {
                        'name':'限值_u',
                        'categories':[ws.name,2,df.columns.get_loc(limit_col_u),len(df)+1,df.columns.get_loc(limit_col_u)],
                        'values':[ws.name,2,0,len(df)+1,0],
                        'line':{'color':'black','width':2.25,'dash_type':'dash'}
                    }
                )
        chart.set_title({
            'name':head[col],
            'name_font':{
                'name':'Arial',
                'size':12,
                'bold':True,
                'color':'black'
            },
        })
        chart.set_x_axis({
            'name':plot[col][0],
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
                'x':0.55,
                'y':0.4,
                'width':0.28,
                'height':0.1,
            },
        })
        chart.set_plotarea({
            'border':{'none':True},
            'layout':{
                'x':0.2,
                'y':0.12,
                'width':0.65,
                'height':0.75,
            },
        })
        if 'x' in col:
            chart_col = j * (chart.width// 63)
            ws.insert_chart(len(df)+7,chart_col,chart,{
                'x_offset':0,
                'y_offset':0,
            })
            j += 1
        else:           
            chart_col = k * (chart.width // 63)
            ws.insert_chart(len(df)+28,chart_col,chart,{
                'x_offset':0,
                'y_offset':0,
            })
            k += 1
#*****************************
data_a = []
indi_keys_disp = list(indicators_disp.keys())
indi_keys_force_e = list(indicators_force_e.keys())
with open(direct_wdisp,'r') as file_d:
    allfile_wdisp = file_d.readlines()
    datachunks_disp = find_chunks(indicators_disp,endflag_disp,allfile_wdisp)
    datachunks_disp = [ [s for s in tmpchunk if s.strip()]
    for tmpchunk in datachunks_disp]
    for onechunk,indi_key in zip(datachunks_disp ,indi_keys_disp):
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
    story = [int(i) for i in stories.split('-') ]
    story_range = range(story[0],story[1]+1)
    df = df[df['fl'].isin(story_range)]
    df['fl'] = df['fl']-story[0]+1   
    df_df,df_id =process_df(df,list(head_df.keys()),list(head_id.keys()))
    output_df(df_df,writer,'内力位移',head_df)
    output_df(df_id,writer,'整体指标',head_id) 
    wb = writer.book
    ws_df = writer.sheets['内力位移']
    ws_id = writer.sheets['整体指标']
    plot_scatter_chart(df_df,wb,ws_df,head_df,plot_df,limit_df)
    plot_scatter_chart(df_id,wb,ws_id,head_id,plot_id,limit_id)
    
 








