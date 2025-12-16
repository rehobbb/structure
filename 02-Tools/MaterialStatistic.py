import openpyxl
import re
import pandas as pd
import zipfile
import os
import tempfile
def repair_excel(file_path):
    # 尝试直接读取
    try:
        print("尝试直接读取Excel文件...")
        df = pd.read_excel(file_path, sheet_name='Sheet1', header=None)
        print("直接读取成功！")
        
    except Exception as e:
        print(f"直接读取失败: {str(e)[:100]}...")
        print("尝试修复文件后读取...")
        
        # 修复文件：删除损坏的样式表
        fixed_file = None
        try:
            # 创建临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 解压原始文件，排除样式表
                with zipfile.ZipFile(file_path, 'r') as zip_in:
                    for item in zip_in.infolist():
                        if item.filename != 'xl/styles.xml':  # 跳过损坏的样式表
                            zip_in.extract(item, temp_dir)
                
                # 生成修复后的临时文件
                fixed_file = os.path.join(temp_dir, 'fixed.xlsx')
                with zipfile.ZipFile(fixed_file, 'w', zipfile.ZIP_DEFLATED) as zip_out:
                    for root, dirs, files in os.walk(temp_dir):
                        for file in files:
                            if file != 'fixed.xlsx':  # 避免递归打包
                                file_path_in_temp = os.path.join(root, file)
                                arcname = os.path.relpath(file_path_in_temp, temp_dir)
                                zip_out.write(file_path_in_temp, arcname)
                
                # 读取修复后的文件
                df = pd.read_excel(fixed_file, sheet_name='Sheet1', header=2)
                df_1 = df[['楼层','构件类别','楼面面积(m2)','合计(kg)']].iloc[1:-2].copy().ffill()            
                print("文件修复并读取成功！")
                
        except Exception as e2:
            raise Exception(f"文件修复失败: {str(e2)}") from e
    return df_1
def read_text(path):
    try:
        with open(path, 'r') as file:
            return file.readlines()
    except Exception as e:
        print(f'文件读取失败:\n{str(e)}')
        # raise
def read_text2(path):
    try:
        with open(path, 'r',encoding='utf-16') as file:
            return file.readlines()
    except Exception as e:
        print(f'文件读取失败:\n{str(e)}')
        # raise
def find_begrund_num(lines):
    for line in lines:
        if '地下室层数' in line:
            num = int(line.split(':')[1].strip())
            break
    return num
def extract_concsteel(chunk,data):
        data.setdefault('楼层',[])
        data.setdefault('板-砼',[])
        data.setdefault('悬板-砼',[])
        data.setdefault('梁-砼',[])
        data.setdefault('柱-砼',[])
        data.setdefault('墙-砼',[])
        data.setdefault('梁-钢',[])
        data.setdefault('柱-钢',[])
        conc_flag,steel_flag = 0,0
        chunk = [ck for ck in chunk if ck.strip()]
        for list in chunk:
            floor_match = re.search(r'第\s*(\d+)自然层',list)
            if floor_match:
                floor = int(floor_match.group(1))
                data['楼层'].append(floor)
                continue
            if re.search('砼等级',list):
                conc_flag = 1
            if re.search('钢等级',list):
                steel_flag = 1
            if conc_flag and floor is not None:
                list_conc = list.split()
                if list_conc[0] == '楼板':
                    floor_conc = [float(l) for l in list_conc[2:] if l.replace('.','',1).isdigit()]
                    data['板-砼'].append(round(sum(floor_conc),1))
                if list_conc[0] == '悬挑板32':
                    floor_conc = [float(l) for l in list_conc[2:] if l.replace('.','',1).isdigit()]
                    data['悬板-砼'].append(round(sum(floor_conc),1))   
                if list_conc[0] == '梁':
                    beam_conc = [float(l) for l in list_conc[2:] if l.replace('.','',1).isdigit()]
                    data['梁-砼'].append(round(sum(beam_conc),1))      
                if list_conc[0] == '柱':
                    beam_conc = [float(l) for l in list_conc[2:] if l.replace('.','',1).isdigit()]
                    data['柱-砼'].append(round(sum(beam_conc),1))                           
                if list_conc[0] == '墙(总计)':
                    wall_conc = [float(l) for l in list_conc[2:] if l.replace('.','',1).isdigit()]
                    data['墙-砼'].append(round(sum(wall_conc),1))
                if '总面积' in list:
                    conc_flag = 0
            if steel_flag and floor is not None:
                list_steel = list.split()
                if list_steel[0] == '梁':
                    beam_steel = [float(l) for l in list_steel[2:] if l.replace('.','',1).isdigit()]
                    data['梁-钢'].append(round(sum(beam_steel),1))      
                if list_steel[0] == '柱':
                    col_steel = [float(l) for l in list_steel[2:] if l.replace('.','',1).isdigit()]
                    data['柱-钢'].append(round(sum(col_steel),1))     
                if '小计' in list:
                    steel_flag = 0                      
                    floor = None
def extract_rebar(df_rebar,data):   
    list=df_rebar['构件类别'].unique().tolist()
    list_ordi= ['板','梁','柱']
    list_wall = [i for i in list if i not in list_ordi]
    list_column = ['楼层','面积',*list_ordi]
    for l in list_column:
        data.setdefault(l,[])
    data['楼层'] = df_rebar['楼层'].unique()
    data['面积'] = df_rebar.groupby('楼层',sort=False)['楼面面积(m2)'].first().tolist()
    for l in list_ordi:
       if df_rebar[df_rebar['构件类别']==l]['合计(kg)'].sum() > 0:
        data[l] = (df_rebar[df_rebar['构件类别']==l]['合计(kg)']/1000).round(1).tolist()
    for l in list_wall:
       if df_rebar[df_rebar['构件类别']==l]['合计(kg)'].sum() > 0:
        data[l] = (df_rebar[df_rebar['构件类别']==l]['合计(kg)']/1000).round(1).tolist()
    max_len = max(len(v) for v in data.values())
    for key in data:
        data[key] += [0] * (max_len - len(data[key])) 
    df_tmp = pd.DataFrame(data)
    df_tmp['墙'] = (df_tmp[list_wall].sum(axis=1)/1000).round(1)
    df_tmp.drop(columns=list_wall,inplace=True)
    
def output(direct,df):
    str = direct.split('\\')[-1].split('-')[1]
    df.to_excel(direct + '/钢筋用量-统计-'+str+'.xlsx',index=False)
    print(df)
    print('统计完成')
def sum_df(df):
    sum_value = df.iloc[0,1:] + df.iloc[1,1:]
    df_sum = pd.DataFrame({
        '范围': ['总量'],  # 范围列赋值
        **sum_value.to_dict()  # 解包数值列的和（自动匹配列名）
    })
    df_total = pd.concat([df,df_sum],axis=0,ignore_index=True)
    return df_total
#主程序       
if __name__ == "__main__":
#定义初始变量，获取文件内容
    direct = input('Please input the YJK model directory:')
    direct = direct.rstrip('\\')
    wmass_path = direct + '/设计结果/wmass.out'  
    quant_path = direct + '/上部结构工程量.txt' 
    rebar_path = direct + '/施工图/钢筋用量.xlsx'
    wmass_lines = read_text(wmass_path)
    quant_lines = read_text2(quant_path)
    df_rebar = repair_excel(rebar_path)
    num_beground = find_begrund_num(wmass_lines)
    rebar_dict = {}
    concsteel_dict = {}
 #获取砼、型钢数据   
    extract_concsteel(quant_lines,concsteel_dict)
    extract_rebar(df_rebar,rebar_dict)
    stories = df_rebar['楼层'].unique()
    stories_beground = stories[:num_beground]
    stories_upground = stories[num_beground:]
    condition_beground = df_rebar['楼层'].isin(stories_beground)
    condition_upground = df_rebar['楼层'].isin(stories_upground)   


    # df_b = df_rebar[condition_beground]
    # df_u = df_rebar[condition_upground]

    # rebar_dict['范围'] = ['地下','地上']
    # rebar_dict['面积'] = []
    # rebar_dict['面积'].append(df_b.groupby('楼层')['楼面面积(m2)'].first().sum())
    # rebar_dict['面积'].append(df_u.groupby('楼层')['楼面面积(m2)'].first().sum())
    # series_b=df_b.groupby('构件类别')['合计(kg)'].sum()/1000
    # series_u=df_u.groupby('构件类别')['合计(kg)'].sum()/1000

    # for l in list :
    #     rebar_dict[l] = []
    #     rebar_dict[l].append(series_b[l] if l in series_b.index else 0)
    #     rebar_dict[l].append(series_u[l] if l in series_u.index else 0)

    # for l in list:
    #     rebar_list[l]=[]
    #     rebar_ele = df_b[(df_b['构件类别'] == l)]['合计(kg)'].sum()/1000
    #     rebar_list[l].append(rebar_ele)
    #     rebar_ele = df_u[(df_u['构件类别'] == l)]['合计(kg)'].sum()/1000
    #     rebar_list[l].append(rebar_ele)


    # df_mat_1 = pd.DataFrame(rebar_dict)
    # df_mat_1.iloc[:] = df_mat_1.iloc[:].round(1)
    # df_mat_1['墙'] = df_mat_1[list_wall].sum(axis=1)
    # df_mat_1['总'] = df_mat_1[list].sum(axis=1)
    # df_mat_1.drop(columns=list_wall,inplace=True)
    # list_colu = df_mat_1.columns[2:].tolist()
    # for l in list_colu:
    #     df_mat_1[f'{l}-单方'] = (df_mat_1[l]/df_mat_1['面积']*1000).round(1)
    # df_total = sum_df(df_mat_1)
    # output(direct,df_total)



