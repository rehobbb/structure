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
        df = pd.read_excel(file_path, sheet_name='Sheet1', header=2)
        df = df[['楼层','构件类别','楼面面积(m2)','合计(kg)']].iloc[1:-2].copy().ffill()            
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
                df = df[['楼层','构件类别','楼面面积(m2)','合计(kg)']].iloc[1:-2].copy().ffill()            
                print("文件修复并读取成功！")
                
        except Exception as e2:
            raise Exception(f"文件修复失败: {str(e2)}") from e
    return df
def read_wmass(path):
    try:
        with open(path, 'r') as file:
            return file.readlines()
    except Exception as e:
        print(f'文件读取失败:\n{str(e)}')
        # raise
def read_quant(path):
    try:
        with open(path, 'r', encoding='utf-16') as f:
            text = f.read()
        # 编译正则：匹配“>第 数字 自然层:”，并保留分隔符
        pattern= re.compile(r'(>第\s*\d+自然层:.*?)(?=>第\s*\d+自然层:|>全楼统计:)',re.DOTALL)
        matches = pattern.findall(text)
        matches = [m.strip() for m in matches]
        return matches
    except Exception as e:
        print(f'文件读取失败:\n{str(e)}')   
def find_begrund_num(lines):
    for line in lines:
        if '地下室层数' in line:
            num = int(line.split(':')[1].strip())
            break
    return num
def find_chunk(startflag,endflag,list):
        chunk = []
        i = 0
        find_target = False
        while i<len(list):
            if startflag in list[i]:
                find_target = True
                chunk.append(list[i])
                i += 1
                continue
            if find_target:
                if endflag in list[i]:
                    break
                else :
                    chunk.append(list[i])
                    i += 1
                    continue
            i += 1
        return chunk
def extract_conc(text,num_beground):
    data_conc = {}
    list_conc = ['楼层','面积','楼板','悬挑板','梁','柱','墙(总计)']
    for l_c in list_conc:
        data_conc.setdefault(l_c,[])
    for chunk in text:
        dict_conc = {}
        floor_match = re.search(r'>第\s*(\d+)自然层:',chunk)
        area_match = re.search(r'面积=\s*([\d.]+)',chunk)
        data_conc['楼层'].append(floor_match.group(1))
        data_conc['面积'].append(round(float(area_match.group(1)),0))
        pattern_conc = r'\s*砼等级(.*?)(钢等级|$)'
        match_conc = re.search(pattern_conc,chunk,re.DOTALL)
        if match_conc:
            chunk_conc_list = match_conc.group(1).split('\n')
        for list in chunk_conc_list:
            for l_c in list_conc[2:]:
                if l_c in list:
                    dict_conc[l_c] = list
        for l_c in list_conc[2:]:
            if l_c not in dict_conc:
                data_conc[l_c].append(0)
            else:
                list_c = dict_conc[l_c].split()
                conc = [float(l) for l in list_c[2:] if l.replace('.','',1).isdigit()]
                data_conc[l_c].append(round(sum(conc),0))
    df = pd.DataFrame(data_conc)
    df['楼板'] = df['楼板'] + df['悬挑板']
    df.rename(columns={'墙(总计)':'墙'},inplace=True)
    df.drop(columns='悬挑板')
    df['总']=df[df.columns[2:]].sum(axis=1)
    list_column = df.columns[2:]
    for l in list_column:
        df[l+'-单方'] = (df[l].div(df['面积'])*1000).round(2)
    
    return data_conc
def extract_steel(text,num_beground):
    data_steel = {}
    list_steel = ['楼层','面积','梁','柱','斜撑']
    for l_s in list_steel:
        data_steel.setdefault(l_s,[])
    for chunk in text:
        dict_steel = {}
        floor_match = re.search(r'>第\s*(\d+)自然层:',chunk)
        area_match = re.search(r'面积=\s*([\d.]+)',chunk)
        data_steel['楼层'].append(floor_match.group(1))
        data_steel['面积'].append(round(float(area_match.group(1)),0))
        pattern_steel = r'\s*钢等级(.*?)$'
        match_steel = re.search(pattern_steel,chunk,re.DOTALL)
        if match_steel:
            chunk_steel_list = match_steel.group(1).split('\n')
        for list in chunk_steel_list:
            for l_s in list_steel[2:]:
                if l_s in list:
                    dict_steel[l_s] = list
        for l_s in list_steel[2:]:
            if l_s not in dict_steel:
                data_steel[l_s].append(0)
            else:
                list_s = dict_steel[l_s].split()
                steel = [float(l) for l in list_s[2:] if l.replace('.','',1).isdigit()]
                data_steel[l_s].append(round(sum(steel),1))  
    return data_steel      
def extract_rebar(df,num_beground):   
    list=df['构件类别'].unique().tolist()
    list_ordi= ['板','梁','柱']
    list_nowall = [i for i in list_ordi if i in list]
    list_wall = [i for i in list if i not in list_nowall]
    list_data = list_nowall + ['墙']
    sr = df.groupby('楼层',sort=False)['楼面面积(m2)'].first()
    new_df = df.pivot_table(index='楼层',columns='构件类别',values='合计(kg)',\
             fill_value=0,aggfunc='first',sort=False).reset_index()
    df_sr = sr.to_frame(name='楼面面积(m2)').reset_index()
    new_df = df_sr.merge(new_df,on='楼层',how='left')
    new_df.rename(columns={'楼面面积(m2)':'面积'},inplace=True)
    new_df['面积'] = new_df['面积'].round(0)
    new_df['墙'] = new_df[list_wall].sum(axis=1)
    new_df[list_data]= new_df[list_data].div(1000).round(0)
    new_df['总'] = new_df[list_data].sum(axis=1)
    new_df.drop(columns=list_wall,inplace=True)
    list_column =[*list_data,'总']
    for l in list_column:
        new_df[l+'-单方'] = (new_df[l].div(new_df['面积'])*1000).round(1)
    stories = new_df['楼层']
    stories_beground = stories[:num_beground]
    stories_upground = stories[num_beground:]
    bool_beground = new_df['楼层'].isin(stories_beground)
    bool_upground = new_df['楼层'].isin(stories_upground)   
    new_df_b = new_df[bool_beground][['面积']+list_column]
    new_df_u = new_df[bool_upground][['面积']+list_column]
    sum_b = new_df_b.sum().to_frame().T
    sum_u = new_df_u.sum().to_frame().T
    sum_all = sum_b + sum_u
    cols = ['范围']+ new_df_b.columns.tolist()
    sum_b['范围'] = '地下'
    sum_u['范围'] = '地上'
    sum_all['范围'] = '合计'   
    df_sum = pd.concat([sum_b,sum_u,sum_all],ignore_index=True)[cols]
    for l in list_column:
        df_sum[l+'-单方'] = (df_sum[l].div(df_sum['面积'])*1000).round(1)
    return new_df,df_sum
def output(writer,df,sheetname):
    df.to_excel(writer,index=False,sheet_name=sheetname)

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
    wmass_lines = read_wmass(wmass_path)
    quant_text = read_quant(quant_path)
    df_rebar = repair_excel(rebar_path)
    num_beground = find_begrund_num(wmass_lines)
    conc_dict = {}
    steel_dict = {}
 #获取砼、型钢数据   
    conc_dict = extract_conc(quant_text,num_beground)
    steel_dict = extract_steel(quant_text,num_beground)
    df_rebar1,df_rebar_sum =extract_rebar(df_rebar,num_beground)
    str = direct.split('\\')[-1].split('-')[1]
    direct_output = direct + '/钢筋用量-统计-'+str+'.xlsx'
    with pd.ExcelWriter(direct_output,engine='xlsxwriter') as writer:
        output(writer,df_rebar_sum,'钢筋用量-统计')
        output(writer,df_rebar1,'钢筋用量')
    print(df_rebar_sum)
    print('统计完成')



