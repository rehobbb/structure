import re
import pandas as pd
import zipfile
import os
import tempfile
from plot_scatter import plot_scatter as ps
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
#读取总信息文件，返回按层文本列表
    try:
        with open(path, 'r') as file:
            return file.readlines()
    except Exception as e:
        print(f'文件读取失败:\n{str(e)}')
def read_quant(path):
#读取材料用量文件，返回按层文本列表
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
def process_df(df,num_beground,scale):
#处理明细df，获取按楼层汇总数据
    decimal = 2
    new_df = df.copy()
    columns_data0 = new_df.columns[2:]
    new_df['总'] = new_df[columns_data0].sum(axis=1)
    columns_data1 = new_df.columns[2:]
    columns_data2 = ['面积',*columns_data1]
    for c in columns_data1:
        new_df[c+'单'] = (new_df[c].div(new_df['面积'])*scale).round(decimal)
    stories = new_df['楼层']
    stories_beground = stories[:num_beground]
    stories_upground = stories[num_beground:]
    bool_beground = new_df['楼层'].isin(stories_beground)
    bool_upground = new_df['楼层'].isin(stories_upground)
    df_b = new_df[bool_beground][columns_data2]
    df_u = new_df[bool_upground][columns_data2]
    sum_b = df_b.sum().to_frame().T
    sum_u = df_u.sum().to_frame().T
    sum_all = sum_b + sum_u
    cols = ['范围'] + columns_data2
    sum_b['范围'] = '地下'
    sum_u['范围'] = '地上'
    sum_all['范围'] = '全楼'
#获取地上、地下、全楼汇总数据
    df_sum = pd.concat([sum_b,sum_u,sum_all],ignore_index=True)[cols].round(0)
    for l in columns_data1:
        df_sum[l+'单'] = (df_sum[l].div(df_sum['面积'])*scale).round(decimal)
    return new_df,df_sum
def extract_conc(text):
#根据型钢明细text，获取按楼层汇总数据
    data_conc = {}
    list_conc = ['楼层','面积','楼板','悬挑板','梁','柱','墙(总计)']
    for l_c in list_conc:
        data_conc.setdefault(l_c,[])
    for chunk in text:
        dict_conc = {}
        floor_match = re.search(r'>第\s*(\d+)自然层:',chunk)
        area_match = re.search(r'面积=\s*([\d.]+)',chunk)
        data_conc['楼层'].append(int(floor_match.group(1)))
        data_conc['面积'].append(round(float(area_match.group(1)),2))
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
                data_conc[l_c].append(round(sum(conc),2))
    df = pd.DataFrame(data_conc)
    df['楼板'] = df['楼板'] + df['悬挑板']
    df.drop(columns='悬挑板',inplace=True)
    df.rename(columns={'墙(总计)':'墙'},inplace=True)
    return df
def extract_steel(text):
#根据型钢明细text，获取按楼层汇总数据
    data_steel = {}
    list_steel = ['楼层','面积','梁','柱','斜撑']
    for l_s in list_steel:
        data_steel.setdefault(l_s,[])
    for chunk in text:
        dict_steel = {}
        floor_match = re.search(r'>第\s*(\d+)自然层:',chunk)
        area_match = re.search(r'面积=\s*([\d.]+)',chunk)
        data_steel['楼层'].append(int(floor_match.group(1)))
        data_steel['面积'].append(round(float(area_match.group(1)),2))
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
                    data_steel[l_s].append(round(sum(steel),2))  
        else:
            for l_s in list_steel[2:]:
                data_steel[l_s].append(0)
    df = pd.DataFrame(data_steel)
    for l in df[2:].columns:
        if df[l].sum() == 0:
            df.drop(columns=l,inplace=True)
    return df    
def extract_rebar(df):   
#根据钢筋明细df，获取按楼层汇总数据
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
    new_df['楼层'] = new_df['楼层'].str.extract(r'(\d+)',expand=False).astype(int)
    new_df['面积'] = new_df['面积'].round(2)
    new_df['墙'] = new_df[list_wall].sum(axis=1)
    new_df.drop(columns=list_wall,inplace=True)
    new_df[list_data]= new_df[list_data].div(1000).round(2)
    return new_df
def output(writer,df,sheetname,startrow=0,col_width=6):
#输出到excel
    df.to_excel(writer,index=False,sheet_name=sheetname,startrow=startrow)
    wb = writer.book
    format_data = wb.add_format({
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
        })
    ws = writer.sheets[sheetname]
    ws.set_column(0,len(df.columns)+10,col_width,format_data) 
def sum_df(df_conc,df_steel,df_rebar,df_conc_sum,\
            df_steel_sum,df_rebar_sum,scope):
#根据三个材料明细df，获取汇总数据   
    scope_global = scope[0]
    scope_permit = scope[1]
    df_sum = pd.DataFrame()
    columns_conc = df_conc.columns[2:].tolist()
    columns_steel = df_steel.columns[2:].tolist()
    columns_rebar = df_rebar.columns[2:].tolist()
    half_conc = len(columns_conc)//2
    half_steel = len(columns_steel)//2
    half_rebar = len(columns_rebar)//2
    col_conc0 = columns_conc[:half_conc]
    col_conc1 = columns_conc[half_conc:]
    col_steel0 = columns_steel[:half_steel]
    col_steel1 = columns_steel[half_steel:]
    col_rebar0 = columns_rebar[:half_rebar]
    col_rebar1 = columns_rebar[half_rebar:]
    df_sum_c0 = df_conc_sum[df_conc_sum['范围']==scope_global][col_conc0]
    df_sum_c0.columns = [ 'C'+col for col in df_sum_c0.columns]
    df_sum_s0 = df_steel_sum[df_steel_sum['范围']==scope_global][col_steel0]
    df_sum_s0.columns = ['S'+ col  for col in df_sum_s0.columns]
    df_sum_r0 = df_rebar_sum[df_rebar_sum['范围']==scope_global][col_rebar0]
    df_sum_r0.columns = ['R' + col  for col in df_sum_r0.columns]
    df_sum_c1 = df_conc_sum[df_conc_sum['范围']==scope_permit][col_conc1]    
    df_sum_c1.columns = ['C'+col for col in df_sum_c1.columns]
    df_sum_s1 = df_steel_sum[df_steel_sum['范围']==scope_permit][col_steel1]
    df_sum_s1.columns = ['S'+ col  for col in df_sum_s1.columns]
    df_sum_r1 = df_rebar_sum[df_rebar_sum['范围']==scope_permit][col_rebar1]
    df_sum_r1.columns = ['R'+ col  for col in df_sum_r1.columns]
    df_sum0 = pd.concat([df_sum_c0,df_sum_s0,df_sum_r0],axis=1).reset_index(drop=True)
    df_sum1 = pd.concat([df_sum_c1,df_sum_s1,df_sum_r1],axis=1).reset_index(drop=True)
    df_sum = pd.concat([df_sum0,df_sum1],axis=1).reset_index(drop=True)
    df_sum.insert(0,'面积',df_conc_sum[df_conc_sum['范围']==scope_global]['面积'].iloc[0])
    return df_sum
#主程序       
def main_program(direct):
# 定义初始变量，获取文件内容
    wmass_path = direct + '/设计结果/wmass.out'  
    quant_path = direct + '/上部结构工程量.txt' 
    rebar_path = direct + '/施工图/钢筋用量.xlsx'
    wmass_lines = read_wmass(wmass_path)
    quant_text = read_quant(quant_path)
    df_rebar_excel = repair_excel(rebar_path)
    num_beground = find_begrund_num(wmass_lines)
    str = direct.split('\\')[-1].split('-')[1]
    direct_output = direct + '/材料用量-统计-'+str+'.xlsx'
 #获取砼、型钢数据   
    df_conc0 = extract_conc(quant_text)
    df_steel0 = extract_steel(quant_text)
    df_rebar0 =extract_rebar(df_rebar_excel)
    df_conc,df_conc_sum = process_df(df_conc0,num_beground,1)
    df_steel,df_steel_sum = process_df(df_steel0,num_beground,1000)
    df_rebar,df_rebar_sum = process_df(df_rebar0,num_beground,1000)
    dict_scope = {
        'global':['全楼','地上'],
        'upground':['地上','地上'],
        }
#获取全楼总量+地上单方
    df_sum0 = sum_df(df_conc,df_steel,df_rebar,\
    df_conc_sum,df_steel_sum,df_rebar_sum,dict_scope['global'])
 #获取地上总量+地上单方
    df_sum1 = sum_df(df_conc,df_steel,df_rebar,\
    df_conc_sum,df_steel_sum,df_rebar_sum,dict_scope['upground'])
#增加版本列
    df_sum0.insert(0,'版本',str)
    df_sum1.insert(0,'版本',str)
#获取简化汇总（地上总量+地上单方）    
    col_simpl = ['版本','面积','C总','S总','R总','C总单','S总单','R总单']
    df_sum_simpl = df_sum1[col_simpl].copy()
    with pd.ExcelWriter(direct_output,engine='xlsxwriter') as writer:
        wb = writer.book
        output(writer,df_sum_simpl,'汇总（地上）',col_width=14)
#输出汇总明细        
        output(writer,df_sum0,'汇总明细',1,col_width=7)
        output(writer,df_sum1,'汇总明细',5,col_width=7)
        ws_hui = writer.sheets['汇总明细']
        ws_hui.merge_range(0,0,0,2,'全楼总量+地上单方')
        ws_hui.merge_range(4,0,4,2,'地上总量+地上单方')
#输出按材料汇总明细        
        output(writer,df_conc_sum,'分材料汇总',1)
        output(writer,df_steel_sum,'分材料汇总',7)
        output(writer,df_rebar_sum,'分材料汇总',13)
        ws_fen = writer.sheets['分材料汇总']
        ws_fen.merge_range(0,0,0,len(df_conc_sum.columns)-1,'混凝土汇总')   
        ws_fen.merge_range(6,0,6,len(df_steel_sum.columns)-1,'型钢汇总')
        ws_fen.merge_range(12,0,12,len(df_rebar_sum.columns)-1,'钢筋汇总')
#输出每层材料用量明细        
        output(writer,df_conc,'砼用量')
        output(writer,df_steel,'型钢用量')
        output(writer,df_rebar,'钢筋用量')
        ws_conc = writer.sheets['砼用量']
        ws_steel = writer.sheets['型钢用量']
        ws_rebar = writer.sheets['钢筋用量']
#绘制散点图
        columns_conc = df_conc.columns[2:].tolist()
        columns_steel = df_steel.columns[2:].tolist()
        columns_rebar = df_rebar.columns[2:].tolist()
        half_conc = len(columns_conc)//2
        half_steel = len(columns_steel)//2
        half_rebar = len(columns_rebar)//2
        col_conc0 = columns_conc[:half_conc]
        col_conc1 = columns_conc[half_conc:]
        col_steel0 = columns_steel[:half_steel]
        col_steel1 = columns_steel[half_steel:]
        col_rebar0 = columns_rebar[:half_rebar]
        col_rebar1 = columns_rebar[half_rebar:]  
        dict_para = {
            'C0':['地上砼总量','砼总量(m3)',1500,500],
            'C1':['地上砼单量','砼单量(m3/m2)',0.8,0.2],
            'S0':['地上型钢总量','型钢(t)',300,50],
            'S1':['地上型钢单量','型钢单量(kg/m2)',150,25],
            'R0':['地上钢筋总量','钢筋(t)',300,50],
            'R1':['地上钢筋单量','钢筋单量(kg/m2)',150,25],
        }    
        ps(df_conc,wb,ws_conc,col_conc0,num_beground,dict_para['C0'])
        ps(df_conc,wb,ws_conc,col_conc1,num_beground,dict_para['C1'])
        ps(df_steel,wb,ws_steel,col_steel0,num_beground,dict_para['S0'])
        ps(df_steel,wb,ws_steel,col_steel1,num_beground,dict_para['S1'])
        ps(df_rebar,wb,ws_rebar,col_rebar0,num_beground,dict_para['R0']) 
        ps(df_rebar,wb,ws_rebar,col_rebar1,num_beground,dict_para['R1'])
    return df_sum0,df_sum1,df_sum_simpl
if __name__ == "__main__":
        direct = input('Please input the YJK model directory:')
        direct = direct.rstrip('\\')
        df_sum0,df_sum1,df_sum_simpl = main_program(direct)
        pd.set_option('display.unicode.ambiguous_as_wide', True)
        pd.set_option('display.unicode.east_asian_width', True)
        print('\n输出地上总量与地上单方用量:\n')
        print('-'*70)
        print(df_sum_simpl.to_string(justify='center'))
        print('-'*70+'\n')
        print('统计完成')



