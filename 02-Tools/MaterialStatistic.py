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
            txt_file = f.read()
        txt_lines = re.split(r'(>第\s*\w+层:)',txt_file)
        result = []
        for i in range(1,len(txt_lines),2):
            mark = txt_lines[i]
            content = txt_lines[i+1] if (i+1) < len(txt_lines) else ''
            result.append(mark+content.strip())
        result[-1] = result[-1].split('>全楼统计:')[0]
        return result
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
def extract_concsteel(chunks,num_beground):
        data.setdefault('楼层',[])
        data.setdefault('板-砼',[])
        data.setdefault('悬板-砼',[])
        data.setdefault('梁-砼',[])
        data.setdefault('柱-砼',[])
        data.setdefault('墙-砼',[])
        data.setdefault('梁-钢',[])
        data.setdefault('柱-钢',[])
        conc_flag,steel_flag = 0,0
        for chunk in chunks:
            


        # chunk = [ck for ck in chunk if ck.strip()]
        # for list in chunk:
        #     floor_match = re.search(r'第\s*(\d+)自然层',list)
        #     if floor_match:
        #         floor = int(floor_match.group(1))
        #         data['楼层'].append(floor)
        #         continue
        #     if re.search('砼等级',list):
        #         conc_flag = 1
        #     if re.search('钢等级',list):
        #         steel_flag = 1
        #     if conc_flag and floor is not None:
        #         list_conc = list.split()
        #         if list_conc[0] == '楼板':
        #             floor_conc = [float(l) for l in list_conc[2:] if l.replace('.','',1).isdigit()]
        #             data['板-砼'].append(round(sum(floor_conc),1))
        #         if list_conc[0] == '悬挑板32':
        #             floor_conc = [float(l) for l in list_conc[2:] if l.replace('.','',1).isdigit()]
        #             data['悬板-砼'].append(round(sum(floor_conc),1))   
        #         if list_conc[0] == '梁':
        #             beam_conc = [float(l) for l in list_conc[2:] if l.replace('.','',1).isdigit()]
        #             data['梁-砼'].append(round(sum(beam_conc),1))      
        #         if list_conc[0] == '柱':
        #             beam_conc = [float(l) for l in list_conc[2:] if l.replace('.','',1).isdigit()]
        #             data['柱-砼'].append(round(sum(beam_conc),1))                           
        #         if list_conc[0] == '墙(总计)':
        #             wall_conc = [float(l) for l in list_conc[2:] if l.replace('.','',1).isdigit()]
        #             data['墙-砼'].append(round(sum(wall_conc),1))
        #         if '总面积' in list:
        #             conc_flag = 0
        #     if steel_flag and floor is not None:
        #         list_steel = list.split()
        #         if list_steel[0] == '梁':
        #             beam_steel = [float(l) for l in list_steel[2:] if l.replace('.','',1).isdigit()]
        #             data['梁-钢'].append(round(sum(beam_steel),1))      
        #         if list_steel[0] == '柱':
        #             col_steel = [float(l) for l in list_steel[2:] if l.replace('.','',1).isdigit()]
        #             data['柱-钢'].append(round(sum(col_steel),1))     
        #         if '小计' in list:
        #             steel_flag = 0                      
        #             floor = None
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
    wmass_lines = read_text(wmass_path)
    quant_chunks = read_text2(quant_path)
    df_rebar = repair_excel(rebar_path)
    num_beground = find_begrund_num(wmass_lines)
    concsteel_dict = {}
 #获取砼、型钢数据   
    extract_concsteel(quant_chunks,num_beground)
    df_rebar1,df_rebar_sum =extract_rebar(df_rebar,num_beground)
    str = direct.split('\\')[-1].split('-')[1]
    direct_output = direct + '/钢筋用量-统计-'+str+'.xlsx'
    with pd.ExcelWriter(direct_output,engine='xlsxwriter') as writer:
        output(writer,df_rebar_sum,'钢筋用量-统计')
        output(writer,df_rebar1,'钢筋用量')
    print(df_rebar_sum)
    print('统计完成')



