import openpyxl
import pandas as pd
import zipfile
import os
import tempfile
directory = 'D:/03-学习/07-PYTHON/'
file_path = directory + '钢筋用量.xlsx'
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
                print("文件修复并读取成功！")
                
        except Exception as e2:
            raise Exception(f"文件修复失败: {str(e2)}") from e
    return df
#主程序       
if __name__ == "__main__":
    num_beground = 4
    rebar_list = {}
    df = repair_excel(file_path)
    df_data = df[['楼层','构件类别','楼面面积(m2)','合计(kg)']].iloc[1:-2].copy().ffill()
    list=df_data['构件类别'].unique().tolist()
    list_ordi= ['板','梁','柱']
    list_wall = [i for i in list if i not in list_ordi]
    stories = df_data['楼层'].unique()
    stories_beground = stories[:num_beground]
    stories_upground = stories[num_beground:]
    condition_beground = df_data['楼层'].isin(stories_beground)
    condition_upground = df_data['楼层'].isin(stories_upground)   
    df_b = df_data[condition_beground]
    df_u = df_data[condition_upground]
    rebar_list['范围'] = ['地下','地上']
    rebar_list['面积'] = []
    rebar_list['面积'].append(df_b.groupby('楼层')['楼面面积(m2)'].first().sum())
    rebar_list['面积'].append(df_u.groupby('楼层')['楼面面积(m2)'].first().sum())
    series_b=df_b.groupby('构件类别')['合计(kg)'].sum()/1000
    series_u=df_b.groupby('构件类别')['合计(kg)'].sum()/1000
    for l in list :
        rebar_list[l] = []
        rebar_list[l].append(series_b[l])
        rebar_list[l].append(series_u[l])
    # for l in list:
    #     rebar_list[l]=[]
    #     rebar_ele = df_b[(df_b['构件类别'] == l)]['合计(kg)'].sum()/1000
    #     rebar_list[l].append(rebar_ele)
    #     rebar_ele = df_u[(df_u['构件类别'] == l)]['合计(kg)'].sum()/1000
    #     rebar_list[l].append(rebar_ele)
    df_rebar = pd.DataFrame(rebar_list)
    df_rebar.iloc[:] = df_rebar.iloc[:].round(1)
    df_rebar['墙'] = df_rebar[list_wall].sum(axis=1)
    df_rebar['总'] = df_rebar[list].sum(axis=1)
    df_rebar.drop(columns=list_wall,inplace=True)
    list_colu = df_rebar.columns[2:].tolist()
    for l in list_colu:
        df_rebar[f'{l}-单方'] = (df_rebar[l]/df_rebar['面积']*1000).round(1)
    df_rebar.to_excel(directory + '钢筋用量-统计.xlsx',index=False)
    print(df_rebar)
    print('统计完成')


