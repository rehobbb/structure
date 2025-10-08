import pandas as pd
import re
from b_config import AppConfig
from c_file_processor import  FileProcessor
from d_data_extractor import DataExtractor
from e_data_processor import DataProcessor
from f_excel_output import ExcelOutput
#1. 主函数
def main():
    cfg0 = AppConfig('')
    print('楼层范围:',f'  {cfg0.stories}')
    print('位移角限值:',f'地震:{cfg0.limit_df['limit_dfe']}',
                          f'风:{cfg0.limit_df['limit_dfw']}')
    print('位移比限值:',f'{cfg0.limit_id['limit_drd']}~\
{cfg0.limit_id["limit_dru"]}')
    type=input(
    '''Select the type of analysis:
1-YJK
2-YJK vs YJK
3-YJI vs Midas\n''')
    #1.1 单独YJK分析
    if type == '1':
        direct=input('Please input the directory of YJK:')
        df_df,df_id,cfg = process_single_yjk(direct)
        #1.1.1 输出Excel
        print('Outputting Excel...')
        with pd.ExcelWriter(cfg.direct+'/Result.xlsx',engine='xlsxwriter') as writer:
            ExcelOutput.write_df(df_df,writer,sheet_name='内力位移',head=cfg.head_df)
            ExcelOutput.write_df(df_id,writer,sheet_name='整体指标',head=cfg.head_id)
            wb = writer.book
            ws_df = writer.sheets['内力位移']
            ws_id = writer.sheets['整体指标']
            ExcelOutput.plot_scatter_chart(df_df,wb,ws_df,cfg.head_df,
                                            cfg.plot_df,cfg.limit_df_float,type)
            ExcelOutput.plot_scatter_chart(df_id,wb,ws_id,cfg.head_id,
                                            cfg.plot_id,cfg.limit_id,type)
        print('Done!')
    #1.2 YJK对比YJK分析
    elif type == '2':
        direct1=input('Please input the directory of YJK-1:')
        direct2=input('Please input the directory of YJK-2:')
        df_df1,df_id1,cfg = process_single_yjk(direct1)
        df_df2,df_id2,cfg = process_single_yjk(direct2)
        df_df = DataProcessor.merge_df(df_df1,df_df2)
        df_id = DataProcessor.merge_df(df_id1,df_id2)
        with pd.ExcelWriter(cfg.direct+'\\Result.xlsx',engine='xlsxwriter') as writer:
            ExcelOutput.write_df(df_df,writer,sheet_name='内力位移',head=cfg.head_df)
            ExcelOutput.write_df(df_id,writer,sheet_name='整体指标',head=cfg.head_id)
            wb = writer.book
            ws_df = writer.sheets['内力位移']
            ws_id = writer.sheets['整体指标']
            ExcelOutput.plot_scatter_chart(df_df,wb,ws_df,cfg.head_df,
                                            cfg.plot_df,cfg.limit_df_float,type)
            ExcelOutput.plot_scatter_chart(df_id,wb,ws_id,cfg.head_id,
                                            cfg.plot_id,cfg.limit_id,type)            
    #1.3 YJI对比Midas分析
    elif type == '3':
        direct = input('Please input the directory of YJK:')
        df_df1,df_id1,cfg = process_single_yjk(direct)
        df_df2,df_id2,cfg = process_single_midas(direct)

#2. 单个YJK目录分析
def process_single_yjk(direct):
    cfg = AppConfig(direct)
    print('Reading files...')
    wmass_lines = FileProcessor.read_file(cfg.direct_wmass)
    wzq_lines = FileProcessor.read_file(cfg.direct_wzq)
    wdisp_lines = FileProcessor.read_file(cfg.direct_wdisp)
    wv02q_lines = FileProcessor.read_file(cfg.direct_wv02q)
    #2.1 初始化数据提取器
    extractor = DataExtractor(cfg)
    data_a = {}
    #2.2 提取wmass数据-风内力，刚度比，抗剪承载力比
    print('Extracting wmass data...')
    #2.2.1 提取风内力
    chunk_wforce = FileProcessor.find_chunk(
        cfg.indicator_wforce,
        cfg.endflag_wforce,
        wmass_lines,
    )
    chunk_wforce = [s for s in chunk_wforce if s.strip()]
    extractor.extract_wforce(chunk_wforce,data_a)
    #2.2.2 提取刚度比
    for line in wmass_lines:
        if '结构总体信息' in line:
            s_tmp = wmass_lines[wmass_lines.index(line)+1]
            s_stru = re.search(r'\s*\w+:\s*(\w+)',s_tmp).group(1)
            break
    chunk_ratios = FileProcessor.find_chunk(
        cfg.indicator_ratios,
        cfg.endflag_ratios,
        wmass_lines,
    )
    chunk_ratios = [s for s in chunk_ratios if s.strip()]
    extractor.extract_ratios(chunk_ratios,data_a,s_stru)
    #2.2.3 提取抗剪承载力比
    chunk_ratiovc = FileProcessor.find_chunk(
        cfg.indicator_ratiovc,
        cfg.endflag_ratiovc,
        wmass_lines,
    )
    chunk_ratiovc = [s for s in chunk_ratiovc if s.strip()]
    extractor.extract_ratiovc(chunk_ratiovc,data_a)
    #2.3 提取wzq数据-地震力,剪重比数据  
    print('Extracting wzq data...')
    chunks_eforce = FileProcessor.find_chunks(
        cfg.indicators_eforce,
        cfg.endflag_eforce,
        wzq_lines,
    )
    chunks_eforce = [[s for s in tmpchunk if s.strip()]
                     for tmpchunk in chunks_eforce]
    for chunk,key in zip(chunks_eforce,cfg.indicators_eforce.keys()):
        extractor.extract_eforce(chunk,data_a,key)   
    #2.4 提取wdisp数据-地震及风位移、位移角
    print('Extracting wdisp data...')
    chunks_disp = FileProcessor.find_chunks(
        cfg.indicators_disp,
        cfg.endflag_disp,
        wdisp_lines,
    )
    chunks_disp = [[s for s in tmpchunk if s.strip()] 
                    for tmpchunk in chunks_disp]
    for chunk,key in zip(chunks_disp,cfg.indicators_disp.keys()):
        extractor.extract_disp(chunk,data_a,key,cfg.f_ratio)
    #2.5 提取wv02q数据-框架倾覆弯矩，剪力占比数据
    #2.5.1 提取框架倾覆弯矩比
    print('Extracting wv02q data...')
    chunk_ratiom = FileProcessor.find_chunk(
        cfg.indicator_ratiom,
        cfg.endflag_ratiom,
        wv02q_lines,
    )
    chunk_ratiom = [s for s in chunk_ratiom if s.strip()]
    extractor.extract_ratiom(chunk_ratiom,data_a)
    #2.5.2 提取框架剪力占比
    chunk_ratiov0 = FileProcessor.find_chunk(
        cfg.indicator_ratiov0,
        cfg.endflag_ratiov0,
        wv02q_lines,
    )
    chunk_ratiov0 = [s for s in chunk_ratiov0 if s.strip()]
    extractor.extract_ratiov0(chunk_ratiov0,data_a)    
    #2.6 处理数据
    print('Processing data...')
    df = pd.DataFrame(data_a)
    df = DataProcessor.map_story(df,cfg.stories)
    df_df,df_id = DataProcessor.process_df(
        df,
        list(cfg.head_df.keys()),
        list(cfg.head_id.keys()),
        cfg,
    )
    return df_df,df_id,cfg
if __name__ == '__main__':
    main()
