import pandas as pd
class DataProcessor:
    #1. 楼层映射
    @staticmethod
    def map_story(df,stories):
        story = [int(i) for i in stories.split('-')]
        story_range = range(story[0],story[1]+1)
        df = df[df['fl'].isin(story_range)].copy()
        df['fl'] = df['fl']-story[0]+1          
        return df    
    #2. 区分内力位移与整体指标
    @staticmethod
    def process_df(df,head_df,head_id,config):
        df_df = df[['fl']+head_df].copy()
        df_id = df[['fl']+head_id].copy()
        df_df =df_df.assign(**config.limit_df_float)
        df_id =df_id.assign(**config.limit_id)
        return df_df,df_id
    @staticmethod
    def merge_df(df1,df2):
        column_limit = [col for col in df1.columns if 'limit' in col]
        column_no_limit = [col for col in df1.columns if 'limit' not in col]
        df_limit = df1[['fl']+column_limit].copy()
        rename_columns1 = {col:f'{col}-1' for col in column_no_limit if col != 'fl'}
        rename_columns2 = {col:f'{col}-2' for col in column_no_limit if col != 'fl'}
        df1_renamed = df1[column_no_limit].rename(columns=rename_columns1)
        df2_renamed = df2[column_no_limit].rename(columns=rename_columns2)
        df_data = pd.merge(df1_renamed,df2_renamed,on = 'fl',how='outer')
        columns_origin = df1[column_no_limit].drop(columns=['fl']).columns.tolist()
        columns_sorted = ['fl'] + [f'{col}-{i}' for col in columns_origin for i in (1,2)]
        df_sorted = df_data[columns_sorted].copy()
        return pd.merge(df_sorted,df_limit,on='fl',how='outer')


