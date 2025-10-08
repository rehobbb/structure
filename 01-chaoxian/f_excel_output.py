from b_config import AppConfig
import re
class ExcelOutput:
    @staticmethod
    def write_df(df,writer,sheet_name,head):
        old_df = df.copy()
        columns_head = list(head.keys())
        columns = df.columns.tolist()
        for col in columns:
            if 'df' in col:
                df[col] = df[col].apply(AppConfig.fraction_to_float)
        df.to_excel(writer,sheet_name=sheet_name,
                    startrow = 2,index =False,header=False)
        ws = writer.sheets[sheet_name]
        for col_idx,col_name in enumerate(columns):
            if col_name == 'fl':
                ws.write(len(df)+4,col_idx,'最大值')
                ws.write(len(df)+5,col_idx,'所在楼层')
            col_name_ori = re.sub(r'-\d$','',col_name)
            if col_name_ori in columns_head:
                ws.write(1,col_idx,'YJK'+col_name[len(col_name_ori):])
                ws.write(len(df)+4,col_idx,
                        old_df[df[col_name]==df[col_name].max()][col_name].values[0])
                ws.write(len(df)+5,col_idx,
                        old_df[df[col_name]==df[col_name].max()]['fl'].values[0])
            elif 'limit' in col_name:
                ws.write(1,col_idx,col_name)
        ws.write(0,0,'楼层')
        columns_ori = [re.sub(r'-\d$','',col) for col in columns]
        start_cols = []
        seen = set()
        for i,col in enumerate(columns_ori):
            if col in columns_head and col not in seen:
                start_cols.append(i)
                seen.add(col)
        len_columns_data = len([col for col in columns if 'limit' not in col])
        for i,name in enumerate(columns_head):          
            start_col = start_cols[i]
            end_col = start_cols[i+1]-1 if i<len(columns_head)-1 else len_columns_data-1
            if end_col - start_col  == 0:
                ws.write(0,start_col,head[name])
                ws.write(len(df)+3,start_col,head[name])
            else:
                ws.merge_range(0,start_col,0,end_col,head[name])
                ws.merge_range(len(df)+3,start_col,len(df)+3,end_col,head[name])
    @staticmethod
    def plot_scatter_chart(df,wb,ws,head,plot,limit,type):
        columns_head = list(head.keys())
        format_data = wb.add_format({
            'align':'center',
            'valign':'vcenter',
            'text_wrap':True,
        })
        ws.set_column(0,len(df.columns)+10,10,format_data)
        j,k = 0,0
        for col in columns_head:
            chart = wb.add_chart({
                'type':'scatter',
                'subtype':'smooth',
            })
            if type == '1':
                chart.add_series({
                    'name':'YJK',
                    'categories':[ws.name,2,df.columns.get_loc(col),
                                len(df)+1,df.columns.get_loc(col)],
                    'values':[ws.name,2,0,len(df)+1,0],
                    'line':{'color':'blue','width':2.25}
                })
            elif type == '2':
                col1 = col+'-1'
                col2 = col+'-2'
                chart.add_series({
                    'name':'YJK1',
                    'categories':[ws.name,2,df.columns.get_loc(col1),
                                len(df)+1,df.columns.get_loc(col1)],
                    'values':[ws.name,2,0,len(df)+1,0],
                    'line':{'color':'blue','width':2.25}
                })
                chart.add_series({
                    'name':'YJK2',
                    'categories':[ws.name,2,df.columns.get_loc(col2),
                                len(df)+1,df.columns.get_loc(col2)],
                    'values':[ws.name,2,0,len(df)+1,0],
                    'line':{'color':'green','width':2.25}
                })
            if plot[col][1] != '':
                limit_col = plot[col][1]
                if limit_col in limit.keys():
                    chart.add_series(
                        {
                            'name':'限值',
                            'categories':[ws.name,2,df.columns.get_loc(limit_col),
                                        len(df)+1,df.columns.get_loc(limit_col)],
                            'values':[ws.name,2,0,
                                        len(df)+1,0],
                            'line':{'color':'red','width':2.25,'dash_type':'dash'}
                        }
                    )
                else:
                    limit_col_d = limit_col+'d'
                    limit_col_u = limit_col+'u'
                    chart.add_series(
                        {
                            'name':'限值',
                            'categories':[ws.name,2,df.columns.get_loc(limit_col_d),
                                        len(df)+1,df.columns.get_loc(limit_col_d)],
                            'values':[ws.name,2,0,
                                        len(df)+1,0],
                            'line':{'color':'red','width':2.25,'dash_type':'dash'}
                        }
                    )
                    chart.add_series(
                        {
                            'name':'限值',
                            'categories':[ws.name,2,df.columns.get_loc(limit_col_u),
                                        len(df)+1,df.columns.get_loc(limit_col_u)],
                            'values':[ws.name,2,0,
                                        len(df)+1,0],
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
                    'line':{
                        'color':'gray',
                        'width':0.5,
                        'dash_type':'long_dash'
                    }
                },
            })
            chart.set_y_axis({
                'name':'楼层',
                'num_font':{'size':10},
                'name_font':{'size':10},
                'major_gridlines':{
                    'visible':True,
                    'line':{
                        'color':'gray',
                        'width':0.5,
                        'dash_type':'long_dash'
                    }
                },
            })
            cm_to_px_width = 36.5
            cm_to_px_heigh = 41
            chart.set_size({
                'width':int(7 * cm_to_px_width),
                'height':int(9.5 * cm_to_px_heigh)
            })
            chart.set_legend({
                'position':'right',
                'font':{
                    'name':'Arial',
                    'size':10,
                    'bold':False,
                    'color':'black'
                },
                'border':{
                    'color':'black',
                    'width':0.7,
                    'dash_type':'solid',
                },
                # 'layout':{
                #     'x':0.55,
                #     'y':0.4,
                #     'width':0.28,
                #     'height':0.07,
                # },
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
                chart_col = j*(chart.width//63)
                ws.insert_chart(len(df)+7,chart_col,chart,{
                    'x_offset':0,
                    'y_offset':0,
                })
                j += 1
            else:
                chart_col = k*(chart.width//63)
                ws.insert_chart(len(df)+28,chart_col,chart,{
                    'x_offset':0,
                    'y_offset':0,
                })
                k += 1
