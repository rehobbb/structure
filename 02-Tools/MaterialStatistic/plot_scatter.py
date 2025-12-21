import random
def plot_scatter(df,wb,ws,head,num_beground,para):
        colors = ['#FF0000', '#00FF00', '#0000FF', '#FFFF00', \
                 '#FF00FF', '#00FFFF', '#FFA500', '#800080']
        chart = wb.add_chart({
            'type':'scatter',
            'subtype':'smooth',
        })
        for idx,col in enumerate(head):
            line_color = colors[idx % len(colors)]
            chart.add_series({
                'name':col,
                'categories':[ws.name,num_beground+1,df.columns.get_loc(col),
                            len(df),df.columns.get_loc(col)],
                'values':[ws.name,num_beground+1,0,len(df),0],
                'line':{'color':line_color,'width':2.25}
            })
        chart.set_title({
            'name':para[0],
            'name_font':{
                'name':'Arial',
                'size':12,
                'bold':True,
                'color':'black'
            },
        })
        chart.set_x_axis({
            'name':para[1],
            'min':0,
            'max':para[2],
            'major_unit':para[3],
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
            'min':num_beground+1,
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
        if '总' in para[0]:
            ws.insert_chart(0,len(df.columns)+1,chart)
        else:
            ws.insert_chart(0,len(df.columns)+7,chart)

