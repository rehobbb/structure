import pandas as pd
import xlsxwriter
data ={
    'fl': [27, 26, 25, 24, 23, 22, 21, 20, 19, 18, 17, 16, 15, 14, 13, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1],
    'ey+_dr': [1.05, 1.03, 1.03, 1.03, 1.03, 1.03, 1.03, 1.03, 1.03, 1.03, 1.03, 1.03, 1.03, 1.03, 1.03, 1.03, 1.03, 1.03, 1.03, 1.04, 1.05, 1.07, 1.08, 1.10, 1.14, 1.11, 1.00]
}
df = pd.DataFrame(data)
excel_file = '位移比.xlsx'
sheet_name = '数据'
with pd.ExcelWriter(excel_file,engine='xlsxwriter') as writer:
    df.to_excel(writer,sheet_name=sheet_name,index=False)
    workbook =writer.book
    worksheet = writer.sheets[sheet_name]
    chart = workbook.add_chart({'type':'scatter','subtype':'smooth'})
    chart.add_series(
        {
            'name':'YJK',
            'categories':[sheet_name,1,1,len(df),1],
            'values':[sheet_name,1,0,len(df),0],
            'line':{'color':'blue','width':2.25}       
        }
    )
    chart.set_title({
        'name':'楼层-位移比关系',
        'name_font': {
        'name': 'Arial',
        'size': 12,
        'bold': True,
        'color': 'black'
        }
        })
    chart.set_x_axis({
        'name':'位移比',
        'num_font':{'size':10},
        'name_font':{'size':10},
        'major_gridlines': {
            'visible': True,
            'line': {'color': 'gray','width':0.5,'dash_type':'long_dash'}
            },          
        'minor_gridlines': {
            'visible': False,
        }
    })
    chart.set_y_axis({
        'name':'楼层',
        'num_font':{'size':10},
        'name_font':{'size':10},
        'major_gridlines': {
            'visible': True,
            'line': {'color': 'gray','width':0.5,'dash_type':'long_dash'}
            },
        'minor_gridlines': {
            'visible': False,
        }
    })
    cm_to_px_width = 36.5
    cm_to_px_height = 41
    chart.set_size({
        'width':int(7.5*cm_to_px_width),
        'height':int(9.5*cm_to_px_height)
    })
    chart.set_legend({
        'position':'right',
        'font': {
        'name': 'Arial',      # 字体名称
        'size': 10,           # 字体大小
        'bold': False,        # 是否加粗
        'italic': False,       # 是否斜体
        'color': 'black'    # 字体颜色
        },
        'border': {
        'color': 'black',   # 边框颜色
        'width': 0.7,         # 边框宽度（磅）
        'dash_type': 'solid'  # 边框样式
        },
        'layout': {
        'x': 0.6,        # 水平位置（0-1）
        'y': 0.4,        # 垂直位置（0-1）
        'width': 0.25,    # 宽度比例
        'height': 0.1    # 高度比例
        },
    })
    chart.set_plotarea({
    'border': {'none': True},
    'layout': {
        'x':      0.2,  # 左边距比例
        'y':      0.15,  # 上边距比例
        'width':  0.65,  # 绘图区宽度比例
        'height': 0.7   # 绘图区高度比例
    }
})
    worksheet.insert_chart('E10',chart)
    worksheet.set_column('A:B',12)