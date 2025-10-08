import pandas as pd
import xlsxwriter

def merge_and_write_data_advanced(data_a, data_b, output_file="merged_data.xlsx"):
    """
    高级版本：将两个数据列表合并并按照特定格式写入Excel
    
    参数:
    data_a -- 第一个字典列表数据
    data_b -- 第二个列表数据
    output_file -- 输出的Excel文件名
    """
    # 创建DataFrame并设置fl列
    df = pd.DataFrame({
        'fl': list(range(27, 0, -1))
    })
    
    # 提取列名（排除fl列）
    columns = ['ex_df', 'ey_df', 'wx_df', 'wy_df']
    
    # 将data_a和data_b转换为DataFrame
    df_a = pd.DataFrame(data_a)
    df_b = pd.DataFrame(data_b)
    
    # 确保数据长度一致
    if len(df_a) != len(df_b) or len(df_a) != len(df):
        raise ValueError("数据长度不一致，请检查数据")
    
    # 创建ExcelWriter对象
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # 创建工作表
        workbook = writer.book
        worksheet = workbook.add_worksheet('合并数据')
        
        # 设置格式
        header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'bg_color': '#E6E6E6'
        })
        
        data_format = workbook.add_format({
            'num_format': '0.000000',
            'border': 1,
            'align': 'right'
        })
        
        fl_format = workbook.add_format({
            'border': 1,
            'align': 'center'
        })
        
        # 写入表头
        worksheet.write(0, 0, 'fl', header_format)
        
        # 写入列标题并合并单元格
        col_idx = 1
        for col in columns:
            # 合并单元格作为主标题
            worksheet.merge_range(0, col_idx, 0, col_idx+1, col, header_format)
            # 写入子标题
            worksheet.write(1, col_idx, 'A', header_format)
            worksheet.write(1, col_idx+1, 'B', header_format)
            col_idx += 2
        
        # 写入fl列数据
        for row_idx, value in enumerate(df['fl'], 2):
            worksheet.write(row_idx, 0, value, fl_format)
        
        # 写入数据A和数据B
        for col_idx, col in enumerate(columns):
            # 计算数据列的位置 (每列占两列)
            data_col = 1 + col_idx * 2
            
            # 写入data_a的数据
            for row_idx, value in enumerate(df_a[col], 2):
                worksheet.write(row_idx, data_col, value, data_format)
            
            # 写入data_b的数据
            for row_idx, value in enumerate(df_b[col], 2):
                worksheet.write(row_idx, data_col + 1, value, data_format)
        
        # 自动调整列宽
        worksheet.set_column(0, 0, 8)  # fl列宽度
        
        # 数据列宽度
        for i in range(len(columns)):
            worksheet.set_column(i*2+1, i*2+2, 12)
        
        # 添加冻结窗格（冻结前两行和第一列）
        worksheet.freeze_panes(2, 1)
        
        print(f"数据已成功写入 {output_file}")

# 示例使用
if __name__ == "__main__":
    # 这里应该替换为您的实际数据
    # data_a 和 data_b 应该是包含27个元素的列表
    # 每个元素是一个字典，包含ex_df, ey_df, wx_df, wy_df四个键
    
    # 示例数据
    data_a = [
        {'ex_df': 0.000409, 'ey_df': 0.000316, 'wx_df': 0.000438, 'wy_df': 0.000376},
        {'ex_df': 0.000508, 'ey_df': 0.000529, 'wx_df': 0.000525, 'wy_df': 0.000540},
        # ... 其他23行数据
        {'ex_df': 0.000100, 'ey_df': 0.000100, 'wx_df': 0.000100, 'wy_df': 0.000100}
    ]
    
    data_b = [
        {'ex_df': 0.000500, 'ey_df': 0.000400, 'wx_df': 0.000500, 'wy_df': 0.000400},
        {'ex_df': 0.000600, 'ey_df': 0.000600, 'wx_df': 0.000600, 'wy_df': 0.000600},
        # ... 其他23行数据
        {'ex_df': 0.000200, 'ey_df': 0.000200, 'wx_df': 0.000200, 'wy_df': 0.000200}
    ]
    
    # 确保数据有27行
    if len(data_a) != 27 or len(data_b) != 27:
        print("警告：数据行数不是27行，结果可能不符合预期")
    
    # 合并并写入数据
    merge_and_write_data_advanced(data_a, data_b, "合并数据.xlsx")