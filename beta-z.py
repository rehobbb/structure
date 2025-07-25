def extract_beta_z():
    beta_z_0 = []   # 存储0度风振系数
    beta_z_90 = []  # 存储90度风振系数
    current_floor = 0
    in_beta_section = False
    wind_direction = None

    try:
        with open('yjkwindforce.txt', 'r', encoding='gbk') as file:
            for line in file:
                # 检测新的楼层开始
                if line.startswith('[第') and '层]' in line:
                    parts = line.split(']')[0].split('第')
                    if len(parts) > 1:
                        current_floor = int(parts[1].replace('层', ''))
                        if current_floor > 30:
                            break  # 超过30层停止处理
                
                # 检测进入风振系数部分
                elif "顺风向风振系数βz:" in line:
                    in_beta_section = True
                    wind_direction = None
                
                # 检测风向变化
                elif in_beta_section and "风向" in line:
                    if "风向0度:" in line:
                        wind_direction = 0
                    elif "风向90度:" in line:
                        wind_direction = 90
                    else:
                        wind_direction = None  # 忽略其他风向
                
                # 提取βz值
                elif in_beta_section and line.strip().startswith('βz') and wind_direction is not None:
                    parts = line.split()
                    if len(parts) >= 2:
                        try:
                            beta_value = float(parts[-1])
                            if wind_direction == 0:
                                beta_z_0.append(beta_value)
                            elif wind_direction == 90:
                                beta_z_90.append(beta_value)
                            wind_direction = None  # 重置风向
                        except ValueError:
                            print(f"格式错误: {line.strip()}")
                
                # 检测风荷载部分开始（表示风振系数部分结束）
                elif "风荷载F(kN):" in line:
                    in_beta_section = False
                    wind_direction = None
    
    except FileNotFoundError:
        print("错误：文件 'yjkwindforce.txt' 不存在！")
        return [], []
    except Exception as e:
        print(f"处理文件时出错: {e}")
        return [], []
    
    return beta_z_0, beta_z_90

# 执行并打印结果
beta_0, beta_90 = extract_beta_z()
print("0° 风振系数 βz 列表：", beta_0)
print("90° 风振系数 βz 列表：", beta_90)
print(f"\n提取完成：0度方向共{len(beta_0)}条，90度方向共{len(beta_90)}条")