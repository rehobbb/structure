import math
import pandas as pd
def cal_pressure_coefficient(H, gamma, c, phi, q, delta, beta, alpha, verbose):
    """
    计算主动土压力系数 Ka (根据GB 50330-2013公式6.2.3)
    
    参数:
    H (float): 挡土墙高度 (m)
    gamma (float): 土体重度 (kN/m³)
    c (float): 土的黏聚力 (kPa)
    phi (float): 土的内摩擦角 (°)
    q (float): 地表均布荷载标准值 (kN/m²)
    delta (float): 土对挡土墙墙背的摩擦角 (°)
    beta (float): 填土表面与水平面的夹角 (°)
    alpha (float): 支撑结构墙背与水平面的夹角 (°)
    verbose (bool): 是否输出计算过程
    
    返回:
    float: 主动土压力系数 Ka
    """
    # 将角度转换为弧度
    phi_rad = math.radians(phi)
    delta_rad = math.radians(delta)
    beta_rad = math.radians(beta)
    alpha_rad = math.radians(alpha)
    
    if verbose:
        print("=== 主动土压力系数 Ka 计算过程 ===")
        print(f"输入参数:")
        print(f"  H = {H} m")
        print(f"  γ = {gamma} kN/m³")
        print(f"  c = {c} kPa")
        print(f"  φ = {phi}°")
        print(f"  q = {q} kN/m²")
        print(f"  δ = {delta}°")
        print(f"  β = {beta}°")
        print(f"  α = {alpha}°")
        print()
    
    # 计算 η (公式 6.2.3-4)
    eta = 2 * c / (gamma * H)
    if verbose:
        print(f"1. 计算 η = 2c/(γH) = 2*{c}/({gamma}*{H}) = {eta:.{de_p}f}")
    
    # 计算 Kq (公式 6.2.3-3)
    numerator_Kq = 2 * q * math.sin(alpha_rad) * math.cos(beta_rad)
    denominator_Kq = gamma * H * math.sin(alpha_rad + beta_rad)
    Kq = 1 + numerator_Kq / denominator_Kq 
    
    if verbose:
        print(f"2. 计算 Kq = 1 + (2ηsinαcosβ)/(γHsin(α+β)) + (2qsinαcosβ)/(γHsin(α+β))")
        print(f"   = 1 + (2*{eta:.{de_p}f}*sin({alpha})*cos({beta}))/({gamma}*{H}*sin({alpha}+{beta})) + (2*{q}*sin({alpha})*cos({beta}))/({gamma}*{H}*sin({alpha}+{beta}))")
        print(f"   = 1 + ({numerator_Kq:.{de_p}f})/({denominator_Kq:.{de_p}f}) + ({2*q*math.sin(alpha_rad)*math.cos(beta_rad):.{de_p}f})/({denominator_Kq:.{de_p}f})")
        print(f"   = {Kq:.{de_p}f}")
    
    # 计算 Ka 的分子部分
    Ka_numerator = math.sin(alpha_rad + beta_rad)
    if verbose:
        print(f"3. 计算 Ka 分子 = sin(α+β) = sin({alpha}+{beta}) = {Ka_numerator:.{de_p}f}")
    
    # 计算 Ka 的分母部分
    Ka_denominator = (math.sin(alpha_rad)**2) * (math.sin(alpha_rad + beta_rad - phi_rad - delta_rad)**2)
    if verbose:
        print(f"4. 计算 Ka 分母 = sin²α * sin²(α+β-φ-δ)")
        print(f"   = sin²({alpha}) * sin²({alpha}+{beta}-{phi}-{delta})")
        print(f"   = {math.sin(alpha_rad)**2:.{de_p}f} * {math.sin(alpha_rad + beta_rad - phi_rad - delta_rad)**2:.{de_p}f}")
        print(f"   = {Ka_denominator:.{de_p}f}")
    
    # 计算 Ka 的大括号内的复杂表达式
    # 第一部分: Kq[sin(α+δ)sin(α-δ)] + sin(φ+δ)sin(φ-β)
    part1 = Kq * (math.sin(alpha_rad + delta_rad) * math.sin(alpha_rad - delta_rad) + 
            math.sin(phi_rad + delta_rad) * math.sin(phi_rad - beta_rad))
    
    if verbose:
        print(f"5. 计算大括号内第一部分: Kq[sin(α+δ)sin(α-δ) + sin(φ+δ)sin(φ-β)]")
        print(f"   = {Kq:.{de_p}f}[sin({alpha}+{delta})sin({alpha}-{delta})] + sin({phi}+{delta})sin({phi}-{beta})")
        print(f"   = {Kq:.{de_p}f}[{math.sin(alpha_rad + delta_rad):.{de_p}f}*{math.sin(alpha_rad - delta_rad):.{de_p}f} \
        + {math.sin(phi_rad + delta_rad):.{de_p}f}*{math.sin(phi_rad - beta_rad):.{de_p}f}]")
        print(f"   = {part1:.{de_p}f}")
    
    # 第二部分: 2ηsinαcosφcos(α+β-φ-δ)
    part2 = 2 * eta * math.sin(alpha_rad) * math.cos(phi_rad) * math.cos(alpha_rad + beta_rad - phi_rad - delta_rad)
    if verbose:
        print(f"6. 计算大括号内第二部分: 2ηsinαcosφcos(α+β-φ-δ)")
        print(f"   = 2*{eta:.{de_p}f}*sin({alpha})*cos({phi})*cos({alpha}+{beta}-{phi}-{delta})")
        print(f"   = {part2:.{de_p}f}")
    
    # 第三部分: 2√Kq sin(α+β) sin(φ-β)+ηsinαcosφ
    part3 = 2 * math.sqrt(Kq * math.sin(alpha_rad + beta_rad) * math.sin(phi_rad - beta_rad)+eta * math.sin(alpha_rad) * math.cos(phi_rad))
    if verbose:
        print(f"7. 计算大括号内第三部分: 2√Kq sin(α+β) sin(φ-β)+ηsinαcosφ")
        print(f"   = 2*√{Kq:.{de_p}f}*sin({alpha}+{beta})*sin({phi}-{beta})+{eta:.{de_p}f}*sin({alpha})*cos({phi})")
        print(f"   = 2*√{Kq:.{de_p}f}*{math.sin(alpha_rad + beta_rad):.{de_p}f}*{math.sin(phi_rad - beta_rad):.{de_p}f} \
           + {eta:.{de_p}f} * {math.sin(alpha_rad):.{de_p}f}*{math.cos(phi_rad):.{de_p}f} ")
        print(f"   = {part3:.{de_p}f}")
    
    # 第四部分:  √[Kq sin(α-δ) sin(φ+δ) + ηsinαcosφ]
    inner_sqrt = Kq * math.sin(alpha_rad - delta_rad) * math.sin(phi_rad + delta_rad) + eta * math.sin(alpha_rad) * math.cos(phi_rad)
    part4 =  math.sqrt(inner_sqrt)
    if verbose:
        print(f"8. 计算大括号内第四部分:  √[Kq sin(α-δ) sin(φ+δ) + ηsinαcosφ]")
        print(f"   = √[{Kq:.{de_p}f}*sin({alpha}-{delta})*sin({phi}+{delta}) + {eta:.{de_p}f}*sin({alpha})*cos({phi})]")
        print(f"   = √[{Kq*math.sin(alpha_rad - delta_rad)*math.sin(phi_rad + delta_rad):.{de_p}f} + {eta*math.sin(alpha_rad)*math.cos(phi_rad):.{de_p}f}]")
        print(f"   = {part4:.{de_p}f}")
    
    # 计算大括号内的总和
    bracket_sum = part1 + part2 - part3 * part4
    if verbose:
        print(f"9. 计算大括号内总和 = 第一部分 + 第二部分 - 第三部分 * 第四部分")
        print(f"   = {part1:.{de_p}f} + {part2:.{de_p}f} - {part3:.{de_p}f} * {part4:.{de_p}f}")
        print(f"   = {bracket_sum:.{de_p}f}")
    
    # 计算最终的 Ka 值
    Ka = (Ka_numerator / Ka_denominator) * bracket_sum
    if verbose:
        print(f"10. 计算 Ka = [分子/分母] * 大括号总和")
        print(f"    = [{Ka_numerator:.{de_p}f}/{Ka_denominator:.{de_p}f}] * {bracket_sum:.{de_p}f}")
        print(f"    = {Ka_numerator/Ka_denominator:.{de_p}f} * {bracket_sum:.{de_p}f}")
        print(f"    = {Ka:.{de_p}f}")
        print(f"最终结果: 主动土压力系数 Ka = {Ka:.2f}")   
    return Ka

# 示例使用
if __name__ == "__main__":
    # 示例参数
    H = 4      # 挡土墙高度 (m)
    gamma = 19  # 土体重度 (kN/m³)
    c = 0      # 土的黏聚力 (kPa)
    phi = 30    # 土的内摩擦角 (°)
    q = 0       # 地表均布荷载标准值 (kN/m²)
    delta = 10 # 土对挡土墙墙背的摩擦角 (°)
    beta = 20  # 填土表面与水平面的夹角 (°)
    alpha = 60  # 支撑结构墙背与水平面的夹角 (°)
    verbose = 1
    de_p = 6
    print("主动土压力系数计算示例")
    print("="*50)   
    # 计算主动土压力系数
    Ka = cal_pressure_coefficient(H, gamma, c, phi, q, delta, beta, alpha,verbose)
# list = list(range(10,90,10))
# df = pd.DataFrame(list,columns=['alpha'])
# df['Ka'] = df['alpha'].apply(lambda x: cal_pressure_coefficient(H, gamma, c, phi, q, delta, beta, x,verbose))
# print(df)
