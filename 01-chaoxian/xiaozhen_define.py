def fraction_to_float(fraction):
    if isinstance(fraction,str) and '/' in fraction:
        numerator,denominator = fraction.split('/')
        try:
            return round(float(numerator) / float(denominator),7)
        except ZeroDivisionError:
            return 0.0
    return float(fraction)
#*****************************
#结构参数
stories = '3-25'
limit_dfe = '1/500'
limit_dfw = '1/500'
limit_dr_d,limit_dr_u = 1.2,1.5
limit_vmr = 0.016
limit_vc = 0.80
limit_stiff = 1.0
limit_md,limit_mu = 0.10,0.50
limit_v0 = 0.10
limit_dfe_f = fraction_to_float(limit_dfe)
limit_dfw_f = fraction_to_float(limit_dfw)
#*****************************
#路径及搜索参数
direct = input('please input the directory:')
direct_wdisp = direct + '\\wdisp.out'
direct_wzq = direct + '\\wzq.out'
direct_wmass = direct + '\\wmass.out'
direct_wv02q = direct + '\\wv02q.out'
s_structure = ''
d_index = 3
e_index = 3
w_index = 4
ratiov_index = 4
m_index = 3
v0_index = 7
f_ratio = '规定'
indicators_disp = {
    'ex_df':'X 方向地震作用下的楼层最大位移',
    'ey_df':'Y 方向地震作用下的楼层最大位移' ,
    'wx_df':'+X 方向风荷载作用下的楼层最大位移',
    'wy_df':'+Y 方向风荷载作用下的楼层最大位移',
    'ex_ds':'X 方向地震作用下的楼层最大位移',
    'ey_ds':'Y 方向地震作用下的楼层最大位移',
    'wx_ds':'+X 方向风荷载作用下的楼层最大位移',
    'wy_ds':'+Y 方向风荷载作用下的楼层最大位移',
    'ex+_dr':'X+ 偶然偏心规定水平力作用下的楼层最大位移',
    'ex-_dr':'X- 偶然偏心规定水平力作用下的楼层最大位移',
    'ey+_dr':'Y+ 偶然偏心规定水平力作用下的楼层最大位移',
    'ey-_dr':'Y- 偶然偏心规定水平力作用下的楼层最大位移',
           }
indicators_force_e = {
    'ex':'各层 X 方向的作用力(CQC)',
    'ey':'各层 Y 方向的作用力(CQC)',
}
indicator_force_w = '                           风荷载信息'
indicator_ratio_vc = '楼层抗剪承载力验算'
indicator_ratio_s = '各层刚心、偏心率、相邻层侧移刚度比等计算信息'
indicator_ratio_m = '规定水平力下框架柱、短肢墙地震倾覆力矩百分比'
indicator_ratio_v0 = '框架柱地震剪力及百分比'
endflag_disp = '=='
endflag_force_e = '='
endflag_force_w = '各楼层等效尺寸'
endflag_ratio_v = '**'
endflag_ratio_s = '**'
endflag_ratio_m = '**'
endflag_ratio_v0 = '**'
#*****************************
#输出excel参数
df_list = ['ex_v','ey_v','ex_m','ey_m','wx_v','wy_v','wx_m','wy_m', \
    'ex_df','ey_df','ex_ds','ey_ds','wx_df','wy_df','wx_ds','wy_ds']
id_list = ['ex+_dr','ex-_dr','ey+_dr','ey-_dr','ex_vmr','ey_vmr',\
    'r_vcx','r_vcy','r_sx','r_sy','r_mx','r_my','r_vx','r_vy']
head_df_list = {
    'ex_v':'层剪力-地震X',
    'ey_v':'层剪力-地震Y',
    'ex_m':'层弯矩-地震X',
    'ey_m':'层弯矩-地震Y',
    'wx_v':'层剪力-风X',
    'wy_v':'层剪力-风Y',
    'wx_m':'层弯矩-风X',
    'wy_m':'层弯矩-风Y',
    'ex_df':'位移角-地震X',
    'ey_df':'位移角-地震Y',
    'ex_ds':'位移-地震X',
    'ey_ds':'位移-地震Y',
    'wx_df':'位移角-风X',
    'wy_df':'位移角-风Y',
    'wx_ds':'位移-风X',
    'wy_ds':'位移-风Y',
}
head_id_list = {
    'ex+_dr':'位移比-X+',
    'ex-_dr':'位移比-X-',
    'ey+_dr':'位移比-Y+',
    'ey-_dr':'位移比-Y-',
    'ex_vmr':'剪重比-X',
    'ey_vmr':'剪重比-Y',
    'r_vcx':'抗剪承载力比-X',
    'r_vcy':'抗剪承载力比-Y',
    'r_sx':'侧向刚度比-X',
    'r_sy':'侧向刚度比-Y',
    'r_mx':'框架弯矩与总倾覆力矩比-X',
    'r_my':'框架弯矩与总倾覆力矩比-Y',
    'r_vx':'框架柱地震剪力比-X',
    'r_vy':'框架柱地震剪力比-Y',
}