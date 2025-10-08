
class AppConfig:
    #1. 转换分数为浮点数
    @staticmethod
    def fraction_to_float(fraction):
        if isinstance(fraction,str) and '/' in fraction:
            numerator,denominator = fraction.split('/')
            try:
                return round(float(numerator) / float(denominator),7)
            except ZeroDivisionError:
                return 0.0
        return float(fraction)   
    def __init__(self,direct):
    #2. 结构参数
        self.stories = '3-25'
        self.limit_df = {
            'limit_dfe':'1/500',
            'limit_dfw':'1/500',
        }
        self.limit_id = {
            'limit_drd':1.2,
            'limit_dru':1.5,
            'limit_vmr':0.016,
            'limit_vc':0.80,
            'limit_stf':1.0,
            'limit_md':0.10,
            'limit_mu':0.50,
            'limit_v0d':0.10,
            'limit_v0u':0.20,
        }
        self.limit_df_float = {
            'limit_dfe':AppConfig.fraction_to_float(self.limit_df['limit_dfe']),
            'limit_dfw':AppConfig.fraction_to_float(self.limit_df['limit_dfw']),    
        }
    #2. 路径参数
        self.direct = direct
        self.direct_wmass = self.direct + '/设计结果/wmass.out'  
        self.direct_wzq = self.direct + '/设计结果/wzq.out'              
        self.direct_wdisp = self.direct + '/设计结果/wdisp.out'
        self.direct_wv02q = self.direct + '/设计结果/wv02q.out'
    #3. 搜索参数
    #3.1 结构体系
        self.s_structure = ''
    #3.2 定位参数
        self.w_index = 4
        self.ratiov_index = 4
        self.d_index = 3
        self.e_index = 3
        self.m_index = 3
        self.v0_index = 7
    #3.3 数据开头标识
        self.indicator_wforce = '                           风荷载信息'
        self.indicator_ratios = '各层刚心、偏心率、相邻层侧移刚度比等计算信息'
        self.indicator_ratiovc = '楼层抗剪承载力验算'   
        self.indicators_eforce = {
            'ex':'各层 X 方向的作用力(CQC)',
            'ey':'各层 Y 方向的作用力(CQC)',
        }           
        self.f_ratio = '规定'     
        self.indicators_disp = {
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
        self.indicator_ratiom = '规定水平力下框架柱、短肢墙地震倾覆力矩百分比'
        self.indicator_ratiov0 = '框架柱地震剪力及百分比'
    #3.4 数据结尾标识
        self.endflag_wforce = '各楼层等效尺寸'
        self.endflag_ratios = '**'            
        self.endflag_ratiovc = '**'
        self.endflag_eforce = '='        
        self.endflag_disp = '=='
        self.endflag_ratiom = '**'
        self.endflag_ratiov0 = '**'
    #4. 输出excel参数
    #4.1 内力位移数据头
        self.head_df = {
            'ex_v':'层剪力-EX',
            'ey_v':'层剪力-EY',
            'ex_m':'层弯矩-EX',
            'ey_m':'层弯矩-EY',
            'wx_v':'层剪力-WX',
            'wy_v':'层剪力-WY',
            'wx_m':'层弯矩-WX',
            'wy_m':'层弯矩-WY',
            'ex_df':'位移角-EX',
            'ey_df':'位移角-EY',
            'ex_ds':'位移-EX',
            'ey_ds':'位移-EY',
            'wx_df':'位移角-WX',
            'wy_df':'位移角-WY',
            'wx_ds':'位移-WX',
            'wy_ds':'位移-WY',
        }
    #4.2 整体参数数据头
        self.head_id = {
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
            'r_mx':'框架倾覆力矩比-X',
            'r_my':'框架倾覆力矩比-Y',
            'r_vx':'框架剪力比-X',
            'r_vy':'框架剪力比-Y',
        }
    #5. 绘制图表参数
    #5.1 内力位移数据标题
        self.plot_df = {
            'ex_v':['楼层剪力(kN)',''],
            'ey_v':['楼层剪力(kN)',''],
            'ex_m':['倾覆弯矩(kN.m)',''],
            'ey_m':['倾覆弯矩(kN.m)',''],
            'wx_v':['楼层剪力(kN)',''],
            'wy_v':['楼层剪力(kN)',''],
            'wx_m':['倾覆弯矩(kN.m)',''],
            'wy_m':['倾覆弯矩(kN.m)',''],
            'ex_df':['位移角','limit_dfe'],
            'ey_df':['位移角','limit_dfw'],
            'ex_ds':['位移(mm)',''],
            'ey_ds':['位移(mm)',''],
            'wx_df':['位移角','limit_dfw'],
            'wy_df':['位移角','limit_dfw'],
            'wx_ds':['位移(mm)',''],
            'wy_ds':['位移(mm)',''],
        }
    #5.2 整体参数数据标题
        self.plot_id = {
            'ex+_dr':['扭转位移比','limit_dr'],
            'ex-_dr':['扭转位移比','limit_dr'],
            'ey+_dr':['扭转位移比','limit_dr'],
            'ey-_dr':['扭转位移比','limit_dr'],
            'ex_vmr':['剪重比','limit_vmr'],
            'ey_vmr':['剪重比','limit_vmr'],
            'r_vcx':['抗剪承载力比','limit_vc'],
            'r_vcy':['抗剪承载力比','limit_vc'],
            'r_sx':['侧向刚度比','limit_stf'],
            'r_sy':['侧向刚度比','limit_stf'],
            'r_mx':['框架倾覆力矩比','limit_m'],
            'r_my':['框架倾覆力矩比','limit_m'],
            'r_vx':['框架剪力比','limit_v0'],
            'r_vy':['框架剪力比','limit_v0'],
        }
