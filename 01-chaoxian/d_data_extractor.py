import re 
class DataExtractor:
    def __init__(self,config):
        self.config = config
    #1. 提取wmass数据-风内力，刚度比，抗剪承载力比
    #1.1 提取风内力
    def extract_wforce(self,chunk,data):
        data.setdefault('fl',[])
        data.setdefault('wx_v',[])
        data.setdefault('wy_v',[])
        data.setdefault('wx_m',[])
        data.setdefault('wy_m',[])
        index = self.config.w_index
        i = 0
        while i < len(chunk):
            parts = chunk[i].split()
            if parts[0].isdigit():
                data['fl'].append(int(parts[0])) 
                parts1 = chunk[i+1].split()
                value_vx = float(parts[index])
                value_mx = float(parts[index+1])
                value_vy = float(parts1[index-2])
                value_my = float(parts1[index-1])
                if not data['fl']:
                    data['fl'].append(floor) 
                data['wx_v'].append(round(value_vx,0))
                data['wx_m'].append(round(value_mx,0))
                data['wy_v'].append(round(value_vy,0))
                data['wy_m'].append(round(value_my,0))                  
                i += 1
                continue
            i += 1    
    #1.2 提取刚度比比
    def extract_ratios(self,chunk,data,stru):
        data.setdefault('r_sx',[])
        data.setdefault('r_sy',[])
        if '剪' in stru:
            rs_x = 'Ratx2'
            rs_y = 'Raty2'
        else:
            rs_x = 'Ratx1'
            rs_y = 'Raty1'
        for list in chunk:
            floor_match = re.search(r'Floor No\.\s*(\d+)',list)
            if floor_match:
                floor = int(floor_match.group(1))
                continue
            rs_match = re.search(rs_x+r'=\s*([\d.]+)\s*'+rs_y+r'=\s*([\d.]+)',list)
            if rs_match and floor is not None:
                ratio_x = float(rs_match.group(1))
                ratio_y = float(rs_match.group(2))
                data['r_sx'].append(round(ratio_x,2))
                data['r_sy'].append(round(ratio_y,2))
                floor = None   
    #1.3 提取抗剪承载力比
    def extract_ratiovc(self,chunk,data):
        data.setdefault('r_vcx',[])
        data.setdefault('r_vcy',[])
        index = self.config.ratiov_index
        i =0 
        while i < len(chunk):
            parts = chunk[i].split()
            if parts[0].isdigit():
                if not data['fl']:
                    data['fl'].append(int(parts1[0])) 
                ratio_vx = float(parts[index])
                ratio_vy = float(parts[index+1])
                data['r_vcx'].append(round(ratio_vx,2))
                data['r_vcy'].append(round(ratio_vy,2))                  
                i += 1
                continue
            i += 1  
    #2. 提取wzq数据-地震内力及剪重比
    def extract_eforce(self,chunk,data,key):
        data.setdefault(key+'_v',[])
        data.setdefault(key+'_vmr',[])
        data.setdefault(key+'_m',[])
        index = self.config.e_index       
        i = 0 
        v_pattern = r'(\d+\.?\d*)\(\s*(\d+\.?\d*)%\)'
        while i < len(chunk):
            parts = re.split(r'\s{2,}',chunk[i].strip())
            if parts[0].isdigit():
                value_v = float(re.search(v_pattern,parts[index]).group(1))
                value_vmr = float(re.search(v_pattern,parts[index]).group(2))/100
                value_m = float(parts[index+1])
                data[key+'_v'].append(round(value_v,0))
                data[key+'_vmr'].append(round(value_vmr,4))
                data[key+'_m'].append(round(value_m,0))                  
                i += 1
                continue
            i += 1                                      
    #3. 提取wdisp数据-地震及风位移、位移角
    def extract_disp(self,chunk,data,key,ratio):
        data.setdefault(key,[])
        index = self.config.d_index
        ratio = self.config.f_ratio
        i = 0 
        while i < len(chunk):
            parts1 = chunk[i].split()
            if parts1[0].isdigit() and int(parts1[0])<100:
                if self.is_contained(ratio,chunk):
                    parts2 = chunk[i+1].split()
                    value = max(float(parts1[index+2]),float(parts2[index]))
                else:
                    if 'ds' in key:
                        value = round(float(parts1[index]),2)
                    else:
                        parts2 = re.split(r'\s{2,}',chunk[i+1].strip())
                        if 'w' in key:
                            value = parts2[index+1].replace('/ ','/')
                        else:
                            value = parts2[index].replace('/ ','/')
                data[key].append(value)      
                i += 1
                continue
            i += 1
    #4. 提取wv02q数据-框架倾覆弯矩比，剪力占比
    #4.1 框架倾覆弯矩比
    def extract_ratiom(self,chunk,data):
        data.setdefault('r_mx',[])
        data.setdefault('r_my',[])
        index = self.config.m_index
        i = 0 
        while i < len(chunk):
            parts = chunk[i].split()
            if parts[0].isdigit():
                if parts[index-1] == 'X':
                    value_mx = float(parts[index].strip('%'))/100
                    data['r_mx'].append(round(value_mx,2))
                else:
                    value_my = float(parts[index].strip('%'))/100
                    data['r_my'].append(round(value_my,2))
                i += 1
                continue
            i += 1
    #4.2 框架剪力占比
    def extract_ratiov0(self,chunk,data):
        data.setdefault('r_vx',[])
        data.setdefault('r_vy',[])
        index = self.config.v0_index
        i = 0 
        while i < len(chunk):
            parts = chunk[i].split()
            if parts[0].isdigit():
                if parts[index-5] == 'X':
                    value_v0x = float(parts[index].strip('%'))/100
                    data['r_vx'].append(round(value_v0x,2))
                else:
                    value_v0y = float(parts[index].strip('%'))/100
                    data['r_vy'].append(round(value_v0y,2))
                i += 1
                continue
            i += 1
    @staticmethod
    def is_contained(target,string_list):
        return any(target in s for s in string_list)            