import pandas as pd
import MaterialStatistic as ms
from pathlib import Path
father_direct = input('Please input the YJK models directory:')
father_direct = father_direct.rstrip('\\')
direct_output = father_direct + '/材料用量-统计.xlsx'
flag = 0 #0代表全楼总量，1代表上部总量
subdirs_names = [
    dir
    for dir in Path(father_direct).iterdir()
    if dir.is_dir()
]
subdirs_names.sort(
    key = lambda x: x.name.split('-')[-1]
)
df_sum0 = pd.DataFrame()
df_sum1 = pd.DataFrame()
df_sum_simpl = pd.DataFrame()
for l in subdirs_names:
    df_sum0_,df_sum1_,df_sum_simpl_ = ms.main_program(str(l))
    if flag == 0:
       df_sum0 = pd.concat([df_sum0,df_sum0_],axis=0)
    else:
       df_sum1 = pd.concat([df_sum1,df_sum1_],axis=0)
with pd.ExcelWriter(direct_output,engine='xlsxwriter') as writer:
    col_width = 6
    if flag == 0:
        ms.output(writer,df_sum0,'全楼总量',col_width=col_width)
    else:
        ms.output(writer,df_sum1,'上部总量',col_width=col_width)
