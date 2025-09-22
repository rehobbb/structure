import pandas as pd
test = {
    'fl':[1,2,3],
    '位移':{
        'YJK_X':[1,2,3],
        'MIDAS_X':[4,5,6],
        'YJK_Y':[7,8,9],
        'MIDAS_Y':[10,11,12]
    },
    '剪力':{
        'YJK_X':[1,2,3],
        'MIDAS_X':[4,5,6],
        'YJK_Y':[7,8,9],
        'MIDAS_Y':[10,11,12]
    }
}
df = pd.DataFrame(test)
df.to_excel('test.xlsx',index=False)