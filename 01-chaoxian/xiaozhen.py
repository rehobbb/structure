import re
direct = input('please input the directory:')
with open(direct + '\wdisp.out','r') as file:
    lines = file.readlines()
str1 = 'X 方向地震作用下的楼层最大位移'
str2 = '=='
results = []
find_target = False
i = 0
while i < len(lines):
    line = lines[i].strip()
    if str1 in line:
        find_target = True
        i += 3
        continue
    if find_target:
         if line and line[0].isdigit():
            parts_1 = lines[i].split()
            if int(parts_1[0]) < 100:
                floor = int(parts_1[0])
                parts_2 = re.split(r'\s{2,}',lines[i+1].strip())
                disp = parts_2[3]
                if '/ ' in disp:
                    disp = disp.replace('/ ','/')
                results.append([floor,disp])
                i += 1
                continue
    if str2 in line:
        break
    i += 1
for floor,disp in results:
    print(floor,disp)


