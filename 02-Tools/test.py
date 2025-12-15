import re
str='  砼等级 	个数		C35		C60'
s = re.search('砼等级',str).group()
print(s)
