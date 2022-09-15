import re
test ='Gonzalez(1990,1993)'
date_all = re.findall(r"(\d{4})",test)
print(date_all)