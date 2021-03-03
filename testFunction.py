import xlrd
import time
import random
import datetime

# 初始化目标工作簿
wb = xlrd.open_workbook("words.xlsx")
# 读出要默写哪一份
file = open("config_entozh.txt")
piece = file.read()
'''
file.close()
# 开始默写
# sh1 = wb.sheet_by_name(piece)
'''
piece = int(piece)

# 打开配置文件
file = open("config_entozh.txt", "w")
# 如果配置文件序号越界置1
if piece < len(wb.sheets()):
	piece = piece+1
else:
	piece = 1
file.write('%d' % piece)
print(piece)
file.close()