import xlrd
import xlwt
from xlutils.copy import copy
import datetime
import time
import random

# print(sh1.cell_value(0, 1))


'''
title = sh1.cell_value(1, 0)
print("词组:" + title)
key = sh1.cell_value(1, 1)
answer = input("答案:")
if key == answer:
    print("正确")
'''


# 英译中模式
def entozh():
    # 初始化目标工作簿
    wb = xlrd.open_workbook("words.xlsx")
    sh1 = wb.sheet_by_name("词组")
    # 创建错题数组
    wrong = []

    print("1.正序\n2.随机")
    choose = input("请选择:")
    if choose == "1":
        row = 1
        while row < sh1.nrows:
            title = sh1.cell_value(row, 0)
            print("词组:" + title)
            key = sh1.cell_value(row, 1)
            key = key.split('；')
            answer = input("答案:")
            if answer.strip() in key:
                print("正确")
                row += 1
            else:
                wrong.append(row)
                answer = input("错误,请再试一次：")
                while answer.strip() not in key:
                    answer = input("错误,请再试一次：")
                print("正确")
                row += 1
        if len(wrong) != 0:
            # 错题订正
            count = 1
            while count <= len(wrong):
                rows = wrong[count]
                title = sh1.cell_value(rows, 1)
                print("词组:" + title)
                keys = sh1.cell_value(rows, 0)
                answer = input("答案:")
                if keys == answer:
                    print("正确")
                    rows += 1
                else:
                    wrong.append(rows)
                    answer = input("错误,请再试一次：")
                    while answer != keys:
                        answer = input("错误,请再试一次：")
                    print("正确")
                    rows += 1

    else:
        serial = []
        count = 0
        # 生成随机题号数组
        while count < sh1.nrows:
            serial.append(random.randint(1, sh1.nrows))
            count = count + 1

        count = 0
        while count < sh1.nrows:
            # 把随机题号读取到row中
            row = serial[count]
            title = sh1.cell_value(row, 0)
            print("词组:" + title)
            key = sh1.cell_value(row, 1)
            key = key.split('；')
            answer = input("答案:")
            if answer in key:
                print("正确")
                count += 1
            else:
                wrong.append(row)
                answer = input("错误,请再试一次：")
                while answer not in key:
                    answer = input("错误,请再试一次：")
                print("正确")
                count += 1
        if len(wrong) != 0:
            # 错题订正
            count = 1
            while count <= len(wrong):
                rows = wrong[count]
                title = sh1.cell_value(rows, 1)
                print("词组:" + title)
                keys = sh1.cell_value(rows, 0)
                answer = input("答案:")
                if keys == answer:
                    print("正确")
                    rows += 1
                else:
                    wrong.append(rows)
                    answer = input("错误,请再试一次：")
                    while answer != keys:
                        answer = input("错误,请再试一次：")
                    print("正确")
                    rows += 1

    print("单词全部默写完啦！")
    print("\n\n")
    time.sleep(1)
    main()

def zhtoen():
    wb = xlrd.open_workbook("words.xlsx")
    sh1 = wb.sheet_by_name("词组")
    # 创建错题数组
    wrong = []

    print("1.正序\n2.随机")
    choose = input("请选择:")
    if choose == "1":
        rows = 1
        while rows < sh1.nrows:
            title = sh1.cell_value(rows, 1)
            print("词组:" + title)
            keys = sh1.cell_value(rows, 0)
            answer = input("答案:")
            if keys == answer:
                print("正确")
                rows += 1
            else:
                wrong.append(rows)
                answer = input("错误,请再试一次：")
                while answer != keys:
                    answer = input("错误,请再试一次：")
                print("正确")
                rows += 1
        if len(wrong) != 0:
            # 错题订正
            count = 1
            while count <= len(wrong):
                rows = wrong[count]
                title = sh1.cell_value(rows, 1)
                print("词组:" + title)
                keys = sh1.cell_value(rows, 0)
                answer = input("答案:")
                if keys == answer:
                    print("正确")
                    rows += 1
                else:
                    wrong.append(rows)
                    answer = input("错误,请再试一次：")
                    while answer != keys:
                        answer = input("错误,请再试一次：")
                    print("正确")
                    rows += 1


    else:
        serial = []
        count = 0
        while count < sh1.nrows:
            serial.append(random.randint(1, sh1.nrows))
            count = count + 1

        count = 0
        while count < sh1.nrows:
            row = serial[count]
            title = sh1.cell_value(row, 1)
            print("词组:" + title)
            keys = sh1.cell_value(row, 0)
            answer = input("答案:")
            if keys == answer:
                print("正确")
                count += 1
            else:
                wrong.append(title)
                answer = input("错误,请再试一次：")
                while answer != keys:
                    answer = input("错误,请再试一次：")
                print("正确")
                count += 1

        if len(wrong) != 0:
            # 错题订正
            count = 1
            while count <= len(wrong):
                rows = wrong[count]
                title = sh1.cell_value(rows, 1)
                print("词组:" + title)
                keys = sh1.cell_value(rows, 0)
                answer = input("答案:")
                if keys == answer:
                    print("正确")
                    rows += 1
                else:
                    wrong.append(rows)
                    answer = input("错误,请再试一次：")
                    while answer != keys:
                        answer = input("错误,请再试一次：")
                    print("正确")
                    rows += 1

    print("单词全部默写完啦！")
    print("\n\n")
    time.sleep(1)
    main()


# 打卡函数有bug，待修改

def signin():
    # rexcel = xlrd.open_workbook("words.xlsx")
    oldWb = xlrd.open_workbook("words.xlsx")
    newWb = copy(oldWb)
    sh1 = newWb.sheet_by_name("打卡")

    row = 0
    col = 0
    value = sh1.cell_value(row, col)
    while value != "null":
        row += 1
    sh1.write(row, col, datetime.time.today())
    print("打卡成功！")
    newWb.save("words.xlsx")


# 主界面函数
def main():
    """
    print("+----------------------+")
    print("|\t\t欢迎使用\t\t   |")
    print("|1.英译中\t\t\t   |\n|2.中译英\t\t\t   |\n|3.打卡(bug)\t\t   |\n|4.混合模式(先挖个坑)   |")
    print("+----------------------+")
    :return:
    """
    # print("+----------------------+")
    print("欢迎使用")
    print("1.英译中\n2.中译英\n3.打卡(bug)\n4.混合模式(先挖个坑)")
    # print("+----------------------+")
    choose = input("请选择：")
    if choose == "1":
        entozh()
    elif choose == "2":
        zhtoen()
    elif choose == '3':
        signin()


# 调用部分
main()
