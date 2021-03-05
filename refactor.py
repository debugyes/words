# 重构
import xlrd
import time
import random
import datetime


# 词组英译中模式
def groupentozh():
    # 初始化目标工作簿
    wb = xlrd.open_workbook("group.xlsx")
    # 读出要默写哪一份
    file = open("group_config_entozh.txt")
    piece = file.read()
    file.close()
    # 开始默写
    sh1 = wb.sheet_by_name(piece)

    print("1.正序\n2.随机\n3.返回\n")
    choose = input("请选择:")
    if choose == "1":
        indexs = list(range(0, sh1.nrows))  # 生成顺序数列
        groupentozh_function(indexs)
    elif choose == "2":
        groupentozh_function(random.sample(range(0, sh1.nrows), sh1.nrows))  # 生成随机数列
    else:
        main()  # 返回主菜单

    print("单词全部默写完啦！")
    print("\n\n")

    piece = int(piece)

    # 打开配置文件
    file = open("group_config_entozh.txt", "w")
    piece = int(piece)
    # 如果配置文件序号越界置1
    if piece < len(wb.sheets()):
        piece = piece + 1
    else:
        piece = 1
    file.write('%d' % piece)
    file.close()

    time.sleep(1)
    main()


# 词组英译中功能实现
def groupentozh_function(indexes):
    # 初始化目标工作簿
    wb = xlrd.open_workbook("group.xlsx")
    # 读出要默写哪一份
    file = open("group_config_entozh.txt")
    piece = file.read()
    file.close()
    # 开始默写
    sh1 = wb.sheet_by_name(piece)

    for row in indexes:
        title = sh1.cell_value(row, 0)
        print("(" + str(row + 1) + '/' + str(sh1.nrows) + ")" + "词组:" + title)
        key = sh1.cell_value(row, 1)
        key = key.split('；')
        answer = input("答案:")
        while answer == '\n':
            answer = input("答案:")
        if answer.strip() in key:
            print("正确")
        else:
            answer = input("错误,请再试一次：")
            while answer.strip() not in key:
                answer = input("错误,请再试一次：")
            print("正确")


# 词组中译英模式
def groupzhtoen():
    # 初始化目标工作簿
    wb = xlrd.open_workbook("group.xlsx")
    # 读出要默写哪一份
    file = open("group_config_zhtoen.txt")
    piece = file.read()
    file.close()
    # 开始默写
    sh1 = wb.sheet_by_name(piece)

    print("1.正序\n2.随机\n3.返回\n")
    choose = input("请选择:")
    if choose == "1":
        indexs = list(range(0, sh1.nrows))
        groupzhtoen_funtion(indexs)
    elif choose == "2":
        groupzhtoen_funtion(random.sample(range(0, sh1.nrows), sh1.nrows))
    else:
        main()
    print("单词全部默写完啦！")
    print("\n\n")

    # 打开配置文件
    file = open("group_config_zhtoen.txt", "w")
    piece = int(piece)
    # 如果配置文件序号越界置1
    if piece < len(wb.sheets()):
        piece = piece + 1
    else:
        piece = 1
    file.write('%d' % piece)
    file.close()

    time.sleep(1)
    main()


# 词组中译英功能实现
def groupzhtoen_funtion(indexes):
    # 初始化目标工作簿
    wb = xlrd.open_workbook("group.xlsx")
    # 读出要默写哪一份
    file = open("group_config_zhtoen.txt")
    piece = file.read()
    file.close()
    # 开始默写
    sh1 = wb.sheet_by_name(piece)

    for row in indexes:
        title = sh1.cell_value(row, 1)
        print("(" + str(row + 1) + '/' + str(sh1.nrows) + ")" + "词组:" + title)
        key = sh1.cell_value(row, 0)
        key = key.split('；')
        answer = input("答案:")
        while answer == '\n':
            answer = input("答案:")
        if answer.strip() in key:
            print("正确")
        else:
            answer = input("错误,请再试一次：")
            while answer.strip() not in key:
                answer = input("错误,请再试一次：")
            print("正确")


# 单词中译英模式
def wordzhtoen():
    # 初始化目标工作簿
    wb = xlrd.open_workbook("word.xlsx")
    # 读出要默写哪一份
    file = open("word_config_zhtoen.txt")
    piece = file.read()
    file.close()
    # 开始默写
    sh1 = wb.sheet_by_name(piece)

    print("1.正序\n2.随机\n3.返回\n")
    choose = input("请选择:")
    if choose == "1":
        indexs = list(range(0, sh1.nrows))  # 生成顺序数列
        wordzhtoen_function(indexs)
    elif choose == "2":
        wordzhtoen_function(random.sample(range(0, sh1.nrows), sh1.nrows))  # 生成随机数列
    else:
        main()  # 返回主菜单

    print("单词全部默写完啦！")
    print("\n\n")

    piece = int(piece)

    # 打开配置文件
    file = open("word_config_zhtoen.txt", "w")
    piece = int(piece)
    # 如果配置文件序号越界置1
    if piece < len(wb.sheets()):
        piece = piece + 1
    else:
        piece = 1
    file.write('%d' % piece)
    file.close()

    time.sleep(1)
    main()
    # 完成


# 单词中译英功能实现
def wordzhtoen_function(indexes):
    # 初始化目标工作簿
    wb = xlrd.open_workbook("word.xlsx")
    # 读出要默写哪一份
    file = open("word_config_zhtoen.txt")
    piece = file.read()
    file.close()
    # 开始默写
    sh1 = wb.sheet_by_name(piece)

    for row in indexes:
        title = sh1.cell_value(row, 1)
        print("(" + str(row + 1) + '/' + str(sh1.nrows) + ")" + "词组:" + title)
        key = sh1.cell_value(row, 0)
        key = key.split('；')
        answer = input("答案:")
        while answer == '\n':
            answer = input("答案:")
        if answer.strip() in key:
            print("正确")
        else:
            answer = input("错误,请再试一次：")
            while answer.strip() not in key:
                answer = input("错误,请再试一次：")
            print("正确")
    # 完成


# 单词英译中模式
def wordentozh():
    # 初始化目标工作簿
    wb = xlrd.open_workbook("word.xlsx")
    # 读出要默写哪一份
    file = open("word_config_entozh.txt")
    piece = file.read()
    file.close()
    # 开始默写
    sh1 = wb.sheet_by_name(piece)

    print("1.正序\n2.随机\n3.返回\n")
    choose = input("请选择:")
    if choose == "1":
        indexs = list(range(0, sh1.nrows))  # 生成顺序数列
        wordentozh_function(indexs)
    elif choose == "2":
        wordentozh_function(random.sample(range(0, sh1.nrows), sh1.nrows))  # 生成随机数列
    else:
        main()  # 返回主菜单

    print("单词全部默写完啦！")
    print("\n\n")

    piece = int(piece)

    # 打开配置文件
    file = open("word_config_entozh.txt", "w")
    piece = int(piece)
    # 如果配置文件序号越界置1
    if piece < len(wb.sheets()):
        piece = piece + 1
    else:
        piece = 1
    file.write('%d' % piece)
    file.close()

    time.sleep(1)
    main()


# 单词英译中模式实现
def wordentozh_function(indexes):
    # 初始化目标工作簿
    wb = xlrd.open_workbook("word.xlsx")
    # 读出要默写哪一份
    file = open("word_config_zhtoen.txt")
    piece = file.read()
    file.close()
    # 开始默写
    sh1 = wb.sheet_by_name(piece)

    for row in indexes:
        title = sh1.cell_value(row, 1)
        print("(" + str(row + 1) + '/' + str(sh1.nrows) + ")" + "词组:" + title)
        key = sh1.cell_value(row, 0)
        key = key.split('；')
        answer = input("答案:")
        while answer == '\n':
            answer = input("答案:")
        if answer.strip() in key:
            print("正确")
        else:
            answer = input("错误,请再试一次：")
            while answer.strip() not in key:
                answer = input("错误,请再试一次：")
            print("正确")


# 打卡函数
def checkin():
    now = datetime.datetime.now()
    file = open("checkin.txt", "a+")
    # file.write(str(now) + '\n')
    if file.write(str(now) + '\n') != 0:
        print("打卡成功")
    else:
        print("打卡失败")


# 学习模式
def viewmode():
    # 初始化目标工作簿
    wb = xlrd.open_workbook("group.xlsx")
    # 读出要默写哪一份
    file = open("group_config_entozh.txt")
    piece = file.read()
    file.close()
    # 开始默写
    sh1 = wb.sheet_by_name(piece)

    row = 0
    while row < sh1.nrows:
        title = sh1.cell_value(row, 0)
        print("(" + str(row + 1) + '/' + str(sh1.nrows) + ")" + "词组:" + title + '\t' + sh1.cell_value(row, 1) + '\n')
        row += 1
        time.sleep(3)


# 主界面函数
def main():
    print("+-----------------------+")
    print("|       Welcome         |")
    print("+-----------------------+")
    print("|1.词组英译中\n|2.词组中译英\n|3.单词英译中\n|4.单词中译英\n|5.打卡\n|6.背诵模式\n|7.退出程序")
    print("+-----------------------+")

    choose = input("请选择：")
    if choose == "1":
        groupentozh()
    elif choose == "2":
        groupzhtoen()
    elif choose == "3":
        wordentozh()
    elif choose == "4":
        wordzhtoen()
    elif choose == "5":
        checkin()
    elif choose == "6":
        viewmode()
    else:
        exit()


if __name__ == '__main__':
    main()


