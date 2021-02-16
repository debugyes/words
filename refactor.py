import xlrd
import time
import random
import datetime


# 英译中模式
def entozh():
    # 初始化目标工作簿
    wb = xlrd.open_workbook("words.xlsx")
    # 读出要默写哪一份
    file = open("config_entozh.txt")
    piece = file.read()
    file.close()
    # 开始默写
    sh1 = wb.sheet_by_name(piece)

    print("1.正序\n2.随机\n3.返回\n")
    choose = input("请选择:")
    if choose == "1":
        indexs = list(range(0, sh1.nrows))
        entozh_function(indexs)
    elif choose == "2":
        entozh_function(random.sample(range(0, sh1.nrows), sh1.nrows))
    else:
        main()
    print("单词全部默写完啦！")
    print("\n\n")

    # 打开配置文件
    file = open("config_entozh.txt", "w")
    # 如果配置文件序号越界置1
    if int(piece) + 1 <= len(wb.sheets()):
        piece = str(int(piece) + 1)
    else:
        piece = 1
    file.write(piece)

    time.sleep(1)
    main()


def entozh_function(indexes):
    # 初始化目标工作簿
    wb = xlrd.open_workbook("words.xlsx")
    # 读出要默写哪一份
    file = open("config_entozh.txt")
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


# 中译英模式
def zhtoen():
    # 初始化目标工作簿
    wb = xlrd.open_workbook("words.xlsx")
    # 读出要默写哪一份
    file = open("config_entozh.txt")
    piece = file.read()
    file.close()
    # 开始默写
    sh1 = wb.sheet_by_name(piece)

    print("1.正序\n2.随机\n3.返回\n")
    choose = input("请选择:")
    if choose == "1":
        indexs = list(range(0, sh1.nrows))
        zhtoen_funtion(indexs)
    elif choose == "2":
        zhtoen_funtion(random.sample(range(0, sh1.nrows), sh1.nrows))
    else:
        main()
    print("单词全部默写完啦！")
    print("\n\n")
    # 打开配置文件
    file = open("config_zhtoen.txt", "w")
    # 如果配置文件序号越界置1
    if int(piece) + 1 <= len(wb.sheets()):
        piece = str(int(piece) + 1)
    else:
        piece = 1
    file.write(str(piece))

    time.sleep(1)
    main()


def zhtoen_funtion(indexes):
    # 初始化目标工作簿
    wb = xlrd.open_workbook("words.xlsx")
    # 读出要默写哪一份
    file = open("config_entozh.txt")
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



def clockin():
    now = datetime.datetime.now()
    file = open("clockin.ext", "a+")
    file.write(str(now) + '\n')


def viewmode():
    # 初始化目标工作簿
    wb = xlrd.open_workbook("words.xlsx")
    # 读出要默写哪一份
    file = open("config_entozh.txt")
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
    print("|1.英译中\n|2.中译英\n|3.打卡\n|4.背诵模式\n|6.退出程序")
    print("+-----------------------+")

    choose = input("请选择：")
    if choose == "1":
        entozh()
    elif choose == "2":
        zhtoen()
    elif choose == "3":
        clockin()
    elif choose == "4":
        viewmode()
    else:
        exit()


if __name__ == '__main__':
    main()
