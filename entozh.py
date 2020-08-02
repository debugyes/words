import xlrd


def entozh():
    wb = xlrd.open_workbook("words.xlsx")
    sh1 = wb.sheet_by_name("词组")

    row = 1
    while row < sh1.nrows:
        title = sh1.cell_value(row, 0)
        print("词组:" + title)
        key = sh1.cell_value(row, 1)
        key = key.split('；')
        answer = input("答案:")
        if answer in key:
            print("正确")
            row += 1
        else:
            answer = input("错误,请再试一次：")
            while answer not in key:
                answer = input("错误,请再试一次：")
            print("正确")
            row += 1
    print("单词全部默写完啦！")
    print("\n\n")
    input("please input any key to exit!")


