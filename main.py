# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

import openpyxl
import datetime

# 一天工作多少秒
TOTAL_WORK_SECOND = 7.5 * 60 * 60
LAUNCH_BREAK_SECOND = 1.5 * 60 * 60


class workItem:
    def __init__(self, name, startTime, endTime):
        self.name = name
        self.startTime = startTime
        self.endTime = endTime


def convertToFloat(startTime, endTime):
    timeOne = datetime.datetime.strptime(startTime, '%H:%M')
    timeTwo = datetime.datetime.strptime(endTime, '%H:%M')
    timeDiffSec = (timeTwo - timeOne).total_seconds()
    if timeOne.hour < 12:
        # 减去午休的时间
        timeDiffSec = timeDiffSec - LAUNCH_BREAK_SECOND

    if timeDiffSec > TOTAL_WORK_SECOND:
        print("加班一天")
    else:
        print("加班半天")


def main():
    sourceFile = "/Users/weigan/Downloads/7月考勤全员.xlsx"
    workbook = openpyxl.load_workbook(sourceFile, read_only=True)
    sheet = workbook.worksheets[0]
    print(sheet.max_row)

    cellOne = sheet.cell(row=5, column=9).value
    cellTwo = sheet.cell(row=6, column=9).value
    convertToFloat(str(cellOne), str(cellTwo))


if __name__ == '__main__':
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
