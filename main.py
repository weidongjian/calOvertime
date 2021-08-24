# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

import openpyxl
import datetime

# 一天工作多少秒
TOTAL_WORK_SECOND = 7.5 * 60 * 60
LAUNCH_BREAK_SECOND = 1.5 * 60 * 60

COLUMN_WEEK = 2
COLUMN_NAME = 3
COLUMN_ID = 4
COLUMN_TIME = 9


class workItem:
    def __init__(self, name, idValue, startTime, endTime):
        self.name = name
        self.idValue = idValue
        self.startTime = startTime
        self.endTime = endTime

    def reset(self):
        self.name = ""
        self.endTime = ""
        self.startTime = ""
        self.idValue = ""


def isWeekend(dataStr):
    return dataStr == "星期六" or dataStr == "星期日"


def calOverTimeReal(item):
    try:
        timeOne = datetime.datetime.strptime(item.startTime, '%H:%M')
        timeTwo = datetime.datetime.strptime(item.endTime, '%H:%M')
    except ValueError:
        return

    timeDiffSec = (timeTwo - timeOne).total_seconds()
    if timeOne.hour < 12:
        # 减去午休的时间
        timeDiffSec = timeDiffSec - LAUNCH_BREAK_SECOND

    if timeDiffSec > TOTAL_WORK_SECOND:
        print(item.name + item.idValue + " 加班一天")
    else:
        print(item.name + item.idValue + " 加班半天")


def updateItem(tempItem, nameValue, idValue, timeValue):
    tempId = tempItem.idValue
    if len(tempId) == 0:
        tempItem.idValue = idValue

    sameID = tempItem.idValue == idValue
    if sameID:
        # 代表是同个用户的加班
        tempItem.name = nameValue
        if len(tempItem.startTime) == 0:
            tempItem.startTime = timeValue
        else:
            tempItem.endTime = timeValue
    else:
        # 不同用户，需要计算加班时间
        calOverTimeReal(tempItem)

    return sameID


def main():
    sourceFile = "/Users/weigan/Downloads/7月考勤全员.xlsx"
    workbook = openpyxl.load_workbook(sourceFile, read_only=True)
    sheet = workbook.worksheets[0]
    maxRow = sheet.max_row + 1
    tempItem = workItem("", "", "", "")
    for num in range(5, maxRow):
        weekValue = sheet.cell(row=num, column=COLUMN_WEEK).value
        if not isWeekend(weekValue):
            continue

        nameValue = sheet.cell(row=num, column=COLUMN_NAME).value
        idValue = sheet.cell(row=num, column=COLUMN_ID).value
        timeValue = sheet.cell(row=num, column=COLUMN_TIME).value
        result = updateItem(tempItem, nameValue, idValue, timeValue)
        if not result:
            tempItem.reset()
            tempItem.startTime = timeValue


if __name__ == '__main__':
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
