import sys
from datetime import datetime

import pyautogui

import schedule
import xlrd
import pyperclip
import time
import PyInstaller

# 定义鼠标事件

file = 'cmd.xls'
# 打开文件
wb = xlrd.open_workbook(filename=file)
# 通过索引获取表格sheet页
sheet1 = wb.sheet_by_index(0)


def mouseClick(clickTimes, lOrR, img, reTry):
    ci = 0
    if reTry == 1:
        while True:
            if img == 'beiqiangla.png':
                if pyautogui.locateCenterOnScreen('kuaisu.png', confidence=0.9) is not None:
                    print("\033[1;36m----恭喜你！预约成功！！！！-----\033[0m")
                    sys.exit()
                print("\033[1;33m常用座位（1）被抢了，我们来预约下一个\033[0m")
                location = pyautogui.locateCenterOnScreen('queren.png', confidence=0.9)
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                time.sleep(0.5)
                i = pyautogui.locateAllOnScreen('ciri.png', confidence=0.9)
                p = 0
                for x, y, h, l in i:
                    p += 1
                    if p == 1:
                        continue
                    pyautogui.click(x, y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                    time.sleep(0.1)
                    if pyautogui.locateCenterOnScreen('kuaisu.png', confidence=0.9) is not None:
                        print("\033[1;36m----恭喜你！预约成功！！！！-----\033[0m")
                        sys.exit()
                    else:
                        location = pyautogui.locateCenterOnScreen('queren.png', confidence=0.9)
                        pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2,
                                        button=lOrR)
                        print("\033[1;33m常用座位（" + str(p) + "）被抢了，我们来预约下一个\033[0m")
                        time.sleep(0.1)
                print("\033[1;31m算了吧 别卷了！\033[0m")
                sys.exit()
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                break
            ci += 1
            if ci == 8:
                location1 = pyautogui.locateCenterOnScreen('shanchu.png', confidence=0.9)
                if location1 is not None:
                    print("\033[1;31m------加载超时！重新加载！-------\033[0m")
                    pyautogui.click(location1.x, location1.y, clicks=clickTimes, interval=0.2, duration=0.2,
                                    button=lOrR)
                    mainWork(sheet1)

            print("\033[1;31m未找到匹配图片,1秒后重试----\033[0m", img)
            time.sleep(1)
    elif reTry == -1:
        while True:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
            time.sleep(0.1)
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                print("重复")
                i += 1
            time.sleep(0.1)


# 数据检查
# cmdType.value  1.0 左键单击    2.0 左键双击  3.0 右键单击  4.0 输入  5.0 等待  6.0 滚轮
# ctype     空：0
#           字符串：1
#           数字：2
#           日期：3
#           布尔：4
#           error：5
def dataCheck(sheet1):
    checkCmd = True
    # 行数检查
    if sheet1.nrows < 2:
        print("没数据啊哥")
        checkCmd = False
    # 每行数据检查
    i = 1
    while i < sheet1.nrows:
        # 第1列 操作类型检查
        cmdType = sheet1.row(i)[0]
        if cmdType.ctype != 2 or (cmdType.value != 1.0 and cmdType.value != 2.0 and cmdType.value != 3.0
                                  and cmdType.value != 4.0 and cmdType.value != 5.0 and cmdType.value != 6.0):
            print('第', i + 1, "行,第1列数据有毛病")
            checkCmd = False
        # 第2列 内容检查
        cmdValue = sheet1.row(i)[1]
        # 读图点击类型指令，内容必须为字符串类型
        if cmdType.value == 1.0 or cmdType.value == 2.0 or cmdType.value == 3.0:
            if cmdValue.ctype != 1:
                print('第', i + 1, "行,第2列数据有毛病")
                checkCmd = False
        # 输入类型，内容不能为空
        if cmdType.value == 4.0:
            if cmdValue.ctype == 0:
                print('第', i + 1, "行,第2列数据有毛病")
                checkCmd = False
        # 等待类型，内容必须为数字
        if cmdType.value == 5.0:
            if cmdValue.ctype != 2:
                print('第', i + 1, "行,第2列数据有毛病")
                checkCmd = False
        # 滚轮事件，内容必须为数字
        if cmdType.value == 6.0:
            if cmdValue.ctype != 2:
                print('第', i + 1, "行,第2列数据有毛病")
                checkCmd = False
        i += 1
    return checkCmd


# 任务
def mainWork(img):
    print()
    i = 1
    while i < sheet1.nrows:
        # 取本行指令的操作类型
        cmdType = sheet1.row(i)[0]
        if cmdType.value == 1.0:
            # 取图片名称
            img = sheet1.row(i)[1].value
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1, "left", img, reTry)
            print("\033[1;36m单击左键", img + "\033[0m")
        # 2代表双击左键
        elif cmdType.value == 2.0:
            # 取图片名称
            img = sheet1.row(i)[1].value
            # 取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(2, "left", img, reTry)
            print("\033[1;36m双击左键", img + "\033[0m")
        # 3代表右键
        elif cmdType.value == 3.0:
            # 取图片名称
            img = sheet1.row(i)[1].value
            # 取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1, "right", img, reTry)
            print("\033[1;36m右键", img + "\033[0m")
            # 4代表输入
        elif cmdType.value == 4.0:
            inputValue = sheet1.row(i)[1].value
            pyperclip.copy(inputValue)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            print("\033[1;36m输入", inputValue + "\033[0m")
            # 5代表等待
        elif cmdType.value == 5.0:
            # 取图片名称
            waitTime = sheet1.row(i)[1].value
            time.sleep(waitTime)
            print("\033[1;36m等待", waitTime, "秒\033[0m")
        # 6代表滚轮
        elif cmdType.value == 6.0:
            # 取图片名称
            scroll = sheet1.row(i)[1].value
            pyautogui.scroll(int(scroll))
            print("\033[1;36m滚轮滑动", int(scroll), "距离\033[0m")
        i += 1


if __name__ == '__main__':
    # 数据检查
    print("\033[1;36m----运行成功！-----\033[0m")

    strTime = input("\033[1;36m请输入开始时间(hh:mm)：\033[0m")
    s = datetime.strptime(strTime + ":00", "%H:%M:%S")
    schedule.every().day.at(strTime).do(mainWork, sheet1)
    checkCmd = dataCheck(sheet1)
    if checkCmd:
        while True:
            print("\033[1;36m\r倒计时:\033[0m", "\033[1;33m",(s - datetime.strptime(time.strftime("%H:%M:%S", time.localtime()), "%H:%M:%S")),
                  "\033[0m", flush=True, end='', )
            schedule.run_pending()
            time.sleep(1)
    else:
        print('\033[1;31m----输入有误或者已经退出-----\033[0m!')
