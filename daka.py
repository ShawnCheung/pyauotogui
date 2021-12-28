import pyautogui
import time
import xlrd
import pyperclip
import argparse
import torch
import cv2
from detect import detect
import numpy as np
import aircv  as ac
import time
import sys,os,signal,datetime
import threading

# check_requirements()


#定义鼠标事件

#pyautogui库其他用法 https://blog.csdn.net/qingfengxd1/article/details/108270159

def mouseClick(clickTimes,lOrR,img,reTry):
    if reTry == 1:
        while True:
            if img=="templates/new.png":
                time.sleep(5)
                pyautogui.click(1880,1000,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                print("click new")
                break
            elif img=="templates/shutdown.png":
                time.sleep(5)
                pyautogui.click(1902,17,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                print("click shutdown")
                break
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                break
            print("未找到匹配图片,0.1秒后重试")
            time.sleep(0.1)
    elif reTry == -1:
        while True:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
            time.sleep(0.1)
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location=pyautogui.locateCenterOnScreen(img,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=lOrR)
                print("重复")
                i += 1
            time.sleep(0.1)




# 数据检查
# cmdType.value  1.0 左键单击    2.0 左键双击  3.0 右键单击  4.0 输入  5.0 等待  6.0 滚轮  7.0 回车键  8.0 滑动解锁
# ctype     空：0
#           字符串：1
#           数字：2
#           日期：3
#           布尔：4
#           error：5
def dataCheck(sheet1):
    checkCmd = True
    #行数检查
    if sheet1.nrows<2:
        print("没数据啊哥")
        checkCmd = False
    #每行数据检查
    i = 1
    while i < sheet1.nrows:
        # 第1列 操作类型检查
        cmdType = sheet1.row(i)[0]
        if cmdType.ctype != 2 or (cmdType.value != 1.0 and cmdType.value != 2.0 and cmdType.value != 3.0 
        and cmdType.value != 4.0 and cmdType.value != 5.0 and cmdType.value != 6.0 and cmdType.value != 7.0 and cmdType.value != 8.0 and cmdType.value != 9.0 and cmdType.value != 10.0):
            print('第',i+1,"行,第1列数据有毛病")
            checkCmd = False
        # 第2列 内容检查
        cmdValue = sheet1.row(i)[1]
        # 读图点击类型指令，内容必须为字符串类型
        if cmdType.value ==1.0 or cmdType.value == 2.0 or cmdType.value == 3.0:
            if cmdValue.ctype != 1:
                print('第',i+1,"行,第2列数据有毛病")
                checkCmd = False
        # 输入类型，内容不能为空
        if cmdType.value == 4.0:
            if cmdValue.ctype == 0:
                print('第',i+1,"行,第2列数据有毛病")
                checkCmd = False
        # 等待类型，内容必须为数字
        if cmdType.value == 5.0:
            if cmdValue.ctype != 2:
                print('第',i+1,"行,第2列数据有毛病")
                checkCmd = False
        # 滚轮事件，内容必须为数字
        if cmdType.value == 6.0:
            if cmdValue.ctype != 2:
                print('第',i+1,"行,第2列数据有毛病")
                checkCmd = False
        i += 1
    return checkCmd

#任务
def mainWork():
    i = 1
    while i < sheet1.nrows:
        #取本行指令的操作类型
        cmdType = sheet1.row(i)[0]
        if cmdType.value == 1.0:
            #取图片名称
            img = sheet1.row(i)[1].value
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1,"left",img,reTry)
            print("单击左键",img)
        #2代表双击左键
        elif cmdType.value == 2.0:
            #取图片名称
            img = sheet1.row(i)[1].value
            #取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(2,"left",img,reTry)
            print("双击左键",img)
        #3代表右键
        elif cmdType.value == 3.0:
            #取图片名称
            img = sheet1.row(i)[1].value
            #取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1,"right",img,reTry)
            print("右键",img) 
        #4代表输入
        elif cmdType.value == 4.0:
            inputValue = sheet1.row(i)[1].value
            pyperclip.copy(inputValue)
            pyautogui.hotkey('ctrl','v')
            time.sleep(0.5)
            print("输入:",inputValue)                                        
        #5代表等待
        elif cmdType.value == 5.0:
            #取图片名称
            waitTime = sheet1.row(i)[1].value
            time.sleep(waitTime)
            print("等待",waitTime,"秒")
        #6代表滚轮
        elif cmdType.value == 6.0:
            #取图片名称
            scroll = sheet1.row(i)[1].value
            pyautogui.scroll(int(scroll))
            print("滚轮滑动",int(scroll),"距离")      
        elif cmdType.value == 7.0:
            #取图片名称
            pyautogui.hotkey('enter')
            print("回车键")   
        elif cmdType.value == 8.0:
            slider = cv2.imread(sheet1.row(i)[1].value)
            imtmp1 = cv2.imread("./templates/model1.png")
            imtmp2 = cv2.imread("./templates/model2.png")
            temple = cv2.imread("./templates/match.png")

            time.sleep(1)
            imgg = pyautogui.screenshot() 
            im = cv2.cvtColor(np.asarray(imgg),cv2.COLOR_RGB2BGR)
            
            conf = ac.find_template(im, temple, 0.2)['confidence']
            print("fisrt conf", conf)
            while (conf>0.9):
                pyautogui.click(1641,366,clicks=1,interval=0.2,duration=0.2,button="left")
                time.sleep(10)
                screenshot = pyautogui.screenshot() 
                im = cv2.cvtColor(np.asarray(screenshot),cv2.COLOR_RGB2BGR)
                pos1 = ac.find_template(im, imtmp1, 0.2)
                pos2 = ac.find_template(im, imtmp2, 0.2)

                crop = im[pos1["rectangle"][0][1]: pos2["rectangle"][3][1], pos1["rectangle"][0][0]: pos2["rectangle"][3][0], :]
                imout = crop[62:62+153, 20:20+274]
                try:
                    cv2.imwrite("temp/crop.jpg", imout)
                except:
                    break        
                with torch.no_grad():
                    out = detect(opt)
                
                print(out)

                pos = ac.find_template(im, slider, 0.2)

                pyautogui.mouseDown(x=int(pos["result"][0]), y=int(pos["result"][1]), button="primary", duration=0)
                pyautogui.mouseUp(x=int(pos["result"][0])+int(out*274)-13, y=int(pos["result"][1]), button="primary", duration=2) 
                time.sleep(10)

                ss = pyautogui.screenshot() 
                del im
                im = cv2.cvtColor(np.asarray(ss),cv2.COLOR_RGB2BGR)
                print("im.shape", im.shape)
                findd = ac.find_template(im, temple, 0.2)
                time.sleep(3)

                if findd is None:
                    print("conf is None")
                    break
                conf = findd['confidence']
                print("second conf", conf)


        elif cmdType.value == 9.0:
            print("xiahau")
            pyautogui.mouseDown(x=1915, y=200, button='primary', duration=0)
            pyautogui.mouseUp(x=1915, y=1000, button='primary', duration=0)
            pyautogui.press('pagedown')
        elif cmdType.value == 10.0:
            pyautogui.hotkey('tab')
            print("tab按键")

        i += 1

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--weights', nargs='+', type=str, default='./best.pt', help='model.pt path(s)')
    parser.add_argument('--source', type=str, default='./temp/', help='source')  # file/folder, 0 for webcam
    parser.add_argument('--img-size', type=int, default=640, help='inference size (pixels)')
    parser.add_argument('--conf-thres', type=float, default=0.25, help='object confidence threshold')
    parser.add_argument('--iou-thres', type=float, default=0.45, help='IOU threshold for NMS')
    parser.add_argument('--device', default='', help='cuda device, i.e. 0 or 0,1,2,3 or cpu')
    parser.add_argument('--view-img', action='store_true', help='display results')
    parser.add_argument('--save-txt', action='store_true', help='save results to *.txt')
    parser.add_argument('--save-conf', action='store_true', help='save confidences in --save-txt labels')
    parser.add_argument('--classes', nargs='+', type=int, help='filter by class: --class 0, or --class 0 2 3')
    parser.add_argument('--agnostic-nms', action='store_true', help='class-agnostic NMS')
    parser.add_argument('--augment', action='store_true', help='augmented inference')
    parser.add_argument('--update', action='store_true', help='update all models')
    parser.add_argument('--project', default='runs/detect', help='save results to project/name')
    parser.add_argument('--name', default='exp', help='save results to project/name')
    parser.add_argument('--exist-ok', action='store_true', help='existing project/name ok, do not increment')
    parser.add_argument('--index', type=int, help='existing project/name ok, do not increment')
    opt = parser.parse_args()
    def after_timeout():
        now_time = datetime.datetime.now()
        hour = now_time.hour
        minute = now_time.minute
        if minute>40:
            print("打卡超时")
            kill_process()
        else:
            os.system("/home/shawn/.pyenv/versions/yolov5/bin/python daka.py --index {}".format(opt.index))

    def kill_process():
        print("KILL THE WORLD HERE!")
        os.kill(os.getpid(), signal.SIGKILL)

    threading.Timer(180, after_timeout).start()
    threading.Timer(182, kill_process).start()
    file = 'daka.xls'
    #打开文件
    wb = xlrd.open_workbook(filename=file)
    #通过索引获取表格sheet页
    print(opt.index-1)
    sheet1 = wb.sheet_by_index(opt.index-1)
    print('欢迎使用不高兴就喝水牌RPA~')
    #数据检查
    checkCmd = dataCheck(sheet1)
    if checkCmd:
        # key=input('选择功能: 1.做一次 2.循环到死 \n')
        mainWork()

        # if key=='1':
        #     #循环拿出每一行指令
        #     mainWork()
        # elif key=='2':
        #     while True:
        #         mainWork()
        #         time.sleep(0.1)
        #         print("等待0.1秒")    
    else:
        print('输入有误或者已经退出!')
    os.kill(os.getpid(), signal.SIGKILL)
    