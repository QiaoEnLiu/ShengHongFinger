import ctypes
from ctypes import *
import threading
import numpy as np
import re
import datetime
import base64
import os
from Crypto.Cipher import AES
from binascii import a2b_hex, b2a_hex
from tkinter import filedialog, Tk
import tkinter.scrolledtext
from tkinter import ttk, simpledialog
from tkinter import *
import tkinter
import sqlite3
from io import BytesIO
from PIL import ImageTk,Image,ImageFilter
import tkinter.messagebox as msgbox

from PyPDF2 import PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

import shutil, tempfile

pdfmetrics.registerFont(TTFont('ChineseFont', 'word_font/Chinese/zh/msjh.ttc'))

left_fingers_text=['左手小指','左手無名指','左手中指','左手食指','左手大拇指']
right_fingers_text=['右手大拇指','右手食指','右手中指','右手無名指','右手小指']

global left_hand, left_Pinky, left_Ring, left_Middle, left_Index, left_Thumb
global right_hand, right_Thumb, right_Index, right_Middle, right_Ring, right_Pinky

left_hand_labels = []
right_hand_labels = []

stop_all_threads = False


# 定义全局变量
global root,text,KEY,SN_list, isOpen, isRunning,feature,\
    HANDLE,imageWidth, imageHeight,rollImageWidth,rollImageHeight,ImageSize,ImgPtr,gets
global lableShowImage1,lableShowImage2,lableShowFingerImage,openButton,\
    ShowImage1,ShowImage2,ShowFingerImage,bmpImage1,bmpImage2,bmpFingerImg
# 初始化全局变量
isOpen = False  # 指纹仪是否打开
isRunning = False  # 是否正在运行线程
root = Tk()  # 定义一个tkinter图形化界面
HANDLE = c_void_p(0)  # 句柄初始化
imageWidth = c_int()  # 指纹图片宽度
imageHeight = c_int()  # 指纹图片高度
rollImageWidth = None  # 指纹图片高度
rollImageHeight = None  # 指纹图片高度
openButton=None  #开始按钮
ImageSize=0  #图片大小
text=None  #信息框
fingers_label=None
lableShowImage1=None  #滚动图片框
lableShowImage2=None  #滚动合成图片框
lableShowFingerImage=None  #滚动合成图片框
ShowImage1=None  #滚动图片
ShowImage2=None  #滚动合成图片
ShowFingerImage=None  #滚动合成图片
bmpImage1=None  #滚动图片数据
bmpImage2=None  #滚动合成图片数据
bmpFingerImg=None  #WMRAPI图片数据
KEY=b'wm20180915144830'
SN_list=[]
ImgPtr=None
gets=None
feature = create_string_buffer(2048)  # 创建2048字节长的指纹特征，且初始化为0


dllinfo=""
# 导入dll
add_path1="WMRAPI.dll"
add_path2="ftrScanAPI.dll"
filelic="WMR08-Plus.lic"
dllExist1=os.path.exists(add_path1)
dllExist2=os.path.exists(add_path2)
fileExist=os.path.exists(add_path2)
if dllExist1==True and dllExist2==True:
    WMdll = WinDLL(add_path1)  ##调用dll文件
    ftrdll = WinDLL(add_path2)  ##调用dll文件
else:
    if dllExist1==False and dllExist2==False:
        dllinfo="WMRAPI.dll、ftrScanAPI.dll不存在\r\n"
    elif dllExist1==False:
        dllinfo="WMRAPI.dll不存在\r\n"
    else:
        dllinfo="ftrScanAPI.dll不存在\r\n"


#region Cache
def clear_finger_cache():
    cacheInfo=""
    try:
        # 定義目標資料夾名稱
        target_folder_name = "fingerCache"

        # 建立完整的資料夾路徑
        target_folder_path = os.path.join(os.getcwd(), target_folder_name)

        # 檢查資料夾是否存在
        if os.path.exists(target_folder_path):
            # 獲取資料夾內所有檔案及子資料夾
            folder_content = os.listdir(target_folder_path)

            # 刪除每個檔案
            for item in folder_content:
                item_path = os.path.join(target_folder_path, item)
                
                # 如果是檔案，則刪除
                if os.path.isfile(item_path):
                    os.remove(item_path)
                # 如果是資料夾，則遞迴清空
                elif os.path.isdir(item_path):
                    shutil.rmtree(item_path)

            print(f"成功清空資料夾內容：{target_folder_path}")
            cacheInfo="成功清空資料夾內容：{target_folder_path}\r\n"
        else:
            print(f"資料夾不存在：{target_folder_path}")
            cacheInfo="資料夾不存在：{target_folder_path}\r\n"
    except Exception as e:
        print(f"清空資料夾內容時發生錯誤：{e}")
        cacheInfo="資料夾不存在：{target_folder_path}\r\n"

    MessageText(cacheInfo)
#endregion

#region UI
def UI():
    
    auxiliary=auxiliaryMeans()
    mean=means()
    root.title('WMR08-Plus')  # 窗口名称
    root.geometry('1920x1080')  # 窗口大小，这里的乘号不是 * ，而是小写英文字母 x
    screen_width = root.winfo_screenwidth()  # 显示屏大小
    screen_height = root.winfo_screenheight()  # 显示屏大小
    root.geometry(f'1920x1080+{round((screen_width - 1920) / 2)}+{round((screen_height - 1080) / 2)}')
    root.resizable(0, 0)  # 固定窗口大小，不能随意伸缩
    root.protocol('WM_DELETE_WINDOW', mean.Exit)
    auxiliary.icoImage()

    ButtonGroup(root)
    messageScrolledText(root)  # 调用信息框
    # FingerprintLibrary()
    showImage1(root)  # 调用滚动采集图片框
    # showImage2(root)  # 调用图片框
    clear_finger_cache()
    root.mainloop()  # 显示图形化界面
#endregion


#region
def update_image(label, image, photo):
    photo = ImageTk.PhotoImage(image)
    label.config(image=photo)
    label.photo = photo  # 保持對photo的引用，防止被垃圾回收
#endregion


#region 滚动采集图片框
def showImage1(root):
    global fingers_label, left_hand, right_hand, lableShowImage1, lableShowImage2, bmpImage1, bmpImage2
    # auxiliary = auxiliaryMeans()
    labFrame = LabelFrame(root, text="滚动采集指纹图像", relief=GROOVE, width=1875, height=950,font=(None,20))  # 设置图片样式
    labFrame.place(relx=0.01, rely=0.1)

    fingers_label = Label(labFrame,text='手指',font=(None,16))
    fingers_label.place(relx=1, rely=0, anchor='e')

    
    lableShowImage1 = Label(labFrame)  # 包含图片的标签
    lableShowImage1.config(text='每幀圖測試',compound=tkinter.BOTTOM,font=(None,20))
    lableShowImage1.place(relx=0, rely=0.025)
    # lableShowImage1.bind("<Button-1>", lambda e: auxiliary.my_label(e, bmpImage1))

    lableShowImage2 = Label(labFrame)  # 包含图片的标签
    lableShowImage2.config(text='每幀圖合成',compound=tkinter.BOTTOM,font=(None,20))
    lableShowImage2.place(relx=0, rely=0.525)
    # lableShowImage2.bind("<Button-1>", lambda e: auxiliary.my_label(e, bmpImage2))


    for i in range(5):
        left_label = Label(labFrame,font=(None,20))
        right_label = Label(labFrame,font=(None,20))

        left_label.place(relx=0.15 + i*0.15, rely=0.025)
        right_label.place(relx=0.15 + i*0.15, rely=0.525)

        left_hand_labels.append(left_label)
        right_hand_labels.append(right_label)

        
    left_hand = left_hand_labels
    right_hand = right_hand_labels

    return lableShowImage1, lableShowImage2
#endregion


#region 图片框
def showImage2(root):
    global lableShowFingerImage, bmpFingerImg
    auxiliary = auxiliaryMeans()
    labFrame = LabelFrame(root, text="指纹图像", relief=GROOVE, width=1000, height=500,font=(None,20))  # 设置图片样式
    labFrame.place(relx=0.4, rely=0.5)

    lableShowFingerImage = Label(labFrame)  # 包含图片的标签
    lableShowFingerImage.place(relx=0.15, rely=0)
    lableShowFingerImage.bind("<Button-1>", lambda e: auxiliary.my_label(e, bmpFingerImg))

    return bmpFingerImg
#endregion

#region messageScrolledText
def messageScrolledText(root):
    global text
    text = tkinter.scrolledtext.ScrolledText(root, width=75, height=5)  # 设置信息框样式
    # text.configure(font=("Arial", 20))
    text.place(relx=1, rely=0.01, anchor='ne')
    text.see(END)  # 信息框处于滚动条最下面信息的位置
    return text
#endregion

#region ButtonGroup
def ButtonGroup(root):
    global openButton  # 使用全局变量
    mean=means()
    thread=threadGroup()
    # command为按钮的点击事件，不能加括号，如command=openButton_Click而不是command=openButton_Click()
    state = tkinter.DISABLED
    font=(None, 16)
    openButton = Button(root, text='打开设备', width=12,height=2, font=font,command=mean.BeginOrClose)
    featureButton = Button(root, text='注册',state = state, width=12, height=2, font=font, command=thread.featureButton_Click)
    onCompareFingerButton = Button(root, text='1:1比对',state = state, width=12, height=2, font=font, command=thread.onCompareFingerButton_Click)
    onCompareIdentifyButton = Button(root, text='1:N识别',state = state, width=12, height=2, font=font, command=thread.onCompareIdentifyButton_Click)
    onCompareIdentifyNButton = Button(root, text='1:N查重', state=state, width=12, height=2, font=font,command=thread.onCompareIdentifyNButton_Click)

    rollStartButton = Button(root, text='滚动采集',state = state, width=12, height=2, font=(None,16), command=thread.rollStartButton_Click)
    stopButton = Button(root, text='停止',state = state, width=12, height=2, font=font, command=mean.Stop)
    SingleDeleteButton = Button(root, text='单一删除',state = state, width=12, height=2, font=font, command=mean.SingleDelete)
    AllDeleteButton = Button(root, text='清空指纹库', state=state, width=12, height=2, font=font,command=mean.AllDelete)
    exitButton = Button(root, text='退出', width=12, height=2, font=font, command=mean.Exit)

    systemLogButton = Button(root, text='系統日誌', width=12, height=2, font=(None,16), command=thread.systemLogButton_Click)
    allFingersButton = Button(root, text='採集所有手指',state = state, width=12, height=2, font=(None,16), command=thread.allFingersButton_Click)
    savePictureFile = Button(root, text='儲存圖檔',state = state, width=12, height=2, font=(None,16), command=thread.savePictureFileButton_Click)
    savePDF = Button(root, text='儲存PDF檔',state = tkinter.DISABLED, width=12, height=2, font=(None,16), command=thread.savePDF_Button_Click)

    # 按钮的位置，范围0-1
    # openButton.place(relx=0.02, rely=0.01)
    # featureButton.place(relx=0.02, rely=0.09)
    # onCompareFingerButton.place(relx=0.02, rely=0.17)
    # onCompareIdentifyButton.place(relx=0.02, rely=0.25)
    # onCompareIdentifyNButton.place(relx=0.02, rely=0.33)

    # rollStartButton.place(relx=0.12, rely=0.01)
    # stopButton.place(relx=0.12, rely=0.09)
    # SingleDeleteButton.place(relx=0.12, rely=0.17)
    # AllDeleteButton.place(relx=0.12, rely=0.25)
    # exitButton.place(relx=0.12, rely=0.33)

    # allFingersButton.place(relx=0.22, rely=0.01)
    # savePictureFile.place(relx=0.22, rely=0.09)
    # savePDF.place(relx=0.22, rely=0.17)


    openButton.place(relx=0.01, rely=0.01)
    rollStartButton.place(relx=0.10, rely=0.01)
    allFingersButton.place(relx=0.19, rely=0.01)
    stopButton.place(relx=0.28, rely=0.01)
    savePDF.place(relx=0.37, rely=0.01)
    systemLogButton.place(relx=0.46, rely=0.01)
    exitButton.place(relx=0.55, rely=0.01)

    return [rollStartButton,featureButton,onCompareFingerButton,onCompareIdentifyButton,onCompareIdentifyNButton,stopButton,SingleDeleteButton,AllDeleteButton, allFingersButton,savePDF]
#endregion

#region threadGroup
class threadGroup:
    
    def rollStartButton_Click(self):
        mean = means()  # 事件类
        open_thread = threading.Thread(target=mean.RollStart)
        open_thread.start()

    def featureButton_Click(self):
        global id, mode
        mean = means()
        auxiliary=auxiliaryMeans()
        entry_int = auxiliary.Input_id()
        if entry_int != None:
            db = Sqlite3()
            sql = "select id from Fingers where id = " + str(entry_int) + ";"
            FingerID = db.Select(sql)
            # 当前用户不存在
            if FingerID == None:
                mode = 0
                id = entry_int
                feature_thread = threading.Thread(target=mean.Feature)
                feature_thread.start()
            else:
                warning = msgbox.askokcancel('确认操作', '该用户已存在，是否覆盖？')  # 返回值true/false
                if warning == True:
                    mode = 1
                    id = entry_int
                    feature_thread = threading.Thread(target=mean.Feature)
                    feature_thread.start()

    def onCompareFingerButton_Click(self):
        mean=means()
        verify_thread = threading.Thread(target=mean.onCompareFinger)
        verify_thread.start()

    def onCompareIdentifyButton_Click(self):
        mean = means()
        verify_thread = threading.Thread(target=mean.onCompareIdentify)
        verify_thread.start()

    def onCompareIdentifyNButton_Click(self):
        mean = means()
        verify_thread = threading.Thread(target=mean.onCompareIdentifyN)
        verify_thread.start()

    #zh-tw
    def allFingersButton_Click(self):
        global stop_all_threads
        stop_all_threads=False
        self.current_finger_index = 0
        self.all_finger_text = left_fingers_text + right_fingers_text
        self.all_finger = left_hand + right_hand

        # 創建 Means 實例
        def start_finger_collection():
            global stop_all_threads
            if not stop_all_threads and self.current_finger_index < len(self.all_finger):
                fingers_text = self.all_finger_text[self.current_finger_index]
                finger = self.all_finger[self.current_finger_index]

                fingers_label.config(text="採集" + fingers_text)
                fingers_label.text = fingers_text
                print(fingers_text + "\r\n")
                MessageText(fingers_text + "\r\n")

                current_means = means()

                # 檢查是否應該停止
                if not stop_all_threads:
                    open_thread = threading.Thread(target=current_means.RollStartAllFingers, args=(finger, fingers_text))
                    open_thread.start()

                    # 定期檢查執行緒是否完成
                    check_thread_completion(open_thread)
                else:
                    print("手動停止\r\n")
                    MessageText("手動停止\r\n")

        def check_thread_completion(thread):
            if thread.is_alive() and not stop_all_threads:
                # 執行緒還在執行中，定時檢查
                root.after(10, lambda: check_thread_completion(thread))
            else:
                # 執行緒已完成，進行下一次迴圈
                self.current_finger_index += 1
                root.after(10, start_finger_collection)

        clear_finger_cache()
        print('採集所有手指')
        MessageText("開始採集所有手指\r\n")

        for i in self.all_finger:
            i.config(text='', image='')

        # 左右手指紋採集
        start_finger_collection()

    def systemLogButton_Click(self):
        # print('系統日誌')
        # MessageText("系統日誌\r\n")
        mean = means()
        verify_thread = threading.Thread(target=mean.systemLogInterface)
        verify_thread.start()

    def savePictureFileButton_Click(self):
        print('儲存圖檔')
        MessageText("儲存圖檔\r\n")

    def savePDF_Button_Click(self):
        print('儲存PDF')
        MessageText("儲存PDF\r\n")
        SavePDFFile.create_pdf(self)

#endregion

#region means
class means:
    finger_images=[]


    def getSNList(self):
        global SN_list
        auxiliary=auxiliaryMeans()
        if fileExist == True:
            files=open(filelic,"rb")
            fileData=files.read()
            SN_list=auxiliary.AES_de(fileData,KEY)
            print(SN_list)
            return SN_list
        else:
            MessageText("WMR08-Plus.lic不存在\r\n")
            return None

    def BeginOrClose(self):
        global message, isOpen, openButton, isRunning
        mean = means()
        auxiliary=auxiliaryMeans()


        if dllinfo != "":
            MessageText(dllinfo)
            return
        WMRAPI = WMRAPI_Dll()
        if isOpen == False:
            mean.getSNList()
            validTime=auxiliary.compareTime(SN_list[0])
            if validTime==True:
                result = WMdll.WM_GetDeviceCount()
                if result == 0:
                    info = "当前环境没有任何指纹设备\r\n"
                    MessageText(info)
                    return
                else:
                    info = "当前电脑中可用设备有" + str(result) + "个\r\n"
                    result = WMdll.WM_Init()
                    if result != 0:
                        info = info + "初始化设备失败，错误码：" + str(result) + "\r\n"
                    else:
                        info = info + "初始化设备成功\r\n"

                    handle = WMRAPI.OpenDevice()
                    if handle[0] != 0:
                        MessageText("打开设备失败\r\n")
                        return

                    SN = WMRAPI.GetSerialNumber()
                    if SN[2] != "":  # 设备有SN
                        if SN[2] in SN_list[1:]:
                            auxiliary.anConfig(root, tkinter.NORMAL)  # 启用按钮
                            # Fingerprint = FingerprintLibrary() #暫時用不到
                            # Fingerprint.Select()  # 显示指纹库数据
                            info = info+"打开设备成功\r\n序列号:" + SN[2]
                            isOpen = True
                            openButton.config(text="关闭设备")
                            MessageText(info+"\r\n")
                        else:
                            WMRAPI.CloseDevice()
                            MessageText("设备未授权\r\n")
                    else:
                        WMRAPI.CloseDevice()
                        MessageText("设备非法\r\n")
            else:
                MessageText("授权文件已过期\r\n")
        else:
            result = WMRAPI.CloseDevice()
            if result == 0:
                isOpen = False
                isRunning=False
                auxiliary.anConfig(root, tkinter.DISABLED)  # 启用按钮
                openButton.config(text="打开设备")
                MessageText("关闭设备成功\r\n")
            else:
                MessageText("设备关闭失败\r\n")
    
    def RollStart(self):
        global isOpen, isRunning, message, ShowImage1, ShowImage2, bmpImage1, bmpImage2
        WMRAPI = WMRAPI_Dll()
        auxiliary=auxiliaryMeans()
        baseImage = None
        WMRAPI.SetOptions()
        startTime = datetime.datetime.now()
        imgResizeRate=0.82
        flag = False
        if isOpen == True:
            isRunning = True
            flag = True
            MessageText("按下手指，开始滚动采集指纹\r\n")
        while (flag):
            flag = isRunning  # 线程开启状态
            if flag == False:
                return
            RAW = WMRAPI.GetFrame()
            if RAW[0] == 1:  # 从手指第一次按下时开始计算
                BMP = WMRAPI.RawToBMP(RAW[1], rollImageWidth, rollImageHeight)
                byte_stream = BytesIO(BMP[1].raw)  # 将二进制转为字节流
                roiImg = Image.open(byte_stream)
                resizeImg1 = roiImg.resize((int(rollImageWidth * imgResizeRate), int(rollImageHeight * imgResizeRate)))
                img = ImageTk.PhotoImage(resizeImg1)
                lableShowImage1.config(image=img)  # 显示最新一张实时图片
                bmpImage1 = roiImg
                # 底图
                baseImage = auxiliary.RemoveBackground(roiImg)
                p = Image.new('RGBA', baseImage.size, (0, 0, 0))
                baseImage = baseImage.filter(ImageFilter.DETAIL)  # 细节增强
                p.paste(baseImage, mask=baseImage)
                resizeImg=p.resize((int(rollImageWidth*imgResizeRate),int(rollImageHeight*imgResizeRate)))
                img2 = ImageTk.PhotoImage(resizeImg)
                lableShowImage2.config(image=img2)  # 显示最新一张显示图片
                ShowImage1 = img
                ShowImage2 = img2
                bmpImage2 = p
                break
        while (flag):
            flag = isRunning  # 线程开启状态
            if flag == False:
                return
            RAW = WMRAPI.GetFrame()
            if RAW[0] == 0:
                MessageText("手指抬起，指纹滚动采集结束\r\n")
                break
            print(RAW)
            if RAW[0] == 1:
                BMP = WMRAPI.RawToBMP(RAW[1],rollImageWidth,rollImageHeight)
                byte_stream = BytesIO(BMP[1].raw)  # 将二进制转为字节流
                roiImg = Image.open(byte_stream)
                roiImg1 = auxiliary.RemoveBackground(roiImg)
                resizeImg1 = roiImg.resize((int(rollImageWidth * imgResizeRate), int(rollImageHeight * imgResizeRate)))
                img = ImageTk.PhotoImage(resizeImg1)
                lableShowImage1.config(image=img)  # 显示最新一张实时图片
                lableShowImage1.update()
                ShowImage1 = img
                bmpImage1 = roiImg
                if (datetime.datetime.now() >= startTime + datetime.timedelta(seconds=0.5)):
                    startTime = datetime.datetime.now()
                    baseImage.paste(roiImg1, mask=roiImg1)
                    p = Image.new('RGBA', baseImage.size, (0, 0, 0))
                    baseImage2 = baseImage.filter(ImageFilter.DETAIL)  # 细节增强
                    p.paste(baseImage2, mask=baseImage2)
                    resizeImg2 = p.resize((int(rollImageWidth * imgResizeRate), int(rollImageHeight * imgResizeRate)))
                    img2 = ImageTk.PhotoImage(resizeImg2)
                    lableShowImage2.config(image=img2)  # 显示最新一张显示图片
                    lableShowImage2.update()
                    ShowImage2 = img2
                    bmpImage2 = p

    # 掃描所有手指
    def RollStartAllFingers(self, finger, text):
        global isOpen, isRunning, message, ShowImage1, ShowImage2, bmpImage1, bmpImage2
        WMRAPI = WMRAPI_Dll()
        auxiliary=auxiliaryMeans()
        baseImage = None
        WMRAPI.SetOptions()
        startTime = datetime.datetime.now()
        imgResizeRate=0.82
        if isOpen == True:
            isRunning = True
            flag = True
            MessageText("按下手指，开始滚动采集指纹\r\n")
        while (flag):
            flag = isRunning  # 线程开启状态
            if flag == False:
                return
            RAW = WMRAPI.GetFrame()
            if RAW[0] == 1:  # 从手指第一次按下时开始计算
                BMP = WMRAPI.RawToBMP(RAW[1], rollImageWidth, rollImageHeight)
                byte_stream = BytesIO(BMP[1].raw)  # 将二进制转为字节流
                roiImg = Image.open(byte_stream)
                resizeImg1 = roiImg.resize((int(rollImageWidth * imgResizeRate), int(rollImageHeight * imgResizeRate)))
                img = ImageTk.PhotoImage(resizeImg1)
                lableShowImage1.config(image=img)  # 显示最新一张实时图片
                bmpImage1 = roiImg
                # 底图
                baseImage = auxiliary.RemoveBackground(roiImg)
                p = Image.new('RGBA', baseImage.size, (0, 0, 0))
                baseImage = baseImage.filter(ImageFilter.DETAIL)  # 细节增强
                p.paste(baseImage, mask=baseImage)
                resizeImg=p.resize((int(rollImageWidth*imgResizeRate),int(rollImageHeight*imgResizeRate)))
                img2 = ImageTk.PhotoImage(resizeImg)
                self.finger_images.append(img2)
                lableShowImage2.config(image=img2)  # 显示最新一张显示图片
                # finger.config(text=text,image=img2,compound=tkinter.BOTTOM)    
                ShowImage1 = img
                ShowImage2 = img2
                bmpImage2 = p
                break
        while (flag):
            flag = isRunning  # 线程开启状态
            if flag == False:
                return
            RAW = WMRAPI.GetFrame()
            if RAW[0] == 0:
                MessageText("手指抬起，指纹滚动采集结束\r\n")
                break
            print(RAW)
            if RAW[0] == 1:
                BMP = WMRAPI.RawToBMP(RAW[1],rollImageWidth,rollImageHeight)
                byte_stream = BytesIO(BMP[1].raw)  # 将二进制转为字节流
                roiImg = Image.open(byte_stream)
                roiImg1 = auxiliary.RemoveBackground(roiImg)
                resizeImg1 = roiImg.resize((int(rollImageWidth * imgResizeRate), int(rollImageHeight * imgResizeRate)))
                img = ImageTk.PhotoImage(resizeImg1)
                lableShowImage1.config(image=img)  # 显示最新一张实时图片
                lableShowImage1.update()
                ShowImage1 = img
                bmpImage1 = roiImg
                if (datetime.datetime.now() >= startTime + datetime.timedelta(seconds=0.5)):
                    startTime = datetime.datetime.now()
                    baseImage.paste(roiImg1, mask=roiImg1)
                    p = Image.new('RGBA', baseImage.size, (0, 0, 0))
                    baseImage2 = baseImage.filter(ImageFilter.DETAIL)  # 细节增强
                    p.paste(baseImage2, mask=baseImage2)
                    resizeImg2 = p.resize((int(rollImageWidth * imgResizeRate), int(rollImageHeight * imgResizeRate)))
                    img2 = ImageTk.PhotoImage(resizeImg2)
                    self.finger_images.append(img2)
                    lableShowImage2.config(image=img2)  # 显示最新一张显示图片
                    lableShowImage2.update()
                    # finger.config(text=text,image=img2,compound=tkinter.BOTTOM)  # 显示最新一张显示图片
                    # finger.update()
                    ShowImage2 = img2
                    bmpImage2 = p
        finger.config(text=text, image=self.finger_images[-1], compound=tkinter.BOTTOM)
        image_format = roiImg.format
        print(f"圖片格式: {text}.{image_format}")
        MessageText(f"圖片格式: {text}.{image_format}\r\n")
        if bmpImage2 is not None:
            output_path = f'fingerCache/{text}.bmp'  # 替換成你想保存的路徑
            bmpImage2.save(output_path)
            print(f"圖檔已成功輸出至: {output_path}")
            MessageText(f"圖檔已成功輸出至: {output_path}\r\n")
        else:
            print("未獲取有效的圖檔來進行輸出")
            MessageText(f"未獲取有效的圖檔來進行輸出\r\n")
                    
        # finger_image = PhotoImage(data=finger.tk.call(finger._w, "image", "-data", finger._w))
        # finger.config(text=text, image=finger_image, compound=tkinter.BOTTOM)

    
    # 注册
    def Feature(self):
        global img, isOpen, isRunning
        isRunning = False
        if isOpen == True:
            re = self.Collect()  # 采集指纹
            if re != 0:  # 采集指纹中途不曾停止，指纹图片数量足够
                self.GetFeature()  # 合成模板
            isRunning = False  # 线程停止
        else:
            result = "指纹设备未打开，无法注册\r\n"
            MessageText(result)

    # 收集指纹
    def Collect(self):
        global ImgPtr, isRunning, ShowFingerImage,bmpFingerImg
        ImgPtr = []  # 每次点击注册都要初始化一次指纹列表
        flag = True
        isRunning = True
        WMRAPI = WMRAPI_Dll()
        try:
            for i in range(3):
                info = "请第" + str(i + 1) + "次按压手指\r\n"
                MessageText(info)

                while flag:
                    result = WMRAPI.GetImage()
                    # result = WMRAPI.GetFrame()
                    flag = isRunning
                    if result[0] == 0:
                        # 显示图像
                        BMP = WMRAPI.RawToBMP(result[2],imageWidth,imageHeight)
                        print("RawToBMP:" + str(BMP))
                        if BMP[0] != 0:
                            info = "图像质量不佳，请再次采集\r\n"
                            MessageText(info)
                            continue
                        ImgPtr.append(BMP[1].raw)  # 将图片的字节数组存入Imgptr
                        byte_stream = BytesIO(ImgPtr[i])  # 将二进制转为字节流
                        roiImg = Image.open(byte_stream)
                        print("hhh" + ImgPtr[i].hex())
                        info = "按压成功，已获取第" + str(i + 1) + "次图像\r\n"
                        break
                    if flag == False:
                        return 0
                # roiImg1 = roiImg.resize((int(imageWidth * 0.95), int(imageHeight * 0.95)))
                img = ImageTk.PhotoImage(roiImg)
                lableShowFingerImage.config(image=img)
                lableShowFingerImage.update()
                ShowFingerImage=img
                bmpFingerImg=roiImg
                MessageText(info)
        except Exception as e:
            print(e)
        return 1

    # 合成模板
    def GetFeature(self):
        Fingerprint = FingerprintLibrary()
        WMRAPI = WMRAPI_Dll()
        try:
            # 合成指纹模板
            result = WMRAPI.GenTemplateWithImage(ImgPtr)
            if result[0] == 0:
                info = "注册指纹模板已生成\r\n"
                str1 = result[1]
                info = info + "注册指纹模板数据:\r\n" + str1 + "\r\n"
                info = info + "用户添加成功，FingerID=" + str(id) + "\r\n"
                db = Sqlite3()
                if mode == 0:
                    db.Insert(id, result[3])
                else:
                    db.Update(id, result[3])
            else:
                info = "指纹模板合成失败，请重新注册\r\n"
            Fingerprint.Select()
            MessageText(info)
        except Exception as e:
            print(e)

    # 1:1比对
    def onCompareFinger(self):
        global isOpen, isRunning, gets
        isRunning = False
        if gets == None:
            info = "请选择要比对的指纹模板..\r\n"
            MessageText(info)
            # lableShowFingerImage.config(image=fingerImg)  # 显示最新一张图片在图片框
        else:
            info = "请按压手指进行比对..\r\n"
            WMRAPI = WMRAPI_Dll()
            # images.config(image=img)  # 显示最新一张图片在图片框
            MessageText(info)
            result = self.meansDoVerify()  # 获取比对信息
            # print("result",result)
            if result[1] != 0:  # 获取到指纹特征图片
                db = Sqlite3()
                sql = "select data from Fingers where id = " + gets + ";"
                template = db.Select(sql)
                print(template[0])
                # 进行比对
                Ver = WMRAPI.Verify(template[0])
                if Ver[0] == 0:
                    info = result[0] + "比对成功，比对分数：" + str(Ver[1].value) + "\r\n"
                else:
                    info = result[0] + "比对失败，如需重试请点击 开始比对按钮，比对分数：" + str(Ver[1].value) + "\r\n"
                # img = ImageTk.PhotoImage(result[1])  # 更新图片
                # lableShowFingerImage.config(image=img)  # 显示最新一张图片
                MessageText(info)
            else:
                info = "获取指纹特征失败"
                MessageText(info)

    # 1:N识别
    def onCompareIdentify(self):
        global isOpen, isRunning,ShowFingerImage,bmpFingerImg
        isRunning = False

        info = "请按压手指进行识别..\r\n"
        WMRAPI = WMRAPI_Dll()
        # images.config(image=img)  # 显示最新一张图片在图片框
        MessageText(info)
        result = self.meansDoVerify()  # 获取比对信息
        if result[1] != 0:  # 获取到指纹特征图片
            db = Sqlite3()
            templates = db.SelectAll()
            Score = 0
            id = 0
            for i in templates:
                # 进行比对
                print(i[1])
                Ver = WMRAPI.Verify(i[1])
                if Ver[0] == 0:
                    if Ver[1].value > Score:
                        Score = Ver[1].value
                        id = i[0]
            if Score == 0:
                info = result[0] + "识别失败。。。\r\n"
            else:
                info = result[0] + "识别成功，FingerID=" + str(id) + "，得分：" + str(Score) + "\r\n"
            # roiImg1 = result[1].resize((int(imageWidth.value * 0.95), int(imageHeight.value * 0.95)))
            img = ImageTk.PhotoImage(result[1])  # 更新图片
            lableShowFingerImage.config(image=img)  # 显示最新一张图片
            ShowFingerImage=img
            bmpFingerImg=result[1]
            MessageText(info)

    # 1:N查重
    def onCompareIdentifyN(self):
        global isOpen, isRunning,ShowFingerImage,bmpFingerImg
        isRunning = False

        info = "请按压手指进行查重..\r\n"
        WMRAPI = WMRAPI_Dll()
        # images.config(image=img)  # 显示最新一张图片在图片框
        MessageText(info)
        result = self.meansDoVerify()  # 获取比对信息
        if result[1] != 0:  # 获取到指纹特征图片
            db = Sqlite3()
            templates = db.SelectAll()
            Score = {}
            for i in templates:
                # 进行比对
                Ver = WMRAPI.Verify(i[1])
                if Ver[0] == 0:
                    Score[i[0]] = Ver[1].value
                    Score = dict(sorted(Score.items(), key=lambda x: x[1], reverse=True))  # 排序
                if len(Score) > 10:  # 字典中超过十个元素时
                    Score.popitem()  # 将最后面的键对值弹出，即分数最小的弹出
                info = result[0] + "指纹识别成功！识别到以下" + str(len(Score)) + "个用户:\r\n"
            if Score == {}:  # 空字典
                info = result[0] + "未识别到用户！\r\n"
            else:
                for key in Score:
                    info = info + "FingerID=" + str(key) + "，得分：" + str(Score[key]) + "\r\n"
            # roiImg1 = result[1].resize((int(imageWidth.value * 0.95), int(imageHeight.value * 0.95)))
            img = ImageTk.PhotoImage(result[1])  # 更新图片
            lableShowFingerImage.config(image=img)  # 显示最新一张图片
            ShowFingerImage=img
            bmpFingerImg=result[1]
            MessageText(info)

    # 指纹特征
    def meansDoVerify(self):
        global isRunning,lableShowFingerImage,ShowFingerImage,bmpFingerImg
        size = 0
        isRunning = True
        WMRAPI = WMRAPI_Dll()
        while size == 0:
            # 获取指纹图像
            result = WMRAPI.GetImage()
            print("getImg",result[3].value)
            size=result[3].value
            flag = isRunning
            if flag == False:
                return [None, 0]
            if result[0] == 0:
                # 显示图像
                bmp = WMRAPI.RawToBMP(result[2],imageWidth,imageHeight)
                print(bmp)
                byte_stream = BytesIO(bmp[1].raw)  # 将二进制转为字节流
                roiImg = Image.open(byte_stream)  # 转成图片
                # roiImg1=roiImg.resize((int(imageWidth.value*0.95),int(imageHeight.value*0.95)))
                img = ImageTk.PhotoImage(roiImg)
                lableShowFingerImage.config(image=img)
                ShowFingerImage=img
                bmpFingerImg=roiImg

        # 提取指纹特征
        result = WMRAPI.Extract(bmp[1])
        if result[0] != 0:
            nr = "比对指纹特征提取失败，错误码：" + str(result[0]) + "\r\n"
            return [nr, 0]
        str1 = result[1]
        nr = "比对指纹特征数据:\r\n" + str1 + "\r\n"
        return [nr, roiImg]

    def SingleDelete(self):
        global gets,isRunning
        db=Sqlite3()
        db.Delete(gets)
        Fingerprint=FingerprintLibrary()
        Fingerprint.Select()
        isRunning=False

    def AllDelete(self):
        global isRunning
        db=Sqlite3()
        db.DeleteAll()
        Fingerprint=FingerprintLibrary()
        Fingerprint.Select()
        isRunning = False

    def systemLogInterface(self):
        global text
        print('查看系統日誌')
        MessageText("查看系統日誌\r\n")

        # 創建主視窗
        logRoot = tkinter.Tk()
        logRoot.title("系統日誌")

        # 創建 ScrolledText 元件
        scrolled_text = tkinter.scrolledtext.ScrolledText(logRoot, wrap=tkinter.WORD, width=75, height=65)
        scrolled_text.pack(padx=10, pady=10)

        # 設置初始文本
        scrolled_text.insert(tkinter.END, text.get("1.0", tkinter.END))

        # 啟動主迴圈
        logRoot.mainloop()

    def Stop(self):
        global isRunning, stop_all_threads
        stop_all_threads = True
        if isOpen == True:
            isRunning = False
            print("已停止..\r\n")
            MessageText("已停止..\r\n")
        else:
            print("指纹设备未打开\r\n")
            MessageText("指纹设备未打开\r\n")

    def Exit(self):
        global root,isOpen,isRunning
        if isOpen==True:
            WMRAPI = WMRAPI_Dll()
            WMRAPI.CloseDevice()
            isOpen=False
        isRunning=False
        root.destroy()


#endregion

#region auxiliaryMeans
class auxiliaryMeans:
    # 设置按钮是否可用
    def anConfig(self,root, state):
        button = ButtonGroup(root)
        for i in button:
            i.config(state=state)

    # 点击图片进行保存
    def my_label(self,event,bmp):
        fname = filedialog.asksaveasfilename(title=u'另存为', filetypes=[("BMP", ".bmp")])
        # 确认保存
        if fname != "":
            bmp.save(str(fname) + '.bmp', 'BMP')
            MessageText("保存成功\n")

    # 图标
    def icoImage(self):
        ico = b"AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAhRQUFoUUFJ+FFBTthRQU/4UTE+WFFBSLhBQUCoUUFAaFFBSDhRQU5YUUFP+FFBTnhBMTkYUUFBAAAAAAhRQUDIUUFNWFFBT/hRQU/4UUFP+FFBT/hRQU/4UUFLuFFBSxhRQU/4UUFP+FFBT/hRQU/4UUFP+FFBTPhRQUDIUTE3KFFBT/hRQU/4UUFP+FFBT/hRQU/4UUFP+FFBT/hRQU/4UUFP+FFBT/hRQU/4UUFP+FFBT/hRQU/4UUFIOFFBS/hRQU/4UUFP+FFBT/hRQU/4UUFP+FFBT/hRQU/4UUFP+FFBT/hRQU/4UUFP+FFBT/hRQU/4UUFP+FFBTbhRQU4YUUFP+FFBT/hRQU/4UUFP+FFBT/hRQU/4UUFP+FFBT/hRQU/4UUFP+FFBT/hRQU/4UUFP+FFBT/hRQU/YUUFOWFFBT/hRQU/4UUFP+FFBT/hRQUh4UUFP+FFBT/hRQU/4UUFP+FFBSVhRQU8YUUFP+FFBT/hRQU/4UUFP+FFBTlhRQU/4UUFP+FFBT/hRQU/4UUFHCFFBT/hRQU/4UUFP+FFBT/hRQUg4UUFOOFFBT/hRQU/4UUFP+FFBT/hRQU5YUUFP+FFBT/hRQU/4UUFP+FFBRwhRQU/4UUFP+FFBT/hRQU/4UUFIOFFBTjhRQU/4UUFP+FFBT/hRQU/4UUFOWFFBT/hRQU/4UUFP+FFBT/hRQUcIUUFP+FFBT/hRQU/4UUFP+FFBSDhRQU44UUFP+FFBT/hRQU/4UUFP+FFBTlhRQU/4UUFP+FFBT/hRQU/4UUFHCFFBT/hRQU/4UUFP+FFBT/hRQUg4UUFOOFFBT/hRQU/4UUFP+FFBT/hRQU5YUUFP+FFBT/hRQU/4UUFP+FFBRwhRQU/4UUFP+FFBT/hRQU/4UUFIOFFBTjhRQU/4UUFP+FFBT/hRQU/4UUFOWFFBT/hRQU/4UUFP+FFBT/hRQUcIUUFP+FFBT/hRQU/4UUFP+FFBSDhRQU44UUFP+FFBT/hRQU/4UUFP+FFBTlhRQU/4UUFP+FFBT/hRQU/4UUFHCFFBT/hRQU/4UUFP+FFBT/hRQUg4UUFOOFFBT/hRQU/4UUFP+FFBT/hRQU5YUUFP+FFBT/hRQU/4UUFP+FFBRwhRQU/4UUFP+FFBT/hRQU/4UUFIOFFBTjhRQU/4UUFP+FFBT/hRQU/4UUFOWFFBT/hRQU/4UUFP+FFBT/hRQUcIUUFP+FFBT/hRQU/4UUFP+FFBSDhRQU44UUFP+FFBT/hRQU/4UUFP+FFBTRhRQU94UUFPeFFBT3hRQU64UUFF6FFBT3hRQU94UUFPeFFBT3hRQUcIUUFNGFFBT3hRQU94UUFPeFFBTtwYMAAIABAACAAAAAAAAAAAAAAAAAAAAABAAAAAQAAAAEAAAABAAAAAQAAAAEAAAABAAAAAQAAAAEAAAABCAAAA=="
        tmp = open("winuim.ico", "wb+")
        tmp.write(base64.b64decode(ico))
        tmp.close()
        root.iconbitmap("winuim.ico")
        os.remove("winuim.ico")

    def RemoveBackground(self,foreimage):
        BG_color=np.array([0, 0, 0])
        foreimage = foreimage.convert('RGBA')
        foreimageg_array = np.array(foreimage)
        _BG_color = BG_color.reshape([-1, 1, 3])
        _BG_color = np.ones(foreimageg_array[:, :, :3].shape) * _BG_color

        # 对背景色亮度高于前景色的区域进行反相
        pos = np.where(np.max((foreimageg_array[:, :, :3] - _BG_color), axis=2) <= 0)
        _BG_color[pos] = 255 - _BG_color[pos]
        foreimageg_array[pos[0], pos[1], :3] = 255 - foreimageg_array[pos[0], pos[1], :3]

        # 透明度
        foreimageg_Alpha = np.max((foreimageg_array[:, :, :3] - _BG_color) / (255 + 1e-7 - _BG_color), axis=2)
        foreimageg_array[:, :, 3] = foreimageg_Alpha * 255
        foreimageg_Alpha = foreimageg_Alpha.reshape([foreimageg_array.shape[0], -1, 1])
        foreimageg_array[:, :, :3] = np.minimum(_BG_color + (foreimageg_array[:, :, :3] - _BG_color) / foreimageg_Alpha,255)
        # 恢复反相
        foreimageg_array[pos[0], pos[1], :3] = 255 - foreimageg_array[pos[0], pos[1], :3]

        res_img = Image.fromarray(foreimageg_array)
        return res_img

    # AES解码
    def AES_de(self,data, key):
        mode = AES.MODE_ECB  #ECB模式
        cryptos = AES.new(key=key, mode=mode)    #定义秘钥和模式
        plain_text = cryptos.decrypt(a2b_hex(data))  #解密
        clean_text=[]
        result = bytes.decode(plain_text)
        result=result.split("\r\n")
        for i in result:
            clean_text.append(re.sub(r"[\x00-\x1F\x7F]", "", i))
        return clean_text  #第一个元素是时间

    def compareTime(self,timestr):
        nowTime = datetime.datetime.now()
        endTime = datetime.datetime.strptime(timestr, "%Y%m%d%H%M%S")
        print(nowTime < endTime)
        return (nowTime < endTime)   #授权时间还没过则为真


    # 注册FingerID弹出框
    def Input_id(self):
        db = Sqlite3()
        sql = '''SELECT * FROM Fingers ORDER BY id DESC LIMIT 1;'''  # 获取id最大的数据
        if db.Select(sql) != None:
            FingerID = db.Select(sql)[0] + 1
        else:
            FingerID = 1
        # 输入整数
        entry_int = simpledialog.askinteger(title='设置FingerID', prompt='FingerID', initialvalue=FingerID)
        return entry_int
#endregion


#region 插入信息
def MessageText(result):
    global text
    text.insert(END, result)  # 将信息result插入信息框最末尾
    text.see(END)  # 将滚动条置于末尾
    text.update()  # 更新信息框
#endregion

#region Sqlite3
class Sqlite3:
    def __init__(self):
        self.Start()
        self.Create()
        self.Close()

    # 连接WMRAPI数据库，如果没有则会创建该数据库
    def Start(self, path='FPdata.db'):
        self.conn = sqlite3.connect(path)
        self.cursor = self.conn.cursor()

    # 创建数据表
    def Create(self):
        sql = '''create table if not exists Fingers(id int primary key not null,data text not null);'''  # 如果数据表Fingers不存在则创建数据表
        self.cursor.execute(sql)

    def Close(self):
        self.cursor.close()
        self.conn.close()

    # 插入数据表
    def Insert(self, id, data):
        try:
            self.Start()
            sql = '''insert into Fingers(id,data) values(?,?);'''  # 如果数据表Fingers不存在则创建数据表
            self.cursor.execute(sql, (id, data))
            self.conn.commit()
            self.Close()
            return 1
        except Exception as e:
            print('>> Insert Error:', e)
            return 0


    # 查询数据
    def Select(self, sql):
        self.Start()
        self.cursor.execute(sql)
        # 获取结果
        result = self.cursor.fetchone()
        self.Close()
        return result

    # 显示数据表
    def SelectAll(self):
        self.Start()
        sql = '''select * from Fingers;'''  # 如果数据表Fingers不存在则创建数据表
        self.cursor.execute(sql)
        # 获取结果集
        data_all = self.cursor.fetchall()
        self.Close()
        return data_all

    # 更新数据
    def Update(self,id,data):
        self.Start()
        sql = '''UPDATE Fingers SET data = ? WHERE id = ?;'''  # 如果数据表Fingers不存在则创建数据表
        self.cursor.execute(sql, (data, id))
        self.conn.commit()
        self.Close()
        return 1

    # 删除数据
    def Delete(self, id):
        self.Start()
        sql = '''DELETE FROM Fingers WHERE id = ?;'''  # 如果数据表Fingers不存在则创建数据表
        self.cursor.execute(sql, (id,))
        self.conn.commit()
        self.Close()
        return 1

    # 删除数据
    def DeleteAll(self):
        self.Start()
        sql = '''DELETE FROM Fingers;'''  # 如果数据表Fingers不存在则创建数据表
        self.cursor.execute(sql)
        self.conn.commit()
        self.Close()
        return 1
#endregion

#region FingerprintLibrary
class FingerprintLibrary:
    # 指纹库表格
    def __init__(self):
        self.Finger = ttk.Treeview(root, columns=("FingerID", "Template"), show="headings",height=10)  # #创建表格对象
        self.Finger.heading("FingerID", text="FingerID")  # 显示表头
        self.Finger.heading("Template", text="Template")
        self.Finger.column("FingerID", width=250)  # #设置列
        self.Finger.column("Template", width=450)
        self.Finger.place(relx=0.02, rely=0.75)
        self.Finger.bind('<ButtonRelease-1>', self.treeviewClick)  # 绑定单击离开事件===========

    # 获取当前点击行的值
    def treeviewClick(self,event):  # 单击
        global gets,getData
        for item in  self.Finger.selection():
            item_text = self.Finger.item(item, "values")
            gets=item_text[0]
            print(item_text[0])

    def Select(self):
        global gets
        db=Sqlite3()
        data = db.SelectAll()
        for i in data:
            self.Finger.insert("", i[0], values=(i[0], i[1].hex().upper()))
        gets=None
#endregion

#region WMRAPI_Dll
class WMRAPI_Dll:
    def OpenDevice(self):
        # 打开设备
        global HANDLE
        print(HANDLE)
        handle=pointer(HANDLE)
        result = WMdll.WM_OpenDevice(0, handle)  # 调用dll函数

        self.GetImageInfo()
        self.GetImageSize()
        return [result, HANDLE]

        # 获取图像信息

    def GetImageInfo(self):
        global imageWidth, imageHeight  # 使用全局变量
        WM_GetImageInfo = WMdll.WM_GetImageInfo
        WM_GetImageInfo.restype = c_int  # 设置函数的返回值类型
        WM_GetImageInfo.argtypes = (POINTER(c_int), POINTER(c_int))  # 设置函数的参数类型
        result = WM_GetImageInfo(imageWidth, imageHeight)
        print("指纹设备图像大小：", imageWidth.value, imageHeight.value)
        return [result, imageWidth.value, imageHeight.value]

    def GetImageSize(self):
        global ImageSize,rollImageWidth, rollImageHeight  # 使用全局变量
        size = FTRSCAN_IMAGE_SIZE(0, 0, 0)
        Size = pointer(size)
        ftrdll.ftrScanGetImageSize(HANDLE, Size)  #获取图像信息
        rollImageWidth = size.nWidth
        rollImageHeight = size.nHeight
        ImageSize = size.nImageSize
        print("指纹设备图像大小，Width：Height：ImageSize=%d：%d：%d" % (size.nWidth, size.nHeight, size.nImageSize))
    def GetSerialNumber(self):
        try:
            print(HANDLE)
            # result = WMdll2.WM_GetDeviceCount()
            # print("设备数量",result)
            DeviceSN = create_string_buffer(40)  # 创建40字节长的DeviceSN，且初始化为0
            result = WMdll.WM_GetSerialNumber(HANDLE, DeviceSN)
            SN = DeviceSN.raw.decode().strip(b'\x00'.decode())  # 将DeviceSN转换成字符串且去掉后面多余的\x00

            # print("result:",result)
            print("SN:",SN)
            return [result, HANDLE, SN]
        except Exception as e:
            print(e)

    #获取帧数时的图片
    def GetFrame(self):
        m_pBuffer=create_string_buffer(ImageSize)
        bRc=ftrdll.ftrScanGetFrame(HANDLE,m_pBuffer,None)
        # print("bRc:",bRc)
        # print("m_pBuffer:",m_pBuffer.raw)
        return [bRc,m_pBuffer]

    #开启滚动获取
    def ScanRollStart(self):
        bRc=ftrdll.ftrScanRollStart(HANDLE)
        print("bRc:",bRc)
        return bRc

    #获取滚动图片
    def RollGetImage(self):
        m_pBuffer=create_string_buffer(ImageSize)
        dwMilliseconds=5000
        bRc=ftrdll.ftrScanRollGetImage(HANDLE,m_pBuffer,dwMilliseconds)
        print("bRc:",bRc)
        # print("m_pBuffer:",m_pBuffer.raw)
        return [bRc,m_pBuffer]

    def RawToBMP(self,imageBuf,width,height):
        try:
            # BmpImageBuf = create_string_buffer(ImageSize+1078)  # 创建300000字节长的buf，且初始化为0
            BmpImageBuf = create_string_buffer(300000)  # 创建300000字节长的buf，且初始化为0
            size = c_int()  # int类型
            Size = pointer(size)  # 指针
            result = WMdll.WM_RawToBMP(imageBuf, width, height, BmpImageBuf, Size)
            # print("RawToBMP:",(BmpImageBuf.raw).hex()[:2156])
            return [result, BmpImageBuf, size]
        except Exception as e:
            print(e)

    # 获取指纹图像
    def GetImage(self):
        try:
            # print(int(imageWidth*imageHeight))
            imageBuf = create_string_buffer(300000)  # 创建300000字节长的buf，且初始化为0
            size = c_int()
            Size = pointer(size)  # 指针
            result = WMdll.WM_GetImage(HANDLE, 0, imageBuf, Size)
            return [result,HANDLE,imageBuf,size]
        except Exception as e:
            print(e)

    # 合成指纹模板
    def GenTemplateWithImage(self,ImgPtr):
        try:
            Image = (POINTER(c_char) * 3)()  # 定义指针数组
            template = create_string_buffer(2048)  # 创建2048字节长的指纹模板，且初始化为0
            Image[:] = [create_string_buffer(ImgPtr[i], 300000) for i in range(3)]  # 初始化指针数组，ImgPtr[i]内容为300000字节长
            size = c_int()
            Size = pointer(size)
            result = WMdll.WM_GenTemplateWithImage(Image, 3, imageWidth, imageHeight, template, Size)
            print("WM_GenTemplateWithImage:" + str(result), template, "Size:" + str(size))
            Template = template.raw[:].hex()  # 将获取的模板从字节转为16进制字符串
            Template = Template.upper()  # 字母全大小
            print(Template)
            return [result, Template, size, template]
        except Exception as e:
            print(e)

    # 提取指纹特征
    def Extract(self, imageBuf):
        # try:
        global feature

        size = c_int()
        Size = pointer(size)
        result = WMdll.WM_Extract(imageBuf, imageWidth, imageHeight, feature, Size)
        print("WM_Extract:" + str(result), feature, str(size))
        Feature = feature.raw[:size.value].hex()  # 将字节转为16进制字符串
        Feature = Feature.upper()  # 字母全大小
        return [result, Feature, size]

    # 指纹比对
    def Verify(self, templates):
        try:
            score = c_int()
            Score = pointer(score)
            template = create_string_buffer(templates)  # 创建2048字节长的指纹模板，且初始化为0
            print("template")
            print(template.raw.hex())
            print("feature")
            print(feature.raw.hex())

            result = WMdll.WM_Verify(template, feature, Score)
            print("WM_Verify:" + str(result), str(score))
            return [result, score]
        except Exception as e:
            print(e)
    def CloseDevice(self):
        global HANDLE
        print(HANDLE)
        result=WMdll.WM_CloseDevice(HANDLE)
        return result

    #切换图片背景为黑色
    def SetOptions(self):
        FTR_OPTIONS_INVERT_IMAGE=0x00000040
        ftrdll.ftrScanSetOptions(HANDLE,FTR_OPTIONS_INVERT_IMAGE,0)

        # 切换图片背景为黑色
        # ftrdll.ftrScanSetOptions(HANDLE,FTR_OPTIONS_INVERT_IMAGE,FTR_OPTIONS_INVERT_IMAGE)
#endregion

#region FTRSCAN_IMAGE_SIZE
class FTRSCAN_IMAGE_SIZE(ctypes.Structure):
    _fields_ = [
    ("nWidth", ctypes.c_int),
    ("nHeight", ctypes.c_int),
    ("nImageSize", ctypes.c_int),
    ]
#endregion

#region SaveFile



class SavePDFFile:
    def __init__(self):
        self.finger_images = []

    def create_pdf(self):

        fingers=left_fingers_text+right_fingers_text
        allFingers=[]

        # 檢查 'finger' 資料夾中是否有指定檔案，有的話就加入 allHand
        for label_text in fingers:
            image_path = os.path.join("fingerCache", f"{label_text}.bmp")
            if os.path.exists(image_path):
                allFingers.append(label_text)
                print(f"'{label_text}.bmp' 存在於 'fingerCache' 資料夾中")
                MessageText(f"'{label_text}.bmp' 存在於 'fingerCache' 資料夾中\r\n")
            else:
                print(f"警告：'{label_text}.bmp' 不存在於 'fingerCache' 資料夾中")
                MessageText(f"警告：'{label_text}.bmp' 不存在於 'fingerCache' 資料夾中\r\n")

        # 檢查指紋數量是否滿十枚
        if len(allFingers) < 10:
            warning_message = "指紋數量不足十枚，無法建立完整手指 PDF。"
            print(warning_message)
            MessageText(f"{warning_message}\r\n")

            # 顯示警告視窗
            tkinter.messagebox.showwarning("警告", warning_message)
            return

           

        # 使用 filedialog.asksaveasfilename 讓使用者選擇 PDF 的名稱
        pdf_filename = filedialog.asksaveasfilename(
            initialdir=os.getcwd(),
            title="儲存為",
            filetypes=[("PDF files", "*.pdf")],
            defaultextension=".pdf"
        )

        if not pdf_filename:
            print("請輸入 PDF 檔案名稱。")
            MessageText("請輸入 PDF 檔案名稱。\r\n")
            return

        # 使用 reportlab 來建立 PDF
        pdf_canvas = canvas.Canvas(pdf_filename)
        page_width, page_height = pdf_canvas._pagesize

        # 左手、右手
        for i, label in enumerate(allHand):
            text = label.cget("text")
            y_position = page_height * 0.25  # 設置在每頁上半部分的中心
            pdf_canvas.setFont('ChineseFont', 48)
            text_width = pdf_canvas.stringWidth(text, 'ChineseFont')  # 取得文字寬度
            pdf_canvas.drawString((page_width - text_width) / 2, y_position-60, text)  # 將 y_position 調整為文字的底線位置

            # 構建圖檔路徑
            image_path = os.path.join("fingerCache", f"{text}.bmp")

            # 轉換成 PhotoImage
            photo_image = SavePDFFile.load_image(image_path)

            # 調整插入 PDF 中的圖片，等比例縮放至寬度為 100
            desired_width = 400
            image_width, image_height = photo_image.width, photo_image.height
            scale_factor = desired_width / image_width
            scaled_width = desired_width
            scaled_height = int(image_height * scale_factor)
            x_position = (page_width - scaled_width) / 2
            pdf_canvas.drawInlineImage(photo_image, x_position, y_position, width=scaled_width, height=scaled_height)
            pdf_canvas.showPage()  # 換頁

        pdf_canvas.save()

        print(f"成功建立新的 PDF 檔案至 {pdf_filename}")
        MessageText(f"成功建立新的 PDF 檔案至 {pdf_filename}\r\n")

    @staticmethod
    def load_image(filename):
        """
        從檔案載入圖片，轉換為 PhotoImage
        """
        image = Image.open(filename)
        return image

#endregion

if __name__ == '__main__':
    # icoImage()
    UI()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
