import cv2
import sys
from tkinter import E
from PyQt5.QtWidgets import QApplication,QDialog,QFileDialog,QTableWidgetItem
from PyQt5.QtGui import QPalette,QStandardItemModel,QStandardItem
from PyQt5.QtCore import Qt,pyqtSlot,QProcess
from ui_mainWindow import Ui_Dialog
import os
import csv
import win32api as wp
import time,win32com.client,win32con
import win32clipboard as wcb
import win32gui as wg
from PIL import Image, ImageGrab
import asyncio
from threading import Thread


class QmyDialog(QDialog):
    def __init__(self, parent = None):
        super().__init__(parent)
        self.ui = Ui_Dialog()
        self.ui.setupUi(self)
        self.ui.btnSetGamePath.clicked.connect(self.open_file)
        self.ui.btnReadAccount.clicked.connect(self.read_account)
        self.ui.btnStart.clicked.connect(self.initGame)
        self.ui.btnStop.clicked.connect(self.closeWindow)
        self.path = ""
        self.accountPath = ""
        self.accountNum = 0
        self.account = []
        self.passwd = []
        self.log=True
        self.gameWndSize = [1280,720]
        # 登陆界面各元素位置
        self.userNamePos = [150,100]
        self.passWordPos = [150,134]
        self.rememberPos = [125,163]
        self.loginBtnPos = [200,200]
        self.steamBtnPos = [30,19]
        self.setBtnPos = [25,145]
        self.panelPos = [35,127]
        self.newsPos = [220,470]
        self.confirmBtnPos = [575,563]
        self.libBtnPos = [150,53]
        # gameBtnPos = [500,300]
        self.searchBtnPos = [100,184]
        self.listGamePos = [100,245]
        self.startBtnPos = [380,430]
        self.preset_val = 0.6
        self.wait_between_steps = 1
        self.accountGameStatus = True
        self.closeNews = False
        try:
            with open('info.ini', 'r') as f:
                text_lines = f.readlines()
                for l in text_lines:
                    if l.find("[gamePath]")>= 0:
                        self.path = l.split("=")[1].strip().replace('\n', '').replace('\t', '').replace('\r', '').strip()
                        print(self.path)
                        self.ui.textGamePath.setText(self.path)
                    elif l.find("[accountPath]")>= 0:
                        print(l)
                        self.accountPath = l.split("=")[1].strip().replace('\n', '').replace('\t', '').replace('\r', '').strip()
                        self.ui.textAccountPath.setText(self.accountPath)
        except :
            pass
        if self.accountPath != "" :
            self.ui.textAccountPath.setText(self.accountPath)
                            
            rowIndex,colIndex =0,0
            for row in csv.reader(open(self.accountPath,'r')):
                item = QTableWidgetItem()    
                self.ui.tableAccountList.insertRow(rowIndex)
                print(row[0],row[1])
                item.setText(row[0])
                self.account.append(row[0])
                
                self.ui.tableAccountList.setItem(rowIndex, 0, item)
                item = QTableWidgetItem()
                item.setText(row[1])
                self.ui.tableAccountList.setItem(rowIndex, 1, item)
                self.passwd.append(row[1])
                
                rowIndex +=1
                self.accountNum += 1
            # self.ui.tableAccountList.setModel(self.sm)    
            self.gameStatusChange("账号读取成功，共加载"+str(self.accountNum)+"个账号信息")
            
            print(self.account,self.passwd)
    
    def getHWND(self,WindowName,gap,repos=1):
        wndHandle = wg.FindWindow(0,WindowName)
        while wndHandle == 0:
            time.sleep(gap)
            wndHandle = wg.FindWindow(0,WindowName)
        # 窗口置顶
        time.sleep(1)
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('%')            
        if repos:
            wg.SetWindowPos(wndHandle,0,0,0,self.gameWndSize[0],self.gameWndSize[1],win32con.SWP_SHOWWINDOW)
        else:
            l,t,r,b = wg.GetWindowRect(wndHandle)
            wg.SetWindowPos(wndHandle,0,l,t,r-l,b-t,win32con.SWP_SHOWWINDOW)
        if self.log:
            print(WindowName, " Handle : ", wndHandle)
        return wndHandle
    
    def input(self,pos,content,baseL=0,baseT=0):
        wp.SetCursorPos((baseL+pos[0], baseT+pos[1]))
        time.sleep(0.1)
        wp.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        wp.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        wp.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        wp.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(0.5)
        wcb.OpenClipboard()
        wcb.EmptyClipboard()
        wcb.SetClipboardData(wcb.CF_UNICODETEXT,content)
        wp.keybd_event(0x11,0,0,0)
        wp.keybd_event(0x56,0,0,0)
        wp.keybd_event(0x11,0,win32con.KEYEVENTF_KEYUP,0)
        wp.keybd_event(0x56,0,win32con.KEYEVENTF_KEYUP,0)
        wcb.CloseClipboard()
        time.sleep(0.5)

    def click(self,pos,baseL=0,baseT=0):    
        wp.SetCursorPos((baseL+pos[0], baseT+pos[1]))
        time.sleep(0.5)
        wp.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        wp.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(0.5)

    def gameStatusChange(self,str):
        self.ui.labelGameStatus.setText(str)

    def initGame(self,str):
        if self.path != "" and self.accountNum > 0:
            wp.ShellExecute(0, 'open', self.path, '', '', 1)
            self.gameStatusChange("游戏启动中...")
            self.ui.btnReadAccount.setEnabled(False)
            self.ui.btnSetGamePath.setEnabled(False)
            self.ui.btnStart.setEnabled(False)
            thread = Thread(target = self.startGame)
            thread.start()
            # self.startGame()
        elif self.accountNum == 0:
            self.gameStatusChange("未加载账号信息...")
        else:
            self.gameStatusChange("未设置有效游戏路径...")
        

    
    def open_file(self):   
        if self.path == "":  
            fileName,flieType = QFileDialog.getOpenFileName(self,"选取文件",os.getcwd(),"exe file(*.exe)")
        else:
            fileName,flieType = QFileDialog.getOpenFileName(self,"选取文件",self.path.replace("/steam.exe",""),"exe file(*.exe)")
        if  fileName != "":
            self.path = fileName    
            self.ui.textGamePath.setText(self.path)
            self.gameStatusChange("设置游戏路径成功")            


    def read_account(self):
        if self.accountPath == "":
            fileName,flieType = QFileDialog.getOpenFileName(self,"选取文件",os.getcwd(),"CSV file(*.CSV)")
        else:
            fileName,flieType = QFileDialog.getOpenFileName(self,"选取文件",self.accountPath,"CSV file(*.CSV)")
        # print(fileName)
        if fileName != "":
            self.accountPath = fileName
            self.ui.textAccountPath.setText(self.accountPath)            
            # item = QTableWidgetItem()                    
            rowIndex,colIndex =0,0
            for row in csv.reader(open(self.accountPath,'r')):
                item = QTableWidgetItem()    
                self.ui.tableAccountList.insertRow(rowIndex)
                print(row[0],row[1])
                item.setText(row[0])
                self.account.append(row[0])
                
                self.ui.tableAccountList.setItem(rowIndex, 0, item)
                item = QTableWidgetItem()
                item.setText(row[1])
                self.ui.tableAccountList.setItem(rowIndex, 1, item)
                self.passwd.append(row[1])
                
                rowIndex +=1
                self.accountNum += 1
            # self.ui.tableAccountList.setModel(self.sm)    
            self.gameStatusChange("账号读取成功，共加载"+str(self.accountNum)+"个账号信息")
            print(self.account,self.passwd)
        else:
            self.gameStatusChange("请选择正确的文件格式")

    def picCompare(self,w_img_path,d_img_path,returnLoc = 0,thresValue = 220):
        w_image = cv2.imread(w_img_path)
        w_image = cv2.cvtColor(w_image,cv2.COLOR_BGR2GRAY)
        r,w_image = cv2.threshold(w_image,thresValue,255,cv2.THRESH_BINARY)
        d_image = cv2.imread(d_img_path)    
        d_image = cv2.cvtColor(d_image,cv2.COLOR_BGR2GRAY)
        r,d_image  = cv2.threshold(d_image,thresValue,255,cv2.THRESH_BINARY)
        res = cv2.matchTemplate(w_image,d_image,cv2.TM_CCOEFF_NORMED)
        min_val,max_val,min_loc,max_loc = cv2.minMaxLoc(res)
        if max_val > 0.9:
            if returnLoc :
                return max_loc
            return True
        return False

    def checkAccountGameStatus(self):
        self.accountGameStatus
        return True

    def startGame(self):
              
        loginWnd = self.getHWND("Steam 登录",1,0)
        # 获取窗口位置
        l,t,r,b = wg.GetWindowRect(loginWnd)
        while l == 0:
            l,t,r,b = wg.GetWindowRect(loginWnd)
        print(l,t,r,b)
        self.ui.labelGameStatus.setText("获取登录界面，输入用户名...")
        for ac in range(self.accountNum):
            self.input(self.userNamePos,self.account[ac],l,t)
            self.input(self.passWordPos,self.passwd[ac],l,t)
            self.click(self.loginBtnPos,l,t)
            
            # 2、进入主界面
            mainWnd = self.getHWND("Steam",1)
            self.ui.labelGameStatus.setText("登录成功，打开主界面...")
            wg.SetWindowPos(mainWnd,0,0,0,self.gameWndSize[0],self.gameWndSize[1],win32con.SWP_SHOWWINDOW)
            time.sleep(5)
            if self.closeNews:
                self.click(self.steamBtnPos)
                self.click(self.setBtnPos)
                # 3、进入设置界面
                setWnd = self.getHWND("设置",1,0)
                sl,st,sr,sp = wg.GetWindowRect(setWnd)
                while sl == 0:
                    sl,st,sr,sp = wg.GetWindowRect(setWnd)
                print(sl,st,sr,sp)
                self.click(self.panelPos,sl,st)
                self.click(self.newsPos,sl,st)
                self.click(self.confirmBtnPos,sl,st)

            # 4、返回主界面，在库中搜索游戏，并开始
            keyword = "BATTLEGROUNDS"
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.SendKeys('%')
            wg.SetForegroundWindow(mainWnd)
            self.click(self.libBtnPos)
            self.input(self.searchBtnPos,keyword)
            self.click(self.listGamePos)
            self.click(self.startBtnPos)            
            self.ui.labelGameStatus.setText("游戏加载中...")
            # 5、进入游戏界面，可以开始游戏
            gameWnd = self.getHWND("绝地求生 ",1)
            time.sleep(1)
            accountGameCount = 0
            while self.checkAccountGameStatus:
                self.ui.tableAccountList.setItem(0,2,QTableWidgetItem(str(accountGameCount)))
                startGameFound = 0
                countTime = 0
                while startGameFound == 0 :
                    wg.SetWindowPos(gameWnd,0,0,0,self.gameWndSize[0],self.gameWndSize[1],win32con.SWP_SHOWWINDOW)
                    time.sleep(1)
                    countTime+=1
                    window_img = ImageGrab.grab((0,0,self.gameWndSize[0],self.gameWndSize[1]))
                    window_img.save("./com/screenshot.png")
                    dst_img = "./dst/startGame.png"
                    d_img = cv2.imread(dst_img)
                    h,w = d_img.shape[:2]
                    window_img = "./com/screenshot.png"                       
                    print("startGameFound:  ",countTime,"s")
                    p = 0
                    p = self.picCompare(window_img,dst_img,1)
                    if p: 
                        self.ui.labelGameStatus.setText("开始游戏，等待进入游戏...")
                        startGameFound = 1
                        posX,posY = int(p[0]+w/2),int(p[1]+h/2)
                        self.click([posX,posY],0,0)
                gameReadyFound = 0
                countTime = 0
                while gameReadyFound == 0 :
                    wg.SetWindowPos(gameWnd,0,0,0,self.gameWndSize[0],self.gameWndSize[1],win32con.SWP_SHOWWINDOW)
                    time.sleep(1)
                    countTime += 1
                    window_img = ImageGrab.grab((280,500,620,800))
                    window_img.save("./com/gameReady.png")
                    dst_img = "./dst/gameReady.png"
                    window_img = "./com/gameReady.png"
                    print("gameReadyFound:  waiting ",countTime, " s ")
                    p = 0
                    p = self.picCompare(window_img,dst_img,1)            
                    # p= (max_loc[0]+w,max_loc[1]+h)
                    # w_img = cv2.rectangle(w_img,max_loc,p,(0,255,0),2)18626886786
                    if p:
                        self.ui.labelGameStatus.setText("游戏加载完成，倒计时...")
                        print("gameReadyFound")
                        gameReadyFound = 1
                countTo15 = 0
                countTime = 0
                while countTo15 == 0:
                    wg.SetWindowPos(gameWnd,0,0,0,self.gameWndSize[0],self.gameWndSize[1],win32con.SWP_SHOWWINDOW)
                    time.sleep(0.1)
                    countTime += 0.1
                    window_img = ImageGrab.grab((280,500,620,800))
                    window_img.save("./com/countDownScreen.png")
                    dst_img = "./dst/countdown.png"
                    window_img = "./com/countDownScreen.png"
                    p = 0
                    p = self.picCompare(window_img,dst_img,1)
                    print("countTo15:  waiting ",countTime," s ")
                    # p= (max_loc[0]+w,max_loc[1]+h)
                    # w_img = cv2.rectangle(w_img,max_loc,p,(0,255,0),2)18626886786
                    if p:
                        self.ui.labelGameStatus.setText("倒计时时间到...")
                        print("countTo15")
                        countTo15 = 1
                        wp.keybd_event(0x1B,0,0,0)
                        wp.keybd_event(0x1B,0,win32con.KEYEVENTF_KEYUP,0)
                escPressed = 0
                countTime = 0
                while escPressed == 0:
                    wg.SetWindowPos(gameWnd,0,0,0,self.gameWndSize[0],self.gameWndSize[1],win32con.SWP_SHOWWINDOW)
                    time.sleep(0.5)
                    countTime += 1
                    window_img = ImageGrab.grab((0,0,self.gameWndSize[0],self.gameWndSize[1]))
                    window_img.save("./com/exitGame.png")
                    dst_img = "./dst/exitGame.png"
                    window_img = "./com/exitGame.png"
                    d_img = cv2.imread(dst_img)
                    h,w = d_img.shape[:2]            
                    p = 0
                    p = self.picCompare(window_img,dst_img,1,190)
                    if p: 
                        self.ui.labelGameStatus.setText("退出游戏中...")
                        escPressed = 1
                        posX,posY = int(p[0]+w/2),int(p[1]+h/2)
                        self.click([posX,posY])
                exitConfirm = 0
                while exitConfirm == 0:
                    wg.SetWindowPos(gameWnd,0,0,0,self.gameWndSize[0],self.gameWndSize[1],win32con.SWP_SHOWWINDOW)
                    time.sleep(0.5)
                    countTime += 1
                    window_img = ImageGrab.grab((0,0,self.gameWndSize[0],self.gameWndSize[1]))
                    window_img.save("./com/exitConfirm.png")
                    dst_img = "./dst/exitConfirm.png"
                    window_img = "./com/exitConfirm.png"
                    d_img = cv2.imread(dst_img)
                    h,w = d_img.shape[:2]
                    print("exitConfirm:  waiting ",countTime," s")
                    p = self.picCompare("./com/exitConfirm.png","./dst/exitConfirm.png",1)
                    if p:
                        exitConfirm = 0
                        posX,posY = int(p[0]),int(p[1])
                        print(posX,posY)
                accountGameCount += 1
                self.ui.labelGameStatus.setText("刷机完成第",accountGameCount,"次")

        
        
        
        self.ui.btnReadAccount.setEnabled(True)
        self.ui.btnSetGamePath.setEnabled(True)
        self.ui.btnStart.setEnabled(True)
        
        pass
    
    def closeWindow(self):
        with open("info.ini","w") as f:
            f.write("[gamePath]="+self.path+"\n")
            f.write("[accountPath]="+self.accountPath)
            f.close()
        app.exit()
        
    
    


if __name__ == "__main__":
    app = QApplication(sys.argv)
    form = QmyDialog()
    form.show()
    sys.exit(app.exec())
    
