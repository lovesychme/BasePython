#_*_ coding:utf-8 _*_
import win32api,win32com.client,win32gui,win32con
import os
import pythoncom
import re
import time, threading

class JKWind():
    def __init__(self,h=None):
        self.h=h
    def __str__(self):
        c=win32gui.GetClassName(self.h)
        t=win32gui.GetWindowText(self.h)
        return "hand={},{}, class={}, text={}".format(hex(self.h),self.h, c, t)
    def print(self):
        print(self.__str__())
        return self
    def attachSaveAs(self):
        li = []
        win32gui.EnumWindows(lambda h, para: para.append(h), li)
        for h in li:
            c = win32gui.GetClassName(h)
            t = win32gui.GetWindowText(h)
            if '另存为' in t:
                self.h=h
                return True
        return False

    def enumFunc(self,h,para):
        if win32gui.GetParent(h)==para[1]:
            para[0].append(h)
    def enumChilds(self,h=None,ind=0):
        li = []
        if h==None: h=self.h
        if h != None:
            c=win32gui.GetClassName(h)
            t=win32gui.GetWindowText(h)
            for i in range(1,ind):
                print('|   ',end='')
            try:
                print("hand={},  class={}, text={}".format(hex(h), c, t))
            except:
                print("hand={},  class={}, text=ERROR".format(hex(h), c))
            win32gui.EnumChildWindows(h, self.enumFunc, (li, h))
        else:
            win32gui.EnumWindows(self.enumFunc, (li, h))
        for x in li:
            self.enumChilds(x,ind+2)


    def childCl(self, className='',index=1):
        li=[]
        win32gui.EnumChildWindows(self.h,lambda h,para:para.append(h),li)
        for x in li:
            c = win32gui.GetClassName(x)
            t = win32gui.GetWindowText(x)
            if c==className:
                if index<=1:
                    return JKWind(x)
                else:
                    index-=1
        return None


    def parent(self):
        return JKWind(win32gui.GetParent(self.h))
    def childTx(self,text='',index=1):
        li=[]
        win32gui.EnumChildWindows(self.h,lambda h,para:para.append(h),li)
        for x in li:
            c = win32gui.GetClassName(x)
            t = win32gui.GetWindowText(x)
            if text in t:
                if index<=1:
                    return JKWind(x)
                else:
                    index-=1
        return None
    def childTxCl(self,text='', className='',index=1):
        li=[]
        win32gui.EnumChildWindows(self.h,lambda h,para:para.append(h),li)
        for x in li:
            c = win32gui.GetClassName(x)
            t = win32gui.GetWindowText(x)
            if text in t and c==className:
                if index<=1:
                    return JKWind(x)
                else:
                    index-=1
        return None
    @classmethod
    def waitTopWindowTx(cls,text='',timeOut=30):
        tryCount=timeOut
        while True:
            li=[]
            win32gui.EnumWindows(lambda h,para:para.append(h),li)
            for x in li:
                c = win32gui.GetClassName(x)
                t = win32gui.GetWindowText(x)
                if text in t:
                    return JKWind(x)
            tryCount-=1
            if tryCount==0:
                return None
            time.sleep(1)
    @classmethod
    def waitTopWindowTxCl(cls,text='',className='',timeOut=30):
        tryCount=timeOut
        while True:
            li=[]
            win32gui.EnumWindows(lambda h,para:para.append(h),li)
            for x in li:
                try:
                    c = win32gui.GetClassName(x)
                    t = win32gui.GetWindowText(x)
                    if text in t and className==c:
                        return JKWind(x)
                except:
                    pass
            tryCount-=1
            if tryCount==0:
                return None
            time.sleep(1)
    @classmethod
    def waitTopWindowCl(cls,className='',timeOut=30):
        tryCount=timeOut
        while True:
            li=[]
            win32gui.EnumWindows(lambda h,para:para.append(h),li)
            for x in li:
                c = win32gui.GetClassName(x)
                t = win32gui.GetWindowText(x)
                if className==c:
                    return JKWind(x)
            tryCount-=1
            if tryCount==0:
                return None
            time.sleep(1)
    def click(self):
        win32api.PostMessage(self.h,win32con.WM_LBUTTONDOWN,0,0)
        time.sleep(0.5)
        win32api.PostMessage(self.h, win32con.WM_LBUTTONUP, 0, 0)
        return self
    @classmethod
    def moveTo(self,x,y):
        ox, oy = win32api.GetCursorPos()
        ix = (x - ox) / 100
        iy = (y - oy) / 100
        for i in range(100):
            a = int(ox + ix * (i + 1))
            b = int(oy + iy * (i + 1))
            win32api.SetCursorPos((a, b))
            time.sleep(0.01)
    @classmethod
    def clickPos(self,x,y):
        self.moveTo(x,y)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

    @classmethod
    def clickDoublePos(self,x,y):
        self.moveTo(x,y)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

    def click_by_mouse(self):
        rect=win32gui.GetWindowRect(self.h)
        x=int((rect[0]+rect[2])/2)
        y=int((rect[1]+rect[3])/2)
        win32api.SetCursorPos((x,y))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0,0,0)
        time.sleep(0.2)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        return self
    def getText(self):
        size=win32gui.SendMessage(self.h, win32con.WM_GETTEXTLENGTH, 0, 0)+1
        buff=win32gui.PyMakeBuffer(size)
        win32api.SendMessage(self.h, win32con.WM_GETTEXT, size, buff)
        address, length = win32gui.PyGetBufferAddressAndLen(buff[:-1])
        text = win32gui.PyGetString(address, length)
        return text
    def setText(self,text):
        win32gui.SendMessage(self.h,win32con.WM_CHAR,'a',0)
        time.sleep(0.1)
        win32gui.SendMessage(self.h ,win32con.WM_SETTEXT, 0, text)
        return self
    def key_return(self):
        win32gui.PostMessage(self.h,win32con.WM_KEYDOWN,win32con.VK_RETURN,0)
        time.sleep(0.2)
        win32gui.PostMessage(self.h, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        return self
    def pastText(self,text):
        self.setClickBoardText(text)
        self.clickWindow(self.h)
        return self
    def typeChar(self,character):
        win32api.keybd_event(ord(character),0,0,0)
        time.sleep(0.2)
        win32api.keybd_event(ord(character), 0, win32con.KEYEVENTF_KEYUP, 0)
        return self
    def typeKeyCode(self,keyCode):
        win32api.keybd_event(keyCode,0,0,0)
        time.sleep(0.2)
        win32api.keybd_event(keyCode, 0, win32con.KEYEVENTF_KEYUP, 0)
        return self
    def typeWords(self,text):
        for n in text:
            self.typeKeyCode(ord(n))
        return self
    def wait(self ,waitTime):
        time.sleep(waitTime)
        return self
    def moveCursor(self):
        rect=win32gui.GetWindowRect(self.h)
        x=int(((rect[0]+rect[2])/2)*65536/1920)
        y=int(((rect[1]+rect[3])/2)*65536/1080)
        win32api.mouse_event(int(win32con.MOUSEEVENTF_ABSOLUTE | win32con.MOUSEEVENTF_MOVE),x,y,0,0)
        return self
    def setForeWindow(self):
        try:
            win32api.keybd_event(win32con.VK_CONTROL,0,0,0)
            time.sleep(0.3)
            win32api.keybd_event(win32con.VK_CONTROL, 0,win32con.KEYEVENTF_KEYUP , 0)

            # win32gui.ShowWindow(self.h, win32con.SW_RESTORE)
            # win32gui.ShowWindow(self.h, win32con.SW_SHOW)

            win32gui.SetForegroundWindow(self.h)
        except:
            pass
        return self
    def close(self):
        win32gui.PostMessage(self.h,win32con.WM_CLOSE,0,0)
        return self

if __name__=='__main__':
    wind=JKWind()
    wind=wind.waitTopWindowTx('Activation')
    wind.enumChilds()
    wind.close()
