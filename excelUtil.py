
import datetime
import os
import pythoncom
import sys
import time
import win32com.client

from xlConst import *
class xlColor:
    GRAY=0xe8e8e8
    BLUE=0xd59b5b
    LIGHT_BLUE_0=0xf0b000 #stand ligiht blue
    LIGHT_BLUE=0xf7ebdd  #80% light blue
    GOLD=0xc0ff
    LIGHT_GOLD=0xccf2ff  #80% light gold

class ExcelUtil():
    @staticmethod
    def newExcel():
        pythoncom.CoInitialize()
        excel=win32com.client.DispatchEx("excel.application")
        excel.Visible=True
        excel.DisplayAlerts=False
        return excel
    @staticmethod
    def openWkb(excel,f,readOnly=False):
        return excel.Workbooks.Open(f,ReadOnly=readOnly)

    def setAccountingFormat(self,rng):
        rng.NumberFormatLocal = '_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_ '

    def setAlignCenter(self,rng):
        rng.HorizontalAlignment=xlCenter
        rng.VerticalAlignment=xlCenter

    def setBackColor(self,rng,color):
        rng.Interior.Color=color
    def setBackColorBlue(self,rng):
        rng.Interior.Color=xlColor.BLUE

    def setBackColorGold(self,rng):
        rng.Interior.Color=xlColor.GOLD
    def setBackColorLightBlue0(self,rng):
        rng.Interior.Color=xlColor.LIGHT_BLUE_0
    def setBackColorLightBlue(self,rng):
        rng.Interior.Color=xlColor.LIGHT_BLUE
    def setBackColorLightGold(self,rng):
        rng.Interior.Color=xlColor.LIGHT_GOLD

    def setBorders(self,rng):
        for i in [xlEdgeLeft,xlEdgeRight,xlEdgeTop,xlEdgeBottom,xlInsideVertical,xlInsideHorizontal]:
            rng.Borders(i).LineStyle=xlContinuous
            rng.Borders(i).Weight=xlThin

    def getMaxRow(self,sht,col=None):
        if col:
            return sht.Cells(sht.Rows.Count,col).End(xlUp).Row
        else:
            return sht.UsedRange.Rows.Count-sht.UsedRange.Row+1

    def getMaxCol(self,sht,row=None):
        if row:
            return sht.Cells(row,sht.Columns.Count).End(xlToLeft).Column
        else:
            return sht.UsedRange.Columns.Count - sht.UsedRange.Column + 1

    def closeAllExcel(self):
        os.system('taskkill  /F /IM excel.exe /T')
    def mkDirs(self,dir):
        predir,_=os.path.split(dir)
        if not os.path.exists(predir):
            self.mkDirs(predir)
        else:
            os.mkdir(dir)

    def toFloat(self,f):
        try:
            return float(f)
        except:
            return 0.00
    def toPyDate(self, date):
        if isinstance(date,datetime.datetime):
            r:time.struct_time=time.strptime(f'{date.year}-{date.month}-{date.day}','%Y-%m-%d')
            return r
        return None
    def isDateTime(self,date):
        return isinstance(date,datetime.datetime)