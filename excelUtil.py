import sys,os
from xlConst import *
class xlColor:
    GRAY=0xe8e8e8

class ExcelUtil():
    def setBorders(self,rng):
        for i in [xlEdgeLeft,xlEdgeRight,xlEdgeTop,xlEdgeBottom,xlInsideVertical,xlInsideHorizontal]:
            rng.Borders(i).LineStyle=xlContinuous
            rng.Borders(i).Weight=xlThin
    def setAlignCenter(self,rng):
        rng.HorizontalAlignment=xlCenter
        rng.VerticalAlignment=xlCenter

    def setBackColor(self,rng,color):
        rng.Interior.Color=color

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
