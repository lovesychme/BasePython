from  excelUtil import *
from  xlConst import *

class ExcelSht(ExcelUtil):
    KEY_BREAK='_@_'
    SHT_NAME='UnTitle'
    def __init__(self):
        self.titleRow=None
        self.headRow=None
        self.sht=None
        self.headColDic=None
        self.primaryKeys:[]=None   #行数据的关键字段
        self.heads:[]=None
    @property
    def maxRow(self):
        return self.sht.UsedRange.Rows.Count - self.sht.UsedRange.Row + 1
    @property
    def maxCol(self):
        return self.sht.UsedRange.Columns.Count - self.sht.UsedRange.Column + 1
    def getCol(self,head:str):
        return self.headColDic[head]
    def initSht(self, sht):
        self.sht = sht
        dic={}
        for i in range(1,self.maxCol+1):
            val=sht.Cells(self.headRow,i).Value
            if val:
                dic[val]=i
        self.headColDic=dic

        if self.primaryKeys:
            keyCols=[self.headColDic[item] for item in self.primaryKeys]
            for iRow in range(self.headRow+1, self.maxRow):
                key=self.KEY_BREAK.join([sht.Cells(iRow,iCol).Value for iCol in keyCols])
                self.keyRowDic[key]=iRow


    def get_all_data(self):
        sht=self.sht
        data=[]
        for iRow in range(self.headRow,self.maxRow+1):
            dic={}
            for iCol in range(1,self.maxCol+1):
                head=sht.Cells(self.headRow,iCol).Value
                val=sht.Cells(iRow,iCol).Value
                dic[head]=val
            data.append(dic)
        return data

    def deleteRow(self,rowNumber):
        self.sht.Rows(rowNumber).Delete(Shift=xlUp)
    def deleteRngXlUp(self,rng):
        rng.Delete(Shift=xlUp)
    def setValueColor(self,rng,value,color=None):
        rng.Value=value
        rng.Interior.Color=color

    def writeTitle(self, rng, title, fontSize=12, fontBold=True):
        #设置Title格式
        self.setBorders(rng)
        rng.Merge()
        self.setAlignCenter(rng)
        rng.Font.Size = fontSize
        rng.Font.Bold = fontBold
        rng.Value=title

    def writeHeads(self,heads):
        sht=self.sht
        for i,s in enumerate(heads):
            sht.Cells(self.headRow,i+1).Value=s
            sht.Columns(i+1).EntireColumn.AutoFit()
        maxCol=i+1

        rng=sht.Range(sht.Cells(self.headRow,1),sht.Cells(self.headRow,maxCol))
        self.setAlignCenter(rng)
        self.setBackColor(rng,xlColor.GRAY) #灰色
        self.backColor=xlColor.GRAY
        self.setBorders(rng)

    def new(self, sht):
        sht.Name = self.SHT_NAME
        self.sht = sht
        # 填写Heads
        self.writeHeads(self.heads)

        # 填写Title
        rng = sht.Range(sht.Cells(self.titleRow, 1), sht.Cells(self.titleRow, self.maxCol))
        self.writeTitle(rng, sht.Name)
        self.initSht(sht)

    def getItem(self,keys,item):
        key=self.KEY_BREAK.join(keys)
        if key in self.keyRowDic:
            return self.sht.Cells(self.keyRowDic[key],self.headColDic[item]).Value

    def setKeyItem(self,keys,item,value):
        key = self.KEY_BREAK.join(keys)
        if key not in self.keyRowDic:
            self.addKey(key)
        self.sht.Cells(self.keyRowDic[key], self.headColDic[item]).Value = value

    def append(self, dic:dict):
        row=self.maxRow+1
        for k,v in dic.items():
            if k in self.headColDic:
                self.sht.Cells(row,self.headColDic[k]).Value=v