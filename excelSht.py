from  excelUtil import *
from  xlConst import *

class ExcelSht(ExcelUtil):
    KEY_BREAK='_@_'
    SHT_NAME='UnTitle'
    def __init__(self):
        self.titleRow=None
        self.title=None
        self.headRow=None
        self.sht=None
        self.primaryKeys:[]=None   #行数据的关键字段
        self.heads:[]=None

        self.__headColDic:{}= None
        self._keyRowDic:{}=None
    @property
    def maxRow(self):
        return self.sht.UsedRange.Rows.Count - self.sht.UsedRange.Row + 1
    @property
    def maxCol(self):
        return self.sht.UsedRange.Columns.Count - self.sht.UsedRange.Column + 1

    def append(self, dic:dict):
        row=self.maxRow+1
        for k,v in dic.items():
            if k in self.__headColDic:
                self.sht.Cells(row, self.__headColDic[k]).Value=v
    def deleteData(self):
        sht=self.sht
        if self.maxCol<=self.headRow:
            return
        rng=sht.Rows(f'{self.headRow+1}:{self.maxRow}')
        self.deleteRngXlUp(rng)
    def deleteRow(self,rowNumber):
        self.sht.Rows(rowNumber).Delete(Shift=xlUp)
    def deleteRngXlUp(self,rng):
        rng.Delete(Shift=xlUp)

    def getAllDic(self):
        sht=self.sht
        data=[]
        for iRow in range(self.headRow+1,self.maxRow+1):
            dic={}
            for iCol in range(1,self.maxCol+1):
                head=sht.Cells(self.headRow,iCol).Value
                val=sht.Cells(iRow,iCol).Value
                dic[head]=val
            data.append(dic)
        return data

    def getCol(self,head:str):
        return self.__headColDic[head]

    def getKeysRow(self, keys):
        key = self.KEY_BREAK.join([str(k) for k in keys])
        return self._keyRowDic.get(key, None)

    def getKeysData(self, keys):
        key = self.KEY_BREAK.join([str(k) for k in keys])
        row=self._keyRowDic.get(key, None)
        if row:
            dic={}
            sht=self.sht
            for colName,col in self.__headColDic.items():
                dic[colName]=sht.Cells(row,col).Value
            return dic
    def getKeysValue(self, keys, head):
        key=self.KEY_BREAK.join(keys)
        if key in self._keyRowDic:
            return self.sht.Cells(self._keyRowDic[key], self.__headColDic[head]).Value

    def getRowCell(self,row,head):
        col=self.__headColDic.get(head,None)
        if col:
            return self.sht.Cells(row,col)

    def initSht(self, sht):
        self.sht = sht
        dic={}
        for i in range(1,self.maxCol+1):
            val=sht.Cells(self.headRow,i).Value
            if val:
                dic[val]=i
        self.__headColDic=dic

        if self.primaryKeys:
            self._keyRowDic={}
            keyCols=[self.__headColDic[item] for item in self.primaryKeys]
            for iRow in range(self.headRow+1, self.maxRow+1):
                keys=[]
                for iCol in keyCols:
                    val=sht.Cells(iRow,iCol).Value
                    if val:
                        keys.append(str(val))
                if keys:
                    key=self.KEY_BREAK.join(keys)
                    self._keyRowDic[key]=iRow
    def new(self, sht):
        sht.Name = self.SHT_NAME
        self.sht = sht
        # 填写Heads
        self.setHeadsRng(self.heads)

        # 填写Title
        rng = sht.Range(sht.Cells(self.titleRow, 1), sht.Cells(self.titleRow, self.maxCol))
        if self.title:
            self.setTitleRng(rng, self.title)
        else:
            self.setTitleRng(rng, sht.Name)
        self.initSht(sht)

    def setKeysValue(self, keys, head, value):
        key = self.KEY_BREAK.join(keys)
        if key not in self._keyRowDic:
            self.addKey(key)
        self.sht.Cells(self._keyRowDic[key], self.__headColDic[head]).Value = value

    def setRowValue(self,row,head,value):
        col=self.__headColDic.get(head,None)
        if col:
            self.sht.Cells(row,col).Value=value

    def setValueColor(self,rng,value,color=None):
        rng.Value=value
        rng.Interior.Color=color

    def setHeadsRng(self, heads):
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

    def setTitleRng(self, rng, title, fontSize=12, fontBold=True):
        #设置Title格式
        self.setBorders(rng)
        rng.Merge()
        self.setAlignCenter(rng)
        rng.Font.Size = fontSize
        rng.Font.Bold = fontBold
        rng.Value=title
