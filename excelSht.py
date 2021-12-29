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

        self._headColDic:{}= None
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
            if k in self._headColDic:
                self.sht.Cells(row, self._headColDic[k]).Value=v
    def autoFitColumns(self, maxWidth=None):
        sht=self.sht
        for iCol in range(1, self.maxCol + 1):
            sht.Columns(iCol).EntireColumn.AutoFit()
            if sht.Columns(iCol).ColumnWidth > 11.88:
                sht.Columns(iCol).ColumnWidth = 11.88

        if maxWidth:
            for iCol in range(1, self.maxCol + 1):
                if sht.Columns(iCol).ColumnWidth > maxWidth:
                    sht.Columns(iCol).ColumnWidth = maxWidth


    def deleteContentRng(self):
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
        return self._headColDic[head]

    def getKeysRow(self, keys):
        key = self.KEY_BREAK.join([str(k) for k in keys])
        return self._keyRowDic.get(key, None)

    def getKeysData(self, keys):
        key = self.KEY_BREAK.join([str(k) for k in keys])
        row=self._keyRowDic.get(key, None)
        if row:
            dic={}
            sht=self.sht
            for colName,col in self._headColDic.items():
                dic[colName]=sht.Cells(row,col).Value
            return dic
    def getKeysValue(self, keys, head):
        key=self.KEY_BREAK.join(keys)
        if key in self._keyRowDic:
            return self.sht.Cells(self._keyRowDic[key], self._headColDic[head]).Value

    def getRowCell(self,row,head):
        col=self._headColDic.get(head, None)
        if col:
            return self.sht.Cells(row,col)

    def initSht(self, sht):
        self.sht = sht
        maxRow=self.maxRow
        maxCol=self.maxCol

        dic = {}
        heads=sht.Range(sht.Cells(self.headRow,1),sht.Cells(self.headRow,maxCol)).Value[0]
        for i in range(len(heads)):
            val=heads[i]
            if val:
                dic[val]=i+1
        self._headColDic=dic

        if self.primaryKeys:
            self._keyRowDic={}
            keyCols=[self._headColDic[item] for item in self.primaryKeys]
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
        self.setHeadsRng()

        # 填写Title
        self.setTitleRng()
        self.initSht(sht)

    def setKeysValue(self, keys, head, value):
        key = self.KEY_BREAK.join(keys)
        if key not in self._keyRowDic:
            self.addKey(key)
        self.sht.Cells(self._keyRowDic[key], self._headColDic[head]).Value = value

    def setRowValue(self,row,head,value):
        col=self._headColDic.get(head, None)
        if col:
            self.sht.Cells(row,col).Value=value

    def setValueColor(self,rng,value,color=None):
        rng.Value=value
        rng.Interior.Color=color

    def setHeadsRng(self):
        sht=self.sht
        for i,s in enumerate(self.heads):
            sht.Cells(self.headRow,i+1).Value=s
            sht.Columns(i+1).EntireColumn.AutoFit()
        maxCol=i+1

        rng=sht.Range(sht.Cells(self.headRow,1),sht.Cells(self.headRow,maxCol))
        self.setAlignCenter(rng)
        self.setBackColor(rng,xlColor.GRAY) #灰色
        self.backColor=xlColor.GRAY
        self.setBorders(rng)

    def setTitleRng(self, rng=None, title=None, fontSize=12, fontBold=True):
        sht=self.sht

        if not rng:
            rng = sht.Range(sht.Cells(self.titleRow, 1), sht.Cells(self.titleRow, len(self.heads)))

        if not title:
            if self.title:
                title=self.title
            else:
                title=sht.Name

        #设置Title格式
        self.setBorders(rng)
        rng.Merge()
        self.setAlignCenter(rng)
        rng.Font.Size = fontSize
        rng.Font.Bold = fontBold
        rng.Value=title
