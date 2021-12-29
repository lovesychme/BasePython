from excelUtil import *
from xlConst import *
from excelSht import *

class ExcelData(ExcelSht):
    KEY_BREAK = '_@_'
    SHT_NAME = 'UnTitle'

    def __init__(self):
        self.titleRow = None
        self.title = None
        self.headRow = None
        self.sht = None
        self.primaryKeys: [] = None  # 行数据的关键字段
        self.heads: [] = None

        self._headColDic: {} = None
        self._keyRowDic: {} = None

        self.data:[]=None

    @property
    def dataLength(self):
        return len(self.data)

    @property
    def dataWidth(self):
        return len(self.data[0])

    def append(self, dic: dict):
        l=[None for x in range(len(self.heads))]
        for k, v in dic.items():
            if k in self._headColDic:
                l[self._headColDic[k]] = v
        self.data.append(l)

    def commit(self):
        sht=self.sht
        data=self.data
        headRow=self.headRow
        self.deleteContentRng()
        sht.Range(sht.Cells(headRow+1,1),sht.Cells(headRow+len(data),len(data[0]))).Value=data

    def getAllDic(self):
        data=self.data
        dic = []
        for iRow in range(len(data)):
            dic = {}
            for head , iCol in self._headColDic.items():
                val=self.data[iRow][iCol]
                dic[head] = val
            dic.append(dic)
        return dic

    def getCol(self,head:str):
        return self._headColDic[head]

    def getKeysData(self, keys):
        key = self.KEY_BREAK.join([str(k) for k in keys])
        row = self._keyRowDic.get(key, None)
        if row:
            dic = {}
            sht = self.sht
            for colName, col in self._headColDic.items():
                dic[colName] = self.data[row][col]
            return dic

    def getKeysRow(self, keys):
        key = self.KEY_BREAK.join([str(k) for k in keys])
        return self._keyRowDic.get(key, None)

    def getKeysValue(self, keys, head):
        key = self.KEY_BREAK.join(keys)
        if key in self._keyRowDic:
            return self.sht.Cells(self._keyRowDic[key], self._headColDic[head]).Value

    def getRowCell(self, row, head):
        col = self._headColDic.get(head, None)
        if col:
            return self.sht.Cells(row, col)

    def initSht(self, sht):
        self.sht = sht
        maxRow=self.maxRow
        maxCol=self.maxCol

        dic = {}
        heads=sht.Range(sht.Cells(self.headRow,1),sht.Cells(self.headRow,maxCol)).Value[0]
        for i in range(len(heads)):
            val=heads[i]
            if val:
                dic[val]=i
        self._headColDic=dic

        if maxRow>self.headRow:
            lines=sht.Range(sht.Cells(self.headRow+1,1),sht.Cells(maxRow,maxCol)).Value
            data=[]
            for line in lines:
                data.append(list(line))
            self.data=data
            if self.primaryKeys:
                self._keyRowDic={}
                keyCols=[self._headColDic[item] for item in self.primaryKeys]
                for iRow in range(len(self.data)):
                    keys=[]
                    for iCol in keyCols:
                        val=data[iRow][iCol]
                        if val:
                            keys.append(str(val))
                    if keys:
                        key=self.KEY_BREAK.join(keys)
                        self._keyRowDic[key]=iRow

    def new(self, sht):
        sht.Name = self.SHT_NAME
        self.sht = sht
        self.setTitleRng()
        self.setHeadsRng()

        self.initSht(sht)

    def setKeysValue(self, keys, head, value):
        key = self.KEY_BREAK.join(keys)
        if key not in self._keyRowDic:
            self.addKey(key)
        self.data(self._keyRowDic[key], self._headColDic[head]).Value = value

    def setRowValue(self, row, head, value):
        col = self._headColDic.get(head, None)
        if col:
            self.data[row][col] = value
