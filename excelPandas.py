from excelUtil import *
from xlConst import *
from excelSht import *

import pandas as pd

class ExcelPandas(ExcelSht):
    KEY_BREAK = '_@_'
    SHT_NAME = 'UnTitle'

    def __init__(self):
        self.titleRow = None
        self.title = None
        self.headRow = None
        self.sht = None
        self.primaryKeys: [] = None  # 行数据的关键字段
        self.heads: [] = None

        self._headColDic: {} = {}
        self._keyRowDic: {} = {}

        self.data:pd.DataFrame=None

    @property
    def dataLength(self):
        return len(self.data)

    @property
    def dataWidth(self):
        return len(self.data.columns)

    def append(self, dic: dict):
        self.data=self.data.append(dic,ignore_index=True)

    def commit(self):
        sht=self.sht
        data=self.data.values
        headRow=self.headRow
        self.deleteContentRng()
        sht.Range(sht.Cells(headRow+1,1),sht.Cells(headRow+len(data),len(data[0]))).Value=data

    def getAllDic(self):
        data=self.data
        result = []
        for iRow in range(len(data)):
            dic = {}
            for head , iCol in self._headColDic.items():
                val=self.data.iloc[iRow,iCol]
                dic[head] = val
            result.append(dic)
        return result

    def getCol(self,head:str):
        return self._headColDic[head]

    def getKeysData(self, keys):
        key = self.KEY_BREAK.join([str(k) for k in keys])
        row = self._keyRowDic.get(key, None)
        if row!=None:
            dic = {}
            for colName, col in self._headColDic.items():
                dic[colName] = self.data.iloc[row,col]
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
    def getValue(self,row,col):
        return self.data.iloc[row,col]

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

        self._keyRowDic = {}
        if maxRow>self.headRow:
            lines=sht.Range(sht.Cells(self.headRow+1,1),sht.Cells(maxRow,maxCol)).Value
            data=pd.DataFrame(lines, columns=heads)
            self.data=data

            if self.primaryKeys:

                keyCols=[self._headColDic[item] for item in self.primaryKeys]

                for iRow in range(len(data)):
                    keys=[]
                    for iCol in keyCols:
                        val=data.iloc[iRow,iCol]
                        if val:
                            keys.append(str(val))
                    if keys:
                        key=self.KEY_BREAK.join(keys)
                        self._keyRowDic[key]=iRow
        else:
            self.data=pd.DataFrame([],columns=heads)

    def new(self, sht):
        sht.Name = self.SHT_NAME
        self.sht = sht
        self.setTitleRng()
        self.setHeadsRng()

        self.initSht(sht)

    def printData(self):
        dic=self.data.to_dict()
        for x ,v in dic.items():
            print(x,v)
    def setCategories(self,head,categories:[]=None):
        if categories:
            self.data[head]=pd.Categorical(self.data[head],categories=categories)
    def setKeysValue(self, keys, head, value):
        key = self.KEY_BREAK.join(keys)
        col=self.self._headColDic.get(head,None)
        row=self._keyRowDic.get(key,None)
        if row==None or col==None:
            return
        self.data.iloc[row, col] = value

    def setRowValue(self, row, head, value):
        col = self._headColDic.get(head, None)
        if col!=None:
            self.data.iloc[row,col] = value
    def sortValues(self,by:[],ascending=None):
        if ascending:
            self.data.sort_values(by=by,ascending=ascending,inplace=True)
        else:
            self.data.sort_values(by=by, inplace=True)

if __name__=='__main__':
    f=r"C:\Users\p1340814\Desktop\NCS 财务部\demo\record\项目记录表_2021.xlsx"
    t=ExcelPandas()
    t.primaryKeys='姓名,项目号,月份'.split(',')
    excel=t.newExcel()
    wkb=t.openWkb(excel,f,True)
    sht=wkb.Sheets(2)
    t.headRow=2
    t.initSht(sht)
    t.setCategories('月份','7月,5月,4月,6月,8月'.split(','))
    t.setRowValue(0,'姓名','fdasfdafdasfd')
    t.sortValues(['月份'])
    # t.data.sort_values(by=['月份'],ascending=[True],inplace=True)
    t.commit()

    # t.closeAllExcel()