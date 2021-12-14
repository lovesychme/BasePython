
class ExcelSht:
    def __init__(self):
        self.titleRow=None
        self.headRow=None
        self.sht=None
        self.headColDic=None
    @property
    def maxRow(self):
        return self.sht.UsedRange.Rows.Count - self.sht.UsedRange.Row + 1
    @property
    def maxCol(self):
        return self.sht.UsedRange.Columns.Count - self.sht.UsedRange.Column + 1

    def initExcelSht(self, sht):
        dic={}
        for i in range(1,self.maxCol+1):
            val=sht.Cells(self.headRow,i).Value
            if val:
                dic[val]=i
        self.headColDic=dic
        self.sht=sht

    def addData(self,dic:dict):
        row=self.maxRow+1
        for k,v in dic.items():
            if k in self.headColDic:
                self.sht.Cells(row,self.headColDic[k]).Value=v
