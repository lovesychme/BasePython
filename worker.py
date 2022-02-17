from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5 import QtCore

class Worker(QThread):
    msgSin = QtCore.pyqtSignal(object)
    endSin = QtCore.pyqtSignal(object)
    def __init__(self,obj,funcName,args:list=None,kwargs:{}=None,parent=None):
        super(Worker, self).__init__(parent)
        if not args:
            args=tuple()
        if not kwargs:
            kwargs={}
        self.obj=obj
        self.args=args
        self.kwargs=kwargs
        self.func=getattr(obj,funcName)

    def run(self,ignoreException=True):
        obj=self.obj
        if hasattr(obj,'msgSin'):
            obj.msgSin=self.msgSin
        if ignoreException:
            try:
                result=self.func(*self.args,**self.kwargs)
            except Exception as e:
                self.msgSin.emit(str(e))
                self.endSin.emit(str(e))
                return
        else:
            result = self.func(*self.args, **self.kwargs)
        if hasattr(obj,'errMsg') and obj.errMsg:
            self.msgSin.emit(obj.errMsg)
        self.endSin.emit(result)