import yaml
import os

class Config():
    def __init__(self,file='config.yml',folderName='NCS_Config'):
        prefix, suffix = os.path.split(file)
        appdata = os.getenv('appdata')
        if appdata or not os.path.exists(prefix) or not os.access(prefix, os.W_OK):
            prefix = f'{appdata}\\{folderName}'
            self.mkDir(prefix)

        file=f'{prefix}\\{suffix}'

        self.configData=None
        self.file=file
        if os.path.exists(file):
            with open(file,'r') as f:
                self.configData=yaml.load(f, Loader=yaml.Loader)
        if self.configData==None:
            self.configData={}
    def __getKeyData(self,k):
        data=self.configData
        if '.' in k:
            ks=k.split('.')
            k=ks[-1]
            prefixs=ks[0:-1]
            for x in prefixs:
                if x not in data:
                    data[x]={}
                data=data[x]
        return data
    def mkDir(self,dir):
        if os.path.exists(dir):
            return
        prefix,_=os.path.split(dir)
        self.mkDir(prefix)
        os.mkdir(dir)
    def set(self,k,v):
        data = self.__getKeyData(k)
        data[k]=v
    def add(self,k,v):
        data = self.__getKeyData(k)
        if isinstance(data[k],list):
            data[k].append(v)
        else:
            data[k]=[data[k],v]
    def get(self,k,default=None):
        data = self.__getKeyData(k)
        if k in data:
            return data[k]
        else:
            return default
    def save(self):
        with open(self.file,'w') as f:
            yaml.dump(self.configData, f)
    def print(self):
        print(self.configData)
if __name__=="__main__":
    c=Config()
    c.set('ncs.name','nihao')
    c.set('ncs.age', 18)
    c.add('ncs.name','niyao')
    c.print()
    c.save()

