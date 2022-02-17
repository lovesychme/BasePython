import os,sys,time

os.popen(r'"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"  --remote-debugging-port=9222 --user-data-dir=C:\Users\p1340814' )
#--headless
# out=os.popen(r'"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"  --remote-debugging-port=9222')
# print(out.read())
# os.popen(r'"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"  --user-data-dir=C:\Users\p1340814')

# Browser.close


import requests as rq

# url=r'http://localhost:9222/json'
# r=rq.get(url)
# print(r.text)
# print(r.json())

# url=r'http://localhost:9222/devtools/page/A3AF96EDCF3D9CB03C66EF9A566EF38A'
# r=rq.get(url)
# print(r.status_code)
# print(r.text)
#
# url=r'http://localhost:9222/json'
# params={
#     'id':1,
#     'method':'Browser.close'
# }
# r=rq.post(url=url,data=params)
# print(r.text)


import pychrome
br=pychrome.Browser()
x=br.list_tab()
# tab=x[0]
tab=br.new_tab("https:www.baidu.com")
tab.start()

acc=tab.Accessibility
acc.enable()
acc.loadComplete=lambda :print('loadComplete')

dom =tab.DOM
dom.enable()
dom.documentUpdated=lambda :print('documentUpdated ')

net=tab.Network
net.enable()
def loadingFinished(*args,**kwargs):
    requestId = kwargs.get('requestId')
    frameId=kwargs.get('frameId','')
    print(f'loadingFinished,requestId:{requestId},frameId:{frameId}')
net.loadingFinished=loadingFinished

def requestWillBeSent(*args,**kwargs):
    requestId=kwargs.get('requestId')
    frameId = kwargs.get('frameId', '')
    print(f'requestWillBeSent,requestId:{requestId},frameId:{frameId}')
net.requestWillBeSent=requestWillBeSent

def responseReceived(*args,**kwargs):
    requestId = kwargs.get('requestId')
    frameId = kwargs.get('frameId', '')
    print(f'responseReceived,requestId:{requestId},frameId:{frameId}')
net.responseReceived=responseReceived


page=tab.Page
page.enable()
# page.addScriptToEvaluateOnNewDocument(source='alert("Hello")')
x=page.navigate(url="https://mail.163.com", _timeout=5)



x=dom.getDocument()
x=dom.getOuterHTML(nodeId=1)

# print(x)

while True:
    time.sleep(1)

