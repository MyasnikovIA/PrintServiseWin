# https://tutorialmore.com/questions-1202301.htm
# sys.stdout = sys.stderr = open(os.devnull, 'w')
import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
import sys
from flask import Flask, request
from flask_cors import CORS, cross_origin
import json

widthPage = 300
heightPage = 100
printer_name = ""
portServer = 51003
version = '0.1'

from ApplicationServer import *

app = Flask(__name__)
cors = CORS(app)
app.config['CORS_HEADERS'] = 'Content-Type'


@app.errorhandler(404)
def not_found(error):
    return "no service", 404

@app.route("/")
@cross_origin()
def requestFun():
    """
    Функция обработки входящих команд с сервера
    :return:
    """
    global printer_name
    global widthPage
    global heightPage
    print(request.host[:9])
    if request.host[:9] != "127.0.0.1":
        if request.host[:9] != "localhost":
            return "no service", 404
    requestMessage = parseHead()
    if "WidthPage" in requestMessage:
        widthPage = requestMessage.get("WidthPage")
    if "HeightPage" in requestMessage:
        heightPage = requestMessage.get("HeightPage")
    if "PrinterName" in requestMessage:
        printer_name = requestMessage.get("PrinterName")
    if "Print" in requestMessage:
        try:
            res = print_image(requestMessage.get("Print"), printer_name, widthPage, heightPage)
            return json.dumps(res), 200
        except:
            requestMessage["Error"] = "%s" % sys.exc_info()[0]
            return json.dumps(requestMessage), 200
    if "Getprinterlist" in requestMessage:
        return json.dumps(get_print_list()), 200
    if "Message" in requestMessage:
        if '[GetPrinterList]' in requestMessage["Message"]:
            return json.dumps(get_print_list()), 200
    if "Version" in requestMessage:
            requestMessage["Version"] = version
    return json.dumps(requestMessage), 200



class TestService(win32serviceutil.ServiceFramework):
    _svc_name_ = 'BarsPyLocalService'
    _svc_display_name_ = 'BarsPy local service'
    _svc_description_ = """
Сервис предназначен для доступа к локальным ресурсам рабочей станции (Принтеры , сканеры , и.т.д),
через крос-доменные запросов на хоси 127.0.0.1 (localhost)
URL:  http://127.0.0.1:%s/
Пример обращение к сервису через JS:
BarsPySend= function(messageObject,FunCallBack ){
    var host = "http://127.0.0.1:51003/";
    var cspBindRequestCall = new XMLHttpRequest();
    cspBindRequestCall.open('GET',host, true);
    if (typeof FunCallBack === 'function'){ 
         cspBindRequestCall.onreadystatechange = function() {
         if (this.readyState == 4 && this.status == 200) {
            if (typeof FunCallBack === 'function'){
		    try {
				FunCallBack(JSON.parse(decodeURIComponent(this.responseText)));
			} catch (err) {
			    FunCallBack(decodeURIComponent(this.responseText));
			}
          }
         };
       };
    }
    cspBindRequestCall.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
    if (typeof messageObject == "object"){
        for (var prop in messageObject) {
           cspBindRequestCall.setRequestHeader("X-My-"+prop, encodeURI(messageObject[prop]));
        }
    } else{
        cspBindRequestCall.setRequestHeader("X-My-message", encodeURI(messageObject));
    }
    cspBindRequestCall.send();
    return cspBindRequestCall; 
}
BarsPySend({"GetPrinterList":1},function(dat){console.log(dat);}) // получить список принтеров установленных в системе
BarsPySend({"Print":"<h1>Привет Мир-HelloWorld</h1>","widthPage":300,"heightPage":100,"PrinterName":"Microsoft XPS Document Writer"},function(dat){console.log(dat);})
BarsPySend({"Print":"<h1>Привет Мир-HelloWorld</h1>"}) // отправека на печать без получения ответа
    """ % portServer
    _args = None

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)
        self._args = args

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE, servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        self.main(self._args)

    def main(self, args):
        # Вставить инсталляцию плагина
        global portServer
        app.run(host='0.0.0.0', port=portServer)


if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(TestService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(TestService)
