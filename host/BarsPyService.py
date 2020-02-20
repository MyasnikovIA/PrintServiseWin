# https://tutorialmore.com/questions-1202301.htm
# sys.stdout = sys.stderr = open(os.devnull, 'w')
import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
import sys
from flask import Flask
from flask_cors import CORS, cross_origin
import json
import os
import win32print
import win32ui
import win32con
from PIL import Image
from PIL import ImageWin
import tempfile
import imgkit
from flask import request
from urllib.parse import unquote
import traceback

widthPage = 300
heightPage = 100
printer_name = ""
portServer = 51003
version = '0.1'
app = Flask(__name__)
cors = CORS(app)
app.config['CORS_HEADERS'] = 'Content-Type'


@app.errorhandler(404)
def not_found(error):
    return "no service", 404


def parseHead():
    """
     Чтение входных пользовательских аргументов в заголовке запроса
    """
    requestMessage = {}
    for rec in request.headers:
        if "X-My-" in rec[0]:
            message = unquote(unquote(request.headers.get(rec[0])))
            key = rec[0][5:len(rec[0])]
            requestMessage[key] = message
    return requestMessage


def get_print_list():
    """
    функция получения списка принтеров установленных в системе
    :return:
    """
    requestMessage = {}
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
    requestMessage["Printers"] = [printer[2] for printer in printers]
    return requestMessage


def print_image(img, printer_name):
    """процедура печати картинки"""
    requestMessage = {}
    try:
        if printer_name == "":
            printer_name = win32print.GetDefaultPrinter()
    except Exception:
        requestMessage["Error"] = "GetDefaultPrinter:%s" % traceback.format_exc()
        return requestMessage
    hdc = win32ui.CreateDC()
    hdc.CreatePrinterDC(printer_name)
    horzres = hdc.GetDeviceCaps(win32con.HORZRES)
    vertres = hdc.GetDeviceCaps(win32con.VERTRES)
    landscape = horzres > vertres
    if landscape:
        if img.size[1] > img.size[0]:
            img = img.rotate(90, expand=True)
    else:
        if img.size[1] < img.size[0]:
            img = img.rotate(90, expand=True)
    img_width = img.size[0]
    img_height = img.size[1]
    if landscape:
        # we want image width to match page width
        ratio = vertres / horzres
        max_width = img_width
        max_height = (int)(img_width * ratio)
    else:
        # we want image height to match page height
        ratio = horzres / vertres
        max_height = img_height
        max_width = (int)(max_height * ratio)
    # map image size to page size
    hdc.SetMapMode(win32con.MM_ISOTROPIC)
    hdc.SetViewportExt((horzres, vertres));
    hdc.SetWindowExt((max_width, max_height))
    # offset image so it is centered horizontally
    offset_x = (int)((max_width - img_width) / 2)
    offset_y = (int)((max_height - img_height) / 2)
    hdc.SetWindowOrg((-offset_x, -offset_y))
    hdc.StartDoc('Result')
    hdc.StartPage()
    dib = ImageWin.Dib(img)
    dib.draw(hdc.GetHandleOutput(), (0, 0, img_width, img_height))
    hdc.EndPage()
    hdc.EndDoc()
    hdc.DeleteDC()


def htmlurl_to_image(UrlPath="", printer_name="", widthPage=300, heightPage=100):
    """
    Функция вывода HTML на принтер
    :param UrlPath: - адрес запроса
    :param printer_name: - имя принтера
    :return:
    """
    requestMessage = {}
    try:
        filename = tempfile.mktemp(".png")
        imgkit.from_url(UrlPath, filename, options={'width': widthPage, 'height': heightPage})
    except Exception:
        requestMessage["Error"] = "create temp file %s %s" % (filename, traceback.format_exc())
        return requestMessage
    requestMessage = print_local_file(filename, printer_name)
    try:
        os.remove(filename)
    except  Exception:
        print("Remove file from temp :(%s)  %s" % (filename, traceback.format_exc()))
    return requestMessage


def html_to_image(StrPrintHtml="", printer_name="", widthPage=300, heightPage=100):
    """
    Функция вывода текста на принтер
    :param StrPrintHtml: - Текст HTML
    :param printer_name: - имя принтера
    :return:
    """
    requestMessage = {}
    try:
        filename = tempfile.mktemp(".png")
        imgkit.from_string(
            """<!DOCTYPE html><html><head><meta charset="utf-8"><title>Печать</title></head><body>%s</body></html>""" % StrPrintHtml,
            filename, options={'width': widthPage, 'height': heightPage})
    except Exception:
        requestMessage["Error"] = "create temp file %s %s" % (filename, traceback.format_exc())
        return requestMessage
    requestMessage = print_local_file(filename, printer_name)
    try:
        os.remove(filename)
    except  Exception:
        print("Remove file from temp :(%s)  %s" % (filename, traceback.format_exc()))
    return requestMessage


def print_local_file(filename, printer_name=""):
    """
    печать локального файла на принтер
    """
    requestMessage = {}
    try:
        img = Image.open(filename, 'r')
    except Exception:
        requestMessage["Error"] = "Open file from temp :(%s)  %s" % (filename, traceback.format_exc())
        return requestMessage
    try:
        print_image(img, printer_name)
    except Exception:
        requestMessage["Error"] = "Print image : %s" % traceback.format_exc()
        return requestMessage
    return requestMessage


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
        res = html_to_image(requestMessage["Print"], printer_name, widthPage, heightPage)
        return json.dumps(res), 200
    if "Printurl" in requestMessage:
        res = htmlurl_to_image(requestMessage["Printurl"], printer_name)
        return json.dumps(res), 200
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
BarsPySend({"Print":"<h1>Привет Мир-HelloWorld</h1>","widthPage":300,"heightPage":100,"PrinterName":"Brother QL-810W"},function(dat){console.log(dat);})
BarsPySend({"Print":"<h1>Привет Мир-HelloWorld</h1>"+Date(Date.now()).toString()}) // отправека на печать без получения ответа
BarsPySend({"PrintUrl":"http://127.0.0.1/sprint.png" },function(dat){console.log(dat);}) // Печать сайта по URL адресу
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
     # app.run(host='0.0.0.0', port=5000)
     if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(TestService)
        servicemanager.StartServiceCtrlDispatcher()
     else:
        win32serviceutil.HandleCommandLine(TestService)
