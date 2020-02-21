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
import random

from PyQt5.QtGui import QPainter
from PyQt5.QtWidgets import QApplication, QLabel
from PyQt5.QtPrintSupport import QPrinter
from PyQt5.QtCore import QSizeF

widthPage = 300
heightPage = 100
printer_name = ""
portServer = 51003
version = '0.1'
app = Flask(__name__)
cors = CORS(app)
app.config['CORS_HEADERS'] = 'Content-Type'
random.seed(999999, version=2)


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
        # filename = tempfile.mktemp(".png")
        # filename = tempfile.NamedTemporaryFile(suffix='.png').name
        # filename = os.path.join(tempfile._get_default_tempdir(), next(tempfile._get_candidate_names()))+".png"
        filename = "Temp_Html_Images%s.png" % random.randrange(999999)
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
        # filename = tempfile.mktemp(".png")
        # filename = tempfile.NamedTemporaryFile(suffix='.png').name
        # filename = os.path.join(tempfile._get_default_tempdir(), next(tempfile._get_candidate_names()))+".png"
        filename = "Temp_Html_Images%s.png" % random.randrange(999999)
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


def print_text_qt(text="", printer_name="", widthPage=300, heightPage=100):
    requestMessage = {}
    try:
        if printer_name == "":
            printer_name = win32print.GetDefaultPrinter()
    except Exception:
        requestMessage["Error"] = "GetDefaultPrinter:%s" % traceback.format_exc()
        return requestMessage
    try:
        printer_name = "Brother QL-810W"
        app = QApplication(sys.argv)
        label = QLabel()
        # label.setText(
        #    '<img width="300"  height="100"  alt="" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAALQAAAB2CAIAAADujy7aAAABuUlEQVR42u3YwY7CIBRAUZj4/7/MLEgIFlqf1YwOPWdlsKlVrlDNqVNKSSnlnNvjqo6MR47PjuP9OSNniJw/foWRa4tc7d45n32tyDvtRyLjkWuIfErjGX4S7BAH4kAciANxIA7EgTgQB+JAHCAOxIE4EAfiQByIA3EgDsQB4kAciANxIA7EgTgQB+JAHCAOxIE4EAfiQByIA3EgDsSBOEAciANxIA7EgTgQB+JAHIgDxIE4EAfiQByIA3EgDsSBOEAciANxIA7EgTgQB+JAHIgDcYA4EAfiQByIA3EgDsSBOBAHiANxIA7EgTgQB+JAHIgDcYA4EAfiQByIg0/LpRSfAlYOxMGb3FbbJvPdRplzbo/b+CuD/Ugbnw6K4+vKOA6lTuQrg5tZr684HbStfO+a8alX/PvLEMdjwSmZHtZvDW12T8zxSmVc6IZ0utT3c1n72JvdcXw6stKesuAN6d56ML1tnE726W9/fxvrhvR/7zX9FI5LyHgruvD2ccVt5WCpiM9u5MiVtpKL3nOcWPaX/z3y4O0v8z43X9/Nf1bT/7WmR6bD/7sOinn2d5M4sK0gDsQB4uDQLyTS/Ojjiz5LAAAAAElFTkSuQmCC" />')
        label.setText(text)
        # label.resize(widthPage, heightPage)
        printer = QPrinter()
        # printer.setPaperSize(QSizeF(widthPage, heightPage), QPrinter.DevicePixel)
        printer.setPrinterName(printer_name)
        # printer.setPrinterName('Microsoft Print to PDF')
        # printer.setOutputFileName("C:\\!Delete\\!11\\33333.pdf")
        painter = QPainter()
        painter.begin(printer)
        painter.drawPixmap(-15, -15, label.grab())
        painter.end()
        # sys.exit(app.exec_())
    except Exception:
        requestMessage["Error"] = "GetDefaultPrinter:%s" % traceback.format_exc()
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
    if "Qt" in requestMessage:
        res = print_text_qt(requestMessage["Qt"], printer_name, widthPage, heightPage)
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
    #app.run(host='0.0.0.0', port=51003)
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(TestService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(TestService)
