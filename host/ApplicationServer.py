import win32print
import win32ui
import win32con
from PIL import Image
from PIL import ImageWin
import tempfile
import imgkit
from flask import request
from urllib.parse import unquote
import os


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




def print_image(StrPrintHtml="", printer_name="",widthPage=300,heightPage=100):
    """
    Функция вывода текста на принтер
    :param StrPrintHtml: - Текст HTML
    :param printer_name: - имя принтера
    :return:
    """
    requestMessage = {}
    if printer_name == "":
        try:
            printer_name = win32print.GetDefaultPrinter()
        except:
            requestMessage["Error"] = "Error:No select printer"
            return requestMessage
    try:
        filename = tempfile.mktemp(".png")
        htmlStr = """<!DOCTYPE html><html><head><meta charset="utf-8"><title>Печать</title></head><body>%s</body></html>""" % StrPrintHtml
        options = {'width': widthPage, 'height': heightPage }
        imgkit.from_string(htmlStr, filename, options=options)
        img = Image.open(filename, 'r')
    except OSError as e:
        requestMessage["Error"] = "Error:create temp file %s %s" % (filename,  str(e))
        return requestMessage
    hdc = win32ui.CreateDC()
    hdc.CreatePrinterDC(printer_name)
    horzres = hdc.GetDeviceCaps(win32con.HORZRES)
    vertres = hdc.GetDeviceCaps(win32con.VERTRES)
    landscape = horzres > vertres
    if landscape:
        if img.size[1] > img.size[0]:
            print('Landscape mode, tall image, rotate bitmap.')
            img = img.rotate(90, expand=True)
    else:
        if img.size[1] < img.size[0]:
            print('Portrait mode, wide image, rotate bitmap.')
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
    return requestMessage

