import sys
import os
import win32com.client
import psutil
import pythoncom
import time

TARGET = ('EXCEL.EXE')
filepath = ('D:\\Dev\\Python\\convert\\test.xlsx')
a = str("2020-05-18")
try:
    pythoncom.CoInitialize()
    Excel = win32com.client.Dispatch("Excel.Application")
    Excel.Visible = 0
    wb1 = Excel.Workbooks.Open(u'D:\\Dev_autonomous\\Python\\BUP_BOT\\xlsx\\test.xlsx')
    sheet2 = wb1.ActiveSheet
    sheet2.Cells(2,7).value = int(15)
    sheet2.Cells(4,1).value = a
    wb1.ExportAsFixedFormat(0, u'D:\\Dev_autonomous\\Python\\BUP_BOT\\pdf\\test.pdf')
    wb1.Save()
    wb1.Close()
except Exception as e:
    print (e)
finally:
    print('finish')
    Excel.Quit()
    time.sleep(3)
    for proc in psutil.process_iter():
        if proc.name == TARGET:
            proc.kill