import sys
import os
from datetime import datetime
import time
from flask import Flask, request
import win32com.client
import psutil
from num2words import num2words
import threading
import pythoncom
from pythoncom import CoInitialize
import yagmail

TARGET = ('EXCEL.EXE')

app = Flask(__name__)

@app.route('/', methods=['POST', 'GET'])
def form():
    if request.method == 'GET':
        return '''<!doctype html>
                <html>
                <head>
                    <title>Форма</title>
                    <meta charset="utf-8">
                    <style type="text/css">
                            p {
                                font-weight: bold;
                                font-size: 24px; 
                            }
                            body {
                                background: #598bff;
                                position: absolute;
                                top: 50%;
                                left: 50%;
                                margin-right: -50%;
                                transform: translate(-50%, -50%)
                            }
                    </style>
                </head>
                <body>
                    <h1>Форма для выписки счёт-фактуры</h1>
                    <div>
                        <form method="post">
                            <p>Полное название организации</p>
                            <textarea placeholder="ОАО 'Пример'" cols="40" rows="3" name="org_name"></textarea><br/>   
                            <p>Реквизиты</p><br/>
                            <textarea placeholder="Юридический и почтовый адрес, УНП, банковские реквизиты." name="org_requisites" cols="80" rows="7"></textarea><br/>
                            <p>Количество карточек:</p>
                            <input type="text" placeholder="0 - ထ шт." name="cards"><br/>
                            <p>Введите кол-во паучей:</p>
                            <input type="text" placeholder="0 - ထ шт." name="pauchs"><br/>
                            <p>Введите адрес электронной почты для обратной связи</p>
                            <input type="email" placeholder="example@gmail.com" name="email">
                            <br/>
                            <br/>
                            <button type="submit">Отправить</button>
                        </form>
                    </div>
                </body>
                </html>'''
    elif request.method == 'POST':      
        org_name = str((request.form['org_name']))
        org_requisites = str((request.form['org_requisites']))
        cards = int((request.form['cards']))
        pauchs = int((request.form['pauchs']))
        to_email = str((request.form['email']))
        try:
            pythoncom.CoInitialize()
            Excel = win32com.client.Dispatch("Excel.Application")
            Excel.Visible = 0
            current_date = str(datetime.now().date())
            wb2 = Excel.Workbooks.Open(u'D:\\Dev\\Python\\BUP_bot\\xlsx\\test.xlsx')
            sheet2 = wb2.ActiveSheet
            sheet2.Cells(9,3).value = org_name
            sheet2.Cells(10,3).value = org_requisites
            sheet2.Cells(14,4).value = cards
            sheet2.Cells(16,4).value = pauchs
            sheet2.Cells(4,1).value = current_date
            
            price1 = int(sheet2.Cells(14,6).value)        # Цена карточки
            price2 = int(sheet2.Cells(16,6).value)       # Цена паучи
            if cards <= 100:            # Регулирование цены от кол-ва
                sheet2.Cells(14,6).value = 0.7           # Регулирование цены от кол-ва
            else:                           # Регулирование цены от кол-ва
                sheet2.Cells(14,6).value = 0.6         # Регулирование цены от кол-ва
            if pauchs >= 150:                 # Регулирование цены от кол-ва
                sheet2.Cells(16,6).value = 17                 # Регулирование цены от кол-ва
            else:                            # Регулирование цены от кол-ва
                sheet2.Cells(16,6).value = 23                    # Регулирование цены от кол-ва

            
            summ_total = int(sheet2.Cells(18,10).value)*100
            summ_nds = int(sheet2.Cells(20,10).value)*100



            total_str = str(num2words(summ_total, to = 'currency', currency = 'RUB', separator='', lang='ru'))
            nds_str = str(num2words(summ_nds, to = 'currency', currency = 'RUB', separator='', lang='ru'))


            sheet2.Cells(21,3).value = total_str
            sheet2.Cells(22,3).value = nds_str

    
            number = int(sheet2.Cells(2,7).value)
            while number < 10000:
                number += 1
                break

            sheet2.Cells(2,7).value = number
            if cards == 0:
                sheet2.Rows(14).EntireRow.Hidden = True
            elif cards != 0:
                sheet2.Rows(14).EntireRow.Hidden = False
            if pauchs == 0:
                sheet2.Rows(16).EntireRow.Hidden = True
            elif pauchs != 0:
                sheet2.Rows(16).EntireRow.Hidden = False
            number_name = str(number)

            wb2.ExportAsFixedFormat(0, u'D:\\Dev\\Python\\BUP_bot\\pdf\\Счёт-фактура №'+number_name+'от'+current_date+'.pdf')
            wb2.Save()
            wb2.Close()

        except Exception as er:
            print(er)
        finally:
            Excel.Quit()
            time.sleep(0.5)
            for proc in psutil.process_iter():
                if proc.name == TARGET:
                    proc.kill
            time.sleep(0.25)
            pdf = ('D:\\Dev\\Python\\BUP_bot\\pdf\\Счёт-фактура №'+number_name+'от'+current_date+'.pdf')
            yag = yagmail.SMTP('belukrprom.bot@gmail.com','belprom1')
            contents = [
                "Добрый день!",
                "Счёт-фактура прикреплена.", pdf
            ]

            yag.send(to_email, 'ОДО "Белукропром"', contents)
    return "Форма отправлена"

if __name__ == '__main__':
    app.run(port=8080, host='127.0.0.1')



