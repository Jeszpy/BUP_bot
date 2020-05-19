import telebot
import constant
from datetime import date
import os
import psutil
import win32com.client
import threading
import sys
import pythoncom
from pythoncom import CoInitialize 
sys.coinit_flags = 0
bot = telebot.TeleBot(constant.API_TOKEN)

@bot.message_handler(content_types=['document'])    
def handle_file(message):
    try:
        chat_id = message.chat.id
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        src = u'D:\\Python\\Bot\\' + message.document.file_name
        with open(src, 'wb') as new_file:         # Сохраняем присланый от пользователя файл
            new_file.write(downloaded_file)
        try:
            pythoncom.CoInitialize()
            Excel = win32com.client.Dispatch("Excel.Application")
            print('Запускаю процесс Excel')
            Excel.Visible = 0
            today = str(date.today().strftime('%Y-%m-%d'))
            print('Узнаю дату')        #Дата!!!
            wb1 = Excel.Workbooks.Open(u'D:\\Python\\Bot\\Форма.xlsx')
            print('Открываю форму')
            sheet1 = wb1.ActiveSheet
            val1 = str(sheet1.Cells(1,2).value)
            val2 = str(sheet1.Cells(2,2).value)
            val3 = int(sheet1.Cells(3,2).value)
            val4 = int(sheet1.Cells(4,2).value)
            print('Принимаю значения')
            wb1.Save()
            wb1.Close()
            print('Закрыл форму')
            wb2 = Excel.Workbooks.Open(u'D:\\Python\\Bot\\Счет-фактура.xlsx')
            print('Открываю Счет-фактура.xlsx')
            sheet2 = wb2.ActiveSheet
            print('Передаю значения')
            sheet2.Cells(9,3).value = val1
            sheet2.Cells(10,3).value = val2
            sheet2.Cells(14,4).value = val3
            sheet2.Cells(16,4).value = int(val4/100)
            sheet2.Cells(4,1).value = today
            
            price1 = int(sheet2.Cells(14,6).value)        # Цена карточки
            price2 = int(sheet2.Cells(16,6).value)       # Цена паучи
            if int(val3) <= 50:            # Регулирование цены от кол-ва
                sheet2.Cells(14,6).value = 0.6           # Регулирование цены от кол-ва
            else:                           # Регулирование цены от кол-ва
                sheet2.Cells(14,6).value = 0.5         # Регулирование цены от кол-ва
            if int(val4) >= 1000:                 # Регулирование цены от кол-ва
                sheet2.Cells(16,6).value = 15                 # Регулирование цены от кол-ва
            else:                            # Регулирование цены от кол-ва
                sheet2.Cells(16,6).value = 20                    # Регулирование цены от кол-ва

            #cost1 = price1*val3
            #cost2 = price2*val4

            nds1 = int(sheet2.Cells(14,9).value)
            nds2 = int(sheet2.Cells(16,9).value)
            
            summ1 = int(sheet2.Cells(14,10).value)
            summ2 = int(sheet2.Cells(16,10).value)

            total_summ = summ1+summ2
            total_nds = nds1+nds2

            x1 = str('=СуммаПрописью(J18)')
            x2 = str('=СуммаПрописью(J20)')
            sheet2.Cells(21,3).value = x1
            sheet2.Cells(22,3).value = x2
            sheet2.Cells(18,10).value = total_summ
            sheet2.Cells(20,10).value = total_nds
            total_summ1 = str(sheet2.Cells(21,3).value)    # Пропись суммы
            total_nds1 = str(sheet2.Cells(22,3).value)      # Пропись НДС
            print(sheet2.Cells(21,3).value)          # Пропись суммы
            print(sheet2.Cells(22,3).value)      # Пропись НДС
            print(total_summ1)           # Пропись суммы
            print(total_nds1)          # Пропись НДС
            print('Передал значения')
    
            number = int(sheet2.Cells(2,6).value)
            while number < 10000:
                number += 1
                break
            print('Присваиваю №')
            sheet2.Cells(2,6).value = number
            '''if val3 == 0:
                sheet2.Rows(14).EntireRow.Hidden = True
            elif val3 != 0:
                sheet2.Rows(14).EntireRow.Hidden = False
            if val4 == 0:
                sheet2.Rows(16).EntireRow.Hidden = True
            elif val4 != 0:
                sheet2.Rows(16).EntireRow.Hidden = False'''
            number_name = str(number)
            print('Сохраняю Счет-фактура №'+number_name+', '+val1+'.pdf')
            wb2.ExportAsFixedFormat(0, 'D:\\Python\\Bot\\Счет-фактура №'+number_name+', '+val1+'.pdf', 'rb')
            print('Сохраняю Счет-фактура.xlsx')
            wb2.Save()
            #print('Сохранил Счет-фактуру №'+number_name+', '+val1+'.pdf')
            wb2.Close()
            Excel.Quit()
        except Exception as er:
            print(er)
        doc = open(u'D:\\Python\\Bot\\Счет-фактура №'+number_name+', '+val1+'.pdf', 'rb')
        bot.send_document(chat_id, doc)
        print('Отправил документ в ответ')
bot.polling(none_stop=True, interval=0)