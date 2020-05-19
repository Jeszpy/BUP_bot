
# imports!
from datetime import date
import sys
from datetime import datetime
from constant import *
import win32com.client
from num2words import num2words


# values!


# !!! Сумма НДС, Сумма с НДС !!!
a = 499.99

b = str(num2words(a, to = 'currency', currency = 'RUB', separator='', lang='ru'))
print(b.capitalize()+'.')
#


# !!! Date !!!
current_date = datetime.now().date()
print(current_date)

today = str(date.today().strftime('%d.%m.%Y'))
print(today)
#

