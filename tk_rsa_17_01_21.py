from tkinter import filedialog as fd
import colorama, requests, multiprocessing
from tkinter import *

import urllib3
from requests.packages.urllib3.exceptions import InsecureRequestWarning
from termcolor import colored
from default_functions import to_digit, print_col, eternity_cycle_deep
import datetime
from script_rsa_17_01_21 import starting


def start():
    global file_name
    date = date_p.get()
    if file_name == '':
        print_col('[Error] Вы не выбрали файл! Перезапустите программу!', 'red')
        eternity_cycle_deep()
    elif '.xls' not in file_name:
        print_col('[Error] Вы выбрали файл без .xls расширения! Перезапустите программу!', 'red')
        eternity_cycle_deep()
    else:
        print_col(f'[True] Выбранный файл: {file_name}', 'green')

    if date == '':
        date = str(datetime.datetime.now().date().today()).split('-')
        date = date[2] + '.' + date[1] + '.' + date[0]
        print_col(f'[Added] Дата поиска не была указана, поэтому выбрана текущая дата ({date}).', 'yellow')
    elif date[0].isdigit() and date[1].isdigit() and date[2] == '.' and date[3].isdigit() and date[4].isdigit() and \
            date[5] == '.' and date[6].isdigit() and date[7].isdigit() and date[8].isdigit() and date[9].isdigit():
        print_col(f'[True] Указанная дата: {date}', 'green')
    else:
        print_col(f'[Error] Некорректно указана дата: {date}', 'red')

    POS_VIN = to_digit(vins_col.get().upper())
    if str(POS_VIN).isdigit():
        print_col(f'[True] Введенный индекс столбца с данными - {POS_VIN}', 'green')
    else:
        print_col(f'[Error] Неверно введена буква столбца. Перезапустите программу', 'red')
        eternity_cycle_deep()

    antiCaptchaKey = key.get()
    if antiCaptchaKey == '':
        print_col('[True] Так как ключ антикапчи не был введен, проверка будет происходить при помощи CapMonster!', 'green')
    else:
        print_col('[Added] Так как ключ антикапчи был введен, проверка будет происходить при помощи сервисов RuCaptcha'
                  ' или AntiCaptcha!', 'yellow')

    pool_col = (pool_cool_p.get())
    if pool_col == '':
        print_col('[Added] Так как вы не ввели количество используемых Пулов, то выставлено значение по умолчанию'
                  ' (50 пулов)', 'yellow')
        pool_col = 50
    else:
        print_col(f'[True] Будет использовано {pool_col} пулов.', 'green')
        pool_col = int(pool_col)
    delta = (delta_p.get())
    if delta == '':
        delta = 300
        print_col('[Added] Так как вы не ввели через сколько проверок программе делать небольшой технический перерыв, '
                  'то выставлено значение по умолчанию (300 проверок).', 'yellow')
    else:
        delta = int(delta)
        print_col(f'[True] Выбрано деление по {delta} строк в одной проверке.', 'green')
    starting(file_name, POS_VIN, date, antiCaptchaKey, delta, pool_col)
    eternity_cycle_deep()


def insertText():
    global file_name
    file_name = fd.askopenfilename()
    print('[system] Выбранный файл: ' + file_name)


if __name__ == '__main__':
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    multiprocessing.freeze_support()
    requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
    print_col('Здравствуйте! Введите данные для начала работы.', 'green')
    file_name = ''
    window = Tk()
    window.title("РСА | КБ")
    window.geometry('300x300')
    window.resizable(False, False)
    date_p = StringVar()
    vins_col = StringVar()
    key = StringVar()
    delta_p = StringVar()
    pool_cool_p = StringVar()
    file_btn = Button(window, text="Выбор Exel файла", command=insertText, width=25, height=1).place(relx=0.19,
                                                                                                     rely=0.02)

    key_text = Label(window, text='Ключ Anticaptcha(если нужен):').place(relx=0.2, rely=0.12)
    key_entry = Entry(textvariable=key).place(relx=0.29, rely=0.2)

    vin_text = Label(window, text='Столбец с VIN:').place(relx=0.35, rely=0.3)
    vin_entry = Entry(textvariable=vins_col).place(relx=0.29, rely=0.38)

    minus_text = Label(window, text='Дата проверки: (dd.mm.yyyy)').place(relx=0.22, rely=0.48)
    minus_entry = Entry(textvariable=date_p).place(relx=0.29, rely=0.56)

    q_text = Label(window, text='Делить по:').place(relx=0.18, rely=0.66)
    q_entry = Entry(textvariable=delta_p).place(relx=0.07, rely=0.74)

    w_text = Label(window, text='Количество потоков:').place(relx=0.51, rely=0.66)
    w_entry = Entry(textvariable=pool_cool_p).place(relx=0.5, rely=0.74)

    start_btn = Button(window, text="Начать проверку", command=start, width=25, height=1).place(relx=0.19,
                                                                                                   rely=0.86)

    window.mainloop()
