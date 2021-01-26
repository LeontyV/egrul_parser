import json
import os
import re
from time import time
from time import sleep
from io import StringIO
from datetime import datetime as dt
import tkinter as tk
from tkinter import scrolledtext
from tkinter import messagebox as mb
from tkinter import filedialog as fd
from tkinter import Menu
from tkinter import Checkbutton
import base64

from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser
from requests.auth import HTTPProxyAuth
import requests

from requests_wrapper import RequestsWrapper
from logger import logger
from check_cert import get_cert_attr
from info_to_excel import to_file
from proxy import get_random_ua


def egrul_parser(ogrn):
    headers = get_random_ua()
    proxies = None
    check_proxy = False
    data = {'query': ogrn}
    egrul_parser_path = create_folders()
    proxy_settings_path = egrul_parser_path + '/proxy_settings'
    if os.path.exists(proxy_settings_path):
        file = open(proxy_settings_path, 'rb')
        proxies_enc = file.read()
        proxies_dec = base64.b64decode(proxies_enc)
        proxies_dec = proxies_dec.decode('utf-8')
        proxies_dec = proxies_dec.replace('"', '\\"')
        proxies_dec = proxies_dec.replace("'", '"')
        proxies = json.loads(proxies_dec)
        check_proxy = getboolean(proxies['proxy'][1:-1])
        file.close()

        if proxies['login'] and proxies['pass']:
            auth = HTTPProxyAuth(proxies['login'], proxies['pass'])
        else:
            auth = False

        _proxies = f"{proxies['http']}, {proxies['https']}"

        proxies = json.loads('{' + _proxies + '}')

    r = requests.session()
    url = "https://egrul.nalog.ru"

    print(check_proxy)
    if check_proxy:
        print('качаем выписку с прокси')
        resp = r.get(url, proxies=proxies, auth=auth, headers=headers)
        content = r.post(url, data=data, proxies=proxies, auth=auth, headers=headers)
    else:
        print('качаем выписку без прокси')
        resp = r.get(url, headers=headers)
        content = r.post(url, data=data, headers=headers)

    tmp = str(content.content.decode('utf-8'))

    res_dict = json.loads(tmp)

    curr_time = str(int(time()))
    url2 = f'https://egrul.nalog.ru/search-result/{res_dict["t"]}?r={curr_time}&_={curr_time}'
    if check_proxy:
        org_info = r.get(url2, proxies=proxies, auth=auth, headers=headers)
    else:
        org_info = r.get(url2)

    org_info = org_info.content.decode('utf-8')
    logger(org_info)

    try:
        id = re.findall('"t":"(.+)","p', org_info)[0]
    except IndexError:
        err_message()
        return None, org_info

    dwnld_id = id

    curr_time = str(int(time()))
    url3 = f'https://egrul.nalog.ru/vyp-request/{dwnld_id}?r=&_={curr_time}'
    if check_proxy:
        content = r.get(url3, proxies=proxies, auth=auth, headers=headers)
    else:
        content = r.get(url3)

    if content.status_code == 200:
        status = False
        while not status:
            curr_time = str(int(time()))
            url4 = f'https://egrul.nalog.ru/vyp-status/{dwnld_id}?r={curr_time}&_={curr_time}'
            if check_proxy:
                content = r.get(url4, proxies=proxies, auth=auth, headers=headers)
            else:
                content = r.get(url4)

            if content.status_code == 200:
                if content.content.decode('utf-8') == '{"status":"ready"}':
                    status = True
            sleep(1)

        dwnld_url = f'https://egrul.nalog.ru/vyp-download/{dwnld_id}'
        if check_proxy:
            final_content = r.get(dwnld_url, proxies=proxies, auth=auth, headers=headers)
        else:
            final_content = r.get(dwnld_url)

        now_date = dt.now().strftime('%d-%m-%Y')

        path = create_folders()
        if folder_lbl['text'] != 'Укажите папку!' and folder_lbl['text']:
            path = folder_lbl['text']
        path = path + '\\' + now_date
        if not os.path.exists(path):
            os.mkdir(path)

        file_name = f'{path}\{str(ogrn)}.pdf'
        open(file_name, 'wb').write(final_content.content)

        logger(f'Файл "{str(ogrn)}.pdf" сохранен в папке {path}')
        info_lbl['text'] = f'Файл "{str(ogrn)}.pdf" сохранен в папке {path}'

        return file_name, org_info


def pdf_read(path):
    output_string = StringIO()
    with open(path, 'rb') as in_file:
        parser = PDFParser(in_file)
        doc = PDFDocument(parser)
        rsrcmgr = PDFResourceManager()
        device = TextConverter(rsrcmgr, output_string, laparams=LAParams())
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        for page in PDFPage.create_pages(doc):
            interpreter.process_page(page)

        return output_string.getvalue()


def get_values(vyp, org_info):
    values = {}
    vyp_new = []
    vyp = vyp.replace('\x0c', ' ')
    for elem in vyp.split('\n'):
        if elem != '':
            vyp_new.append(elem)

    # ищем дату и номер выписки
    vyp_date_index = vyp_new.index('ВЫПИСКА')
    while True:
        try:
            vyp_date = re.findall('(\d{2}.\d{2}.\d{4})', vyp_new[vyp_date_index])[0]
            values['дата'] = vyp_new[vyp_date_index]
            break
        except IndexError:
            pass

        vyp_date_index += 1

    vyp_num_index = vyp_new.index('дата формирования выписки')
    number_elems = [elem for elem in vyp_new[vyp_num_index:vyp_num_index + 4] if re.findall('([0-9])', elem)]
    values['номер'] = ''.join(number_elems)

    # ищем свидетельство
    try:
        certificate = [elem for elem in vyp_new if 'Серия, номер и дача выдачи свидетельства' in elem][0]
        values['свидетельство'] = vyp_new[vyp_new.index(certificate)][-12:].strip()
        cert_date_index = vyp_new.index(certificate) + 1
        while True:
            try:
                vyp_date = re.findall('(\d{2}.\d{2}.\d{4})', vyp_new[cert_date_index])[0]
                values['свидетельство'] += ' от ' + vyp_date
                break
            except IndexError:
                pass
            cert_date_index += 1

    except IndexError:
        values['свидетельство'] = 'нет'

    res = json.loads(org_info)

    # вытаскиваем дату регистрации
    date_reg = res['rows'][0]['r']

    if values['свидетельство'] == 'нет':
        values['свидетельство'] = vyp_new[vyp_new.index(date_reg) - 1] + ' от ' + date_reg

    # вытаскиваем адрес
    try:
        adress = res['rows'][0]['a']
    except IndexError:
        adress = 'Не нашел'
    except KeyError:
        adress = 'Не нашел'

    # вытаскиваем название
    try:
        name_long = res['rows'][0]['n']
    except IndexError:
        name_long = 'Не нашел'
    except KeyError:
        name_long = 'Не нашел'

    try:
        name_short = res['rows'][0]['c']
    except IndexError:
        name_short = name_long
    except KeyError:
        name_short = 'Не нашел'

    # вытаскиваем реквизиты
    try:
        ogrn = res['rows'][0]['o']
    except IndexError:
        ogrn = 'не нашел'
    except KeyError:
        ogrn = 'Не нашел'

    try:
        inn = res['rows'][0]['i']
    except IndexError:
        inn = 'не нашел'
    except KeyError:
        inn = 'Не нашел'

    # вытаскиваем должность и ФИО
    if res['rows'][0]['k'] == 'ul':
        try:
            fio = res['rows'][0]['g']
            fio = fio[fio.find(':')+1:]
            fio = fio.lower().title().strip()
        except IndexError:
            fio = 'не нашел'
        except KeyError:
            fio = 'Не нашел'

    if res['rows'][0]['k'] == 'fl':
        fio = name_long.lower().title()

    try:
        position = res['rows'][0]['g']
        position = position[:position.find(':')]
        position = position[0].upper() + position[1:].lower().strip()
    except IndexError:
        position = 'не нашел'
    except KeyError:
        position = 'Не нашел'

    if '"' in position:
        index = position.find('"')
        position = position[:index+1] + position[index+1].upper() + position[index+2:]

    values['адрес'] = adress.replace('\\', '')
    values['краткое наименование'] = name_short.replace('\\', '')
    values['полное наименование'] = name_long.replace('\\', '')
    values['огрн'] = ogrn.strip()
    values['инн'] = inn.strip()
    values['фио'] = fio.strip()
    values['должность'] = position.replace('\\', '')

    logger(f'Выписка {ogrn} просмотрена')
    return values


def main(ogrn):
    logger(f'Приступаем к загрузке выписки {ogrn}')
    file_name, org_info = egrul_parser(ogrn)
    if file_name is not None:
        logger(f'Выписка {ogrn} загружена')
        vyp = pdf_read(file_name)
        values = get_values(vyp, org_info)

        return values
    else:
        logger(f'Выписка {ogrn} не загружена')

    return


def clicked(event=None):
    try:
        re.findall('^(\d{13,15})$', str(ogrn_input.get()))[0]
    except IndexError:
        mb.showerror('Ошибка', 'Поле ОГРН должно быть заполнено!')
        return

    result = main(int(ogrn_input.get()))
    if result:
        # очищаем
        props_output.delete(0, 'end')
        cert_output.delete(0, 'end')
        state_output.delete(0, 'end')
        fio_output.delete(0, 'end')
        inn_output.delete(0, 'end')
        ogrn_output.delete(0, 'end')
        name_output.delete(0, 'end')
        short_name_output.delete(0, 'end')
        addres_output.delete(0, 'end')

        # вставляем новые данные
        props_output.insert(0, f"Выписка {result['номер']} от {result['дата']}")
        cert_output.insert(0, f"{result['свидетельство']}")
        state_output.insert(0, f"{result['должность']}")
        fio_output.insert(0, f"{result['фио']}")
        inn_output.insert(0, f"{result['инн']}")
        ogrn_output.insert(0, f"{result['огрн']}")
        name_output.insert(0, f"{result['полное наименование']}")
        short_name_output.insert(0, f"{result['краткое наименование']}")
        addres_output.insert(0, f"{result['адрес']}")


def choice_proxy():
    global prx

    def proxy_save():
        set_dict = {}
        if login.get() == '' or password.get() == '':
            set_dict['http'] = f'"http": "http://{proxy_address.get()}:{port.get()}"'
            set_dict['https'] = f'"https": "https://{proxy_address.get()}:{port.get()}"'
            set_dict['login'] = ''
            set_dict['pass'] = ''
            set_dict['proxy'] = f'"{prx.get()}"'
            enc_str = str(set_dict).encode('utf-8')
            enc_str_b64 = base64.b64encode(enc_str)
        else:
            set_dict['http'] = f'"http": "http://{proxy_address.get()}:{port.get()}"'
            set_dict['https'] = f'"https": "https://{proxy_address.get()}:{port.get()}"'
            set_dict['login'] = login.get()
            set_dict['pass'] = password.get()
            set_dict['proxy'] = f'"{prx.get()}"'
            enc_str = str(set_dict).encode('utf-8')
            enc_str_b64 = base64.b64encode(enc_str)

        egrul_parser_path = create_folders()
        proxy_settings_path = egrul_parser_path + '/proxy_settings'
        proxy_settings = open(proxy_settings_path, 'wb')
        proxy_settings.write(enc_str_b64)

        proxy_settings.close()

        pw.destroy()
        window.deiconify()

    def on_close_proxy():
        pw.destroy()
        window.deiconify()

    def change_check():
        if prx.get():
            prx.set(0)
        else:
            prx.set(1)

    window.withdraw()

    pw = tk.Tk()
    width = 300
    height = 150
    pw.title("Настройки прокси")
    pw.geometry(f'{str(width)}x{str(height)}')
    pw.resizable(0, 0)

    proxy_lbl = Label(pw, text="адрес: ")
    proxy_lbl.grid(column=0, row=0)

    proxy_address = Entry(pw, width=30)
    proxy_address.grid(column=1, row=0, sticky='w')

    prx = BooleanVar()
    prx.set(0)
    proxy_check = Checkbutton(pw, text='Прокси:', variable=prx, onvalue=1, offvalue=0, command=change_check)
    proxy_check.grid(column=3, row=0, sticky='e')

    port_lbl = Label(pw, text="порт: ")
    port_lbl.grid(column=0, row=1)

    port = Entry(pw, width=5)
    port.grid(column=1, row=1, sticky='w')

    login_lbl = Label(pw, text="логин: ")
    login_lbl.grid(column=0, row=2)

    login = Entry(pw, width=20)
    login.grid(column=1, row=2, sticky='w')

    passwrd_lbl = Label(pw, text="пароль: ")
    passwrd_lbl.grid(column=0, row=3)

    password = Entry(pw, show="*", width=20)
    password.grid(column=1, row=3, sticky='w')

    btn = Button(pw, text="Сохранить", command=proxy_save)
    btn.grid(column=1, row=5)

    btn = Button(pw, text="Закрыть", command=on_close_proxy)
    btn.grid(column=0, row=5)

    egrul_parser_path = create_folders()
    proxy_settings_path = egrul_parser_path + '/proxy_settings'
    if os.path.exists(proxy_settings_path):
        file = open(proxy_settings_path, 'rb')
        proxies_enc = file.read()
        proxies_dec = base64.b64decode(proxies_enc)
        proxies_dec = proxies_dec.decode('utf-8')
        proxies_dec = proxies_dec.replace('"', '\\"')
        proxies_dec = proxies_dec.replace("'", '"')
        proxies_dec = json.loads(proxies_dec)
        file.close()
        _login = proxies_dec.get('login')
        _pass = proxies_dec.get('pass')
        if _login is None or _pass is None:
            logger('Ошибка с парсингом параметров логина и пароля прокси')
            _login = ''
            _pass = ''

        _proxy = proxies_dec.get('http')
        try:
            _address = re.findall('(http:\W\W|@|)([a-z0-9.-]+):\d+', _proxy)[0][1]
            _port = re.findall(':(\d+)', _proxy)[0]
        except IndexError:
            logger('Ошибка с парсингом параметров адреса и порта прокси')
            _address = ''
            _port = ''

        proxy_address.insert(0, _address)
        port.insert(0, _port)
        login.insert(0, _login)
        password.insert(0, _pass)
        if getboolean(proxies_dec.get('proxy')[1:-1]):
            proxy_check.select()
            prx.set(1)
        else:
            proxy_check.deselect()
            prx.set(0)

    pw.protocol("WM_DELETE_WINDOW", on_close_proxy)

    pw.mainloop()


def on_close():
    window.destroy()


def err_message():
    mb.showerror('Ошибка', 'Такой ОГРН не найден!')


def open_folder():
    egrul_parser_path = create_folders()
    folder_settings_path = egrul_parser_path + '\\folder_settings'
    set_dict = {}
    folder = fd.askdirectory(title='Выберите папку для сохранения файлов')
    folder_lbl['text'] = folder
    folder_lbl['fg'] = 'green'
    set_dict['folder'] = folder
    enc_str = str(set_dict).encode('utf-8')
    enc_str_b64 = base64.b64encode(enc_str)

    folder_settings = open(folder_settings_path, 'wb')
    folder_settings.write(enc_str_b64)

    folder_settings.close()


def open_cert():
    certificates = fd.askopenfiles(title='Выберите сертификаты для чтения')
    now_date = dt.now().strftime('%d-%m-%Y')
    if folder_lbl['text'] and folder_lbl['text'] != 'Укажите папку':
        path = folder_lbl['text']
        path = path + '/' + now_date
    else:
        path = now_date
        path = f'{os.getcwd()}\{path}'
    if not os.path.exists(path):
        os.mkdir(path)

    for certificate in certificates:
        cert_info = get_cert_attr(certificate.name)
        to_file(path, cert_info)

    certificate_lbl['fg'] = 'green'
    certificate_lbl['text'] = f'Сертификаты успешно распаршены - {len(certificates)} шт.'


def handle_click_ogrn(event):
    _buffer = str(window.clipboard_get()).strip()
    if len(_buffer) == 13 or len(_buffer) == 15:
        ogrn_input.delete(0, 'end')
        ogrn_input.insert(0, str(_buffer))


def handle_click_copy(event):
    window.clipboard_clear()
    window.clipboard_append(event.widget.get())
    window.update()


def create_folders():
    home_path = os.getenv('USERPROFILE')
    egrul_parser_path = home_path + '\egrul_parser'
    if not os.path.exists(egrul_parser_path):
        os.mkdir(egrul_parser_path)

    return egrul_parser_path


def fill_folder_path():
    egrul_parser_path = create_folders()
    folder_settings_path = egrul_parser_path + '\\folder_settings'
    if os.path.exists(folder_settings_path):
        file = open(folder_settings_path, 'rb')
        folder_enc = file.read()
        folder_dec = base64.b64decode(folder_enc)
        folder_dec = folder_dec.decode('utf-8')
        folder_dec = folder_dec.replace('"', '\\"')
        folder_dec = folder_dec.replace("'", '"')
        folder_dec = json.loads(folder_dec)
        file.close()
        folder_lbl['text'] = folder_dec.get('folder')
        folder_lbl['fg'] = 'green'


if __name__ == '__main__':
    from tkinter import *

    proxy = {}

    width = 810
    height = 320

    window = Tk()
    window.title("Запрос ЕГРЮЛ")
    window.geometry(f'{str(width)}x{str(height)}')
    window.resizable(0, 0)

    main_menu = Menu(window)
    window.config(menu=main_menu)
    file_menu = Menu(main_menu, tearoff=0)
    main_menu.add_command(label='Прокси', command=choice_proxy)

    lbl = Label(window, text="Введите ОГРН: ")
    lbl.grid(column=0, row=0)

    ogrn_input = Entry(window, width=25)
    ogrn_input.grid(column=1, row=0, sticky='w')
    ogrn_input.bind('<Button-3>', handle_click_ogrn)

    btn = Button(window, text="Запросить", command=clicked)
    btn.grid(column=2, row=0, sticky='e')

    # далее строки с информацией из выписки
    name_lbl = Label(window, text="Название: ")
    name_lbl.grid(column=0, row=1)
    name_output = Entry(window, width=100)
    name_output.grid(column=1, row=1)
    name_output.bind('<Button-3>', handle_click_copy)

    short_name_lbl = Label(window, text="Название краткое: ")
    short_name_lbl.grid(column=0, row=2)
    short_name_output = Entry(window, width=100)
    short_name_output.grid(column=1, row=2)
    short_name_output.bind('<Button-3>', handle_click_copy)

    addres_lbl = Label(window, text="Адрес: ")
    addres_lbl.grid(column=0, row=3)
    addres_output = Entry(window, width=100)
    addres_output.grid(column=1, row=3)
    addres_output.bind('<Button-3>', handle_click_copy)

    inn_lbl = Label(window, text="ИНН: ")
    inn_lbl.grid(column=0, row=4)
    inn_output = Entry(window, width=100)
    inn_output.grid(column=1, row=4)
    inn_output.bind('<Button-3>', handle_click_copy)

    ogrn_lbl = Label(window, text="ОГРН: ")
    ogrn_lbl.grid(column=0, row=5)
    ogrn_output = Entry(window, width=100)
    ogrn_output.grid(column=1, row=5)
    ogrn_output.bind('<Button-3>', handle_click_copy)

    fio_lbl = Label(window, text="ФИО: ")
    fio_lbl.grid(column=0, row=6)
    fio_output = Entry(window, width=100)
    fio_output.grid(column=1, row=6)
    fio_output.bind('<Button-3>', handle_click_copy)

    state_lbl = Label(window, text="Должность:")
    state_lbl.grid(column=0, row=7)
    state_output = Entry(window, width=100)
    state_output.grid(column=1, row=7)
    state_output.bind('<Button-3>', handle_click_copy)

    cert_lbl = Label(window, text="Свидетельство:")
    cert_lbl.grid(column=0, row=8)
    cert_output = Entry(window, width=100)
    cert_output.grid(column=1, row=8)
    cert_output.bind('<Button-3>', handle_click_copy)

    props_lbl = Label(window, text="Реквизиты:")
    props_lbl.grid(column=0, row=9)
    props_output = Entry(window, width=100)
    props_output.grid(column=1, row=9)
    props_output.bind('<Button-3>', handle_click_copy)

    # Указание папок и файлов
    folder_lbl = Label(window, text="Укажите папку")
    folder_lbl.grid(column=0, row=10, columnspan=2, sticky='w')
    folder_btn = Button(window, text="Выбрать папку", command=open_folder)
    folder_btn.grid(column=2, row=10)

    certificate_lbl = Label(window, text="Укажите сертификаты")
    certificate_lbl.grid(column=0, row=11, columnspan=2, sticky='w')
    certificate_btn = Button(window, text="Сертификаты", command=open_cert)
    certificate_btn.grid(column=2, row=11, sticky='e')

    info_lbl = Label(window, text="", justify=LEFT, fg='green')
    info_lbl.grid(column=0, row=12, columnspan=2, sticky='w')

    close_btn = Button(window, text="Выход", command=on_close)
    close_btn.grid(column=2, row=13, sticky='e')

    fill_folder_path()
    window.bind('<Return>', clicked)
    window.mainloop()
