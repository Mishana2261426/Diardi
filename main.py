# -*- encoding: utf-8 -*-
import threading
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr

from tkinter import *
from tkinter import filedialog, BOTH, END, HORIZONTAL, Tk, scrolledtext, ttk

import smtplib
import time
import openpyxl
import re

mark = 0

def print_email(fr, to, l):
    print(str(l) + ' Отправка: ' + fr + ' к ' + to + ' asd')
    console.configure(state='normal')  # enable insert
    console.insert(END, str(l) + ' Отправка: ' + fr + ' --> ' + to + '\n')
    console.yview(END)  # autoscroll
    console.configure(state='disabled')  # disable editing
    console.pack()


def thread(func):
    def wrapper(*args, **kwargs):
                    current_thread = threading.Thread(target=func, args=args, kwargs=kwargs)
                    current_thread.start()
    return wrapper

def browseFiles():
    global filename

    filename = filedialog.askopenfilename(initialdir="/", title="Select a File", filetypes=(("Excel", "*.xlsx"), ("all files", "*.*")))

    label_file_explorer.configure(text=filename)
    if filename != '':
        button_start['state'] = 'normal'
        console.configure(state='normal')  # enable insert
        console.delete('1.0', END)
        console.pack()
        console.configure(state='disable')

@thread
def mail_start():
    email_text = ''

    button_explore['state'] = 'disable'
    button_start['state'] = 'disable'

    wb = openpyxl.load_workbook(filename)
    to_list = wb['Адреса отправки']
    from_list = wb['Адреса с которых отправить']
    letter_list = wb['Письма']
    settings = wb['Тех. данные']
    j = 1
    k = 1
    l = 0
    m = len(to_list['A'])
    jj = len(from_list['A'])
    kk = len(letter_list['A'])

    server = smtplib.SMTP(settings['B2'].value)
    msg_type = settings['B3'].value

    settings_delay.configure(text='Задержка ' + str(settings['B1'].value) + ' секунд(ы)')
    print('Задержка ' + str(settings['B1'].value) + ' секунд(ы)')
    print(m, jj, kk)
    print(msg_type)



    for i in range(1, m + 1):
        time.sleep(settings['B1'].value)
        msg = MIMEMultipart("alternative")
        print(to_list['A' + str(i)].value)
        msg['To'] = to_list['A' + str(i)].value

        print(from_list['A' + str(j)].value)
        adress = from_list['A' + str(j)].value

        msg['From'] = formataddr((str(Header(letter_list['A' + str(k)].value, 'utf-8')), adress))

        password = from_list['B' + str(j)].value

        print(letter_list['A' + str(k)].value)
        msg['Subject'] = letter_list['B' + str(k)].value


        try:
            # letter_text = MIMEText(letter_list['C' + str(k)].value, "plain", "utf-8")
            # letter_text = letter_list['C' + str(k)].value
            # letter_text = MIMEText(letter_list['C' + str(k)].value, "html")
            letter_text = letter_list['C' + str(k)].value
            if re.search("&email", str(letter_text)):
                letter_text = MIMEText(str(letter_text).replace("&email", msg['To']), msg_type)
            else:
                letter_text = MIMEText(letter_text, msg_type)
            # print(MIMEText(str(letter_text)))
            print(i)
            l = l + 1
            print_email(from_list['A' + str(j)].value, to_list['A' + str(i)].value, l)
            msg.attach(letter_text)

            s = smtplib.SMTP('smtp.gmail.com: 587')
            s.starttls()
            s.login(adress, password)
            s.sendmail(msg['From'], [msg['To']], msg.as_string())

            print('\n')
            if j == jj:
                j = 1
            else:
                j += 1
            if k == kk:
                k = 1
            else:
                k += 1
        except BaseException:
            pass

    button_start['state'] = 'disable'
    button_explore['state'] = 'normal'
    settings_delay.configure(text='')


window = Tk()


# Set window size
window.geometry("700x700")

# Set window background color
window.config(background="white")

# Create a File Explorer label
label_file_explorer = Label(window,
                            text="",
                            fg="blue")



button_explore = Button(window,
                        text="Выберите файл",
                        command=browseFiles)

button_start = Button(window,
                      text='Запуск',
                      command=mail_start)

settings_delay = Label(window,
                           text='',
                           fg="green")


button_start['state'] = 'disable'

# if mark == 1:
#     button_start['state'] = 'normal'

ttk.Separator(window, orient=HORIZONTAL).pack(fill=BOTH)  # line in-between
console = scrolledtext.ScrolledText(window, state='disable')

label_file_explorer.pack()

button_explore.pack()
button_start.pack()
settings_delay.pack()
window.mainloop()

