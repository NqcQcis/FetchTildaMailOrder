#!/usr/bin/env python
# -*- coding: utf-8 -*-

# Блок для подключения библиотек
import email
import imaplib
# Добавляем необходимые подклассы - MIME-типы
import mimetypes  # Импорт класса для обработки неизвестных MIME-типов, базирующихся на расширении файла
import os
import re
import smtplib
from datetime import datetime
from email import encoders  # Импортируем энкодер
from email.header import Header
from email.mime.audio import MIMEAudio  # Аудио
from email.mime.base import MIMEBase  # Общий тип
from email.mime.image import MIMEImage  # Изображения
from email.mime.multipart import MIMEMultipart  # Многокомпонентный объект
from email.mime.text import MIMEText  # Текст/HTML
from email.utils import formataddr

import gspread

# Блок для подключения переменных с файлами (сертификат, номера заказов, файлы для отправки

# Текущий каталог
scriptCurrentFolder = os.getcwd()

# Каталог с файлами
scriptFilesFolder = os.path.join(scriptCurrentFolder, "Files")
scriptLibsFolder = os.path.join(scriptCurrentFolder, "Libs")

# Параметры логина для просмотра ящика
inboxMailServer = {imap server}
inboxMailServerLogin = {Mail login}
inboxMailServerPassword = {password}

# Параметры для отправки писем
outboxMailServer = {smtp server}
outboxMailServerPort = "465"
outboxMailServerSSLEnabled = True
outboxMailServerLogin = {smtp login}
outboxMailServerPassword = {smtp password}
outboxMailServerName = {user name}

# Папка, в которой буду храниться новые письма о заказах
# ordersFolder = "Orders"
ordersFolder = "Orders"
appliedFolder = "Applied"

# Пути к важным для работы файлам
P12KeyPath = str(os.path.join(scriptLibsFolder, "rocketchat.json"))
numOrderPath = os.path.join(scriptLibsFolder, "numOrder.txt")
textMessagePath = os.path.join(scriptLibsFolder, "textMessage.html")

# Пути до файлов с пособиями
PosSip = os.path.join(scriptFilesFolder, "Пособие по сыпи.pdf")
PosZap = os.path.join(scriptFilesFolder, "Пособие по проблемам со стулом.pdf")
PosSrig = os.path.join(scriptFilesFolder, "Пособие по срыгиваниям.pdf")

# Получатели в скрытой копии письма для отслеживания заказов
# receivers = ["semen@yaki-mov.ru", "alexandra@yaki-mova.ru"]
receivers = ["semen@yaki-mov.ru"]


##############################################################
##############################################################
##############################################################

def send_email(addr_to, bcc_to, msg_subj, msg_text, files):
    addr_from = outboxMailServerLogin  # Отправитель
    password = outboxMailServerPassword  # Пароль

    msg = MIMEMultipart()  # Создаем сообщение
    msg['From'] = formataddr((str(Header(outboxMailServerName, 'utf-8')), outboxMailServerLogin))  # Адресат
    msg['To'] = addr_to  # Получатель
    msg['Subject'] = msg_subj  # Тема сообщения
    bcc_to2 = str()
    for r in bcc_to:
        bcc_to2 = bcc_to2 + '"' + r + '",'
    msg['Bcc'] = bcc_to2[0:-1]

    bodyMessage = msg_text  # Текст сообщения
    msg.attach(MIMEText(bodyMessage, 'html', 'utf-8'))  # Добавляем в сообщение текст

    process_attachement(msg, files)

    # ======== Этот блок настраивается для каждого почтового провайдера отдельно ===============================================
    server = smtplib.SMTP_SSL(outboxMailServer, outboxMailServerPort)  # Создаем объект SMTP
    # if outboxMailServerSSLEnabled:
    #    server.starttls()                                      # Начинаем шифрованный обмен по TLS
    # server.set_debuglevel(True)                            # Включаем режим отладки, если не нужен - можно закомментировать
    server.login(addr_from, password)  # Получаем доступ
    server.send_message(msg)  # Отправляем сообщение
    server.quit()  # Выходим
    # ==========================================================================================================================


def process_attachement(msg, files):  # Функция по обработке списка, добавляемых к сообщению файлов
    for f in files:
        if os.path.isfile(f):  # Если файл существует
            attach_file(msg, f)  # Добавляем файл к сообщению
        elif os.path.exists(f):  # Если путь не файл и существует, значит - папка
            dir = os.listdir(f)  # Получаем список файлов в папке
            for file in dir:  # Перебираем все файлы и...
                attach_file(msg, f + "/" + file)  # ...добавляем каждый файл к сообщению


def attach_file(msg, filepath):  # Функция по добавлению конкретного файла к сообщению
    filename = os.path.basename(filepath)  # Получаем только имя файла
    ctype, encoding = mimetypes.guess_type(filepath)  # Определяем тип файла на основе его расширения
    if ctype is None or encoding is not None:  # Если тип файла не определяется
        ctype = 'application/octet-stream'  # Будем использовать общий тип
    maintype, subtype = ctype.split('/', 1)  # Получаем тип и подтип
    if maintype == 'text':  # Если текстовый файл
        with open(filepath) as fp:  # Открываем файл для чтения
            file = MIMEText(fp.read(), _subtype=subtype)  # Используем тип MIMEText
            fp.close()  # После использования файл обязательно нужно закрыть
    elif maintype == 'image':  # Если изображение
        with open(filepath, 'rb') as fp:
            file = MIMEImage(fp.read(), _subtype=subtype)
            fp.close()
    elif maintype == 'audio':  # Если аудио
        with open(filepath, 'rb') as fp:
            file = MIMEAudio(fp.read(), _subtype=subtype)
            fp.close()
    else:  # Неизвестный тип файла
        with open(filepath, 'rb') as fp:
            file = MIMEBase(maintype, subtype)  # Используем общий MIME-тип
            file.set_payload(fp.read())  # Добавляем содержимое общего типа (полезную нагрузку)
            fp.close()
            encoders.encode_base64(file)  # Содержимое должно кодироваться как Base64
    file.add_header('Content-Disposition', 'attachment', filename=filename)  # Добавляем заголовки
    msg.attach(file)  # Присоединяем файл к сообщению


# Читаем файл с текстом письма в html и заносим в переменую для дальнейшей отправки
with open(textMessagePath, "r", encoding='utf-8') as textMessage:
    textMessage = textMessage.read()

# Подключаемся к почтовому серверу
mail = imaplib.IMAP4_SSL(inboxMailServer)
mail.login(inboxMailServerLogin, inboxMailServerPassword)

# Выгружаем список писем
mail.list()
mail.select(ordersFolder)
result, data = mail.search(None, "ALL")
ids = data[0]
id_list = ids.split()

if len(id_list) > 0:
    gc = gspread.service_account(filename=P12KeyPath)
    sh = gc.open("Заказы детская смесь")

# Обрабатываем каждое письмо и заносим в таблицу, оповещаем по электронной почте
for id_message in id_list:
    result, data = mail.fetch(id_message, "(RFC822)")
    raw_email = data[0][1]
    raw_email_string = raw_email.decode('utf-8')

    email_message = email.message_from_string(raw_email_string)

    body = email_message.get_payload(decode=True).decode('utf-8')

    #    print(body)
    # print(id_message)
    # print(sh.sheet1.get('A1'))
    # sh.sheet1.append_row(["adwdawd"])

    # Считываем номер заказа из файла
    numOrder = open(numOrderPath, "r")
    numOrderValue = numOrder.read()
    numOrder.close()

    tovarsArray = []
    body = body.split('\n')
    for stroka in body:
        if "@" in stroka:
            mailCustomer = (stroka.replace("Email: ", "")).replace("<br>", "")
        if "Name: " in stroka:
            nameCustomer = (stroka.replace("Name: ", "")).replace("<br>", "")
        if "Phone: " in stroka:
            phoneCustomer = (
                ((((stroka.replace("Phone: ", "")).replace("<br>", "")).replace("+", "")).replace(" (", "")).replace(
                    ") ", "")).replace("-", "")
        if "Payment time: " in stroka:
            paymentTimeCustomer = '{dt:%d}.{dt:%m}.{dt.year}'.format(
                dt=datetime.strptime(str(stroka.replace("Payment time: ", "")).replace("<br>", ""), "%d %b %Y %H:%M"))
        if "td style" in stroka:
            # print(stroka)
            tovarsArray.append(re.sub(r'\<[^>]*\>', '', stroka))

    # Формируем список файлов для передачи конкретному покупателю
    i = 1
    while i < len(tovarsArray) - 1:
        if "Пособие по сыпи" in tovarsArray[i]:
            tovars = [PosSip]
        if "Пособие проблемы со стулом" in tovarsArray[i]:
            tovars = [PosZap]
        if "Пособие по срыгиваниям" in tovarsArray[i]:
            tovars = [PosSrig]
        if "Пакет Сыпь+проблемы со стулом" in tovarsArray[i]:
            tovars = [PosSip, PosZap]
        if "Пакет Сыпь+срыгивания" in tovarsArray[i]:
            tovars = [PosSip, PosSrig]
        if "Пакет Срыгивания+проблемы со стулом" in tovarsArray[i]:
            tovars = [PosSrig, PosZap]
        if "Пакет Сыпь+проблемы со стулом+срыгивания" in tovarsArray[i]:
            tovars = [PosSip, PosZap, PosSrig]
        if "Товар 1" in tovarsArray[i]:
            tovars = [PosSip]
        if "Товар 2" in tovarsArray[i]:
            tovars = [PosSip, PosZap]
        if "Товар 3" in tovarsArray[i]:
            tovars = [PosSip, PosZap, PosSrig]
        i = i + 5

    # Вычисляем общую сумму заказа
    sumOrder = int()
    i = 4
    while i < len(tovarsArray) - 1:
        sumOrder = sumOrder + int(tovarsArray[i].replace(" RUB", ""))
        i = i + 5

    # Использование функции send_email()
    receiverMessage = mailCustomer
    bccMessage = receivers
    files = tovars
    subjectMessage = "Ваш заказ №" + numOrderValue

    send_email(receiverMessage, bccMessage, subjectMessage, textMessage, files)

    copy_res = mail.copy(id_message, appliedFolder)
    if copy_res[0] == 'OK':
        mail.store(id_message, '+FLAGS', '\\Deleted')
        mail.expunge()

    try:
        paymentTimeCustomer
    except NameError:
        paymentTimeCustomer = '{dt.day}.{dt.month}.{dt.year}'.format(dt=datetime.now())
    sh.sheet1.append_row([numOrderValue, paymentTimeCustomer, nameCustomer, mailCustomer, int(phoneCustomer), sumOrder],
                         "USER_ENTERED")
    numOrderValue = int(numOrderValue) + 1
    numOrder = open(numOrderPath, "w")
    numOrder.write(str(numOrderValue))
    numOrder.close()

mail.logout()
