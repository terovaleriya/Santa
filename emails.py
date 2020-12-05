import smtplib
from email.header import Header
from email.mime.text import MIMEText

import pandas as pd

from matching import matching_table, data_table


def isNaN(value):
    try:
        import math
        return math.isnan(float(value))
    except:
        return False


class From(object):
    def __init__(self, name: str, surname: str, email_address: str):
        self.name = name
        self.surname = surname
        self.email_address = email_address


class To(object):
    def __init__(self, name: str, surname: str, social: str, wish: str, address: str, year: str, index: str,
                 number: str):
        self.name = name
        self.surname = surname
        self.social = social
        self.wish = wish
        self.address = address
        self.index = index
        self.year = year
        self.number = number


# change these as per use
your_email = "fmknsanta@gmail.com"
your_password = "santasanta"

# establishing connection with gmail
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)

# reading the spreadsheet
email_list = pd.read_excel(matching_table)
data = pd.read_excel(data_table)

# getting the FromTo's
froms = email_list['From']
tos = email_list['To']

# iterate through the records
for i in range(len(froms)):
    datafrom = data[data['Фамилия (Surname)'] == froms[i]]
    datato = data[data['Фамилия (Surname)'] == tos[i]]

    From.surname = datafrom['Фамилия (Surname)'].to_string(index=False).lstrip(' ')
    From.name = datafrom['Имя (Name)'].to_string(index=False).lstrip(' ')
    From.email_address = datafrom['Адрес электронной почты (Email)'].to_string(index=False).lstrip(' ')

    To.surname = datato['Фамилия (Surname)'].to_string(index=False).lstrip(' ')
    To.name = datato['Имя (Name)'].to_string(index=False).lstrip(' ')
    To.year = datato['Курс (Year)'].to_string(index=False).lstrip(' ')
    To.social = datato['Аккаунт в социальной сети (vk ID)'].to_string(index=False).lstrip(' ')
    To.address = datato['Адрес проживания (Address)'].to_string(index=False).lstrip(' ')
    To.wish = datato['Пожелания (Wishes)'].to_string(index=False).lstrip(' ')
    To.index = datato['Индекс (Index)'].to_string(index=False).lstrip(' ')
    To.number = datato['Номер телефона (Phone number)'].to_string(index=False).lstrip(' ')

    # the text will be emailed
    text = "Привет, " + From.name + "!\n\n"
    text += "Твоего подопечного зовут " + To.name + " " + To.surname + ". "
    if not (isNaN(To.year)):
        text += "Он учится на " + To.year[0] + " курсе"
        if len(To.year.split()) == 2:
            text += " магистратуры"
        text += "."
    text += "\nВот, что он оставил в качестве пожелания: " + "\"" + To.wish + "\"." + "\n" + "Его адрес: " + To.address + ", " + To.index + ".\n"
    if not (isNaN(To.social)):
        text += "Ссылка на социальную сеть: " + To.social + ".\n"
    else:
        text += "К сожалению, твой подопечный не оставил ссылку на социальную сеть :(\n"
    if not (isNaN(To.number)):
        text += "У нас есть номер телефона твоего подопечного, оставь его транспортной компании, так ему будет проще отслеживать полылку: " + To.number + ".\n"
    else:
        text += "Мы не располагаем номером его телефона.\n"
    text += "\nМы желаем тебе удачи и поздравляем с наступающим Новым годом!\n"

    message = MIMEText(text, 'plain', 'utf-8')
    message['Subject'] = Header('Твой подопечный!', 'utf-8')
    message['From'] = your_email
    message['To'] = From.email_address

    # print(text)

    # # UNCOMMENT TO SEND
    # # sending the email
    # server.sendmail(your_email, [From.email_address], message.as_string())

# close the smtp server
server.close()
