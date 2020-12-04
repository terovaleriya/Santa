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
    def __init__(self, name: str, surname: str, social: str, wish: str, address: str, year: str, index: str):
        self.name = name
        self.surname = surname
        self.social = social
        self.wish = wish
        self.address = address
        self.index = index
        self.year = year


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
    datafrom = data[data['Фамилия'] == froms[i]]
    datato = data[data['Фамилия'] == tos[i]]

    From.surname = datafrom['Фамилия'].to_string(index=False).lstrip(' ')
    From.name = datafrom['Имя'].to_string(index=False).lstrip(' ')
    From.email_address = datafrom['Адрес электронной почты'].to_string(index=False).lstrip(' ')

    To.surname = datato['Фамилия'].to_string(index=False).lstrip(' ')
    To.name = datato['Имя'].to_string(index=False).lstrip(' ')
    To.year = datato['Курс'].to_string(index=False).lstrip(' ')
    To.social = datato['Аккаунт в социальной сети'].to_string(index=False).lstrip(' ')
    To.address = datato['Адрес проживания'].to_string(index=False).lstrip(' ')
    To.wish = datato['Пожелания'].to_string(index=False).lstrip(' ')
    To.index = datato['Индекс'].to_string(index=False).lstrip(' ')

    # the text will be emailed
    text = "Привет, " + From.name + "!\n\n"
    text += "Твоего подопечного зовут " + To.name + " " + To.surname + ". "
    text += "Он учится на " + To.year[0] + " курсе"
    if len(To.year.split()) == 2:
        text += " магистратуры"
    text += ".\n"
    text += "Вот, что он оставил в качестве пожелания: " + "\"" + To.wish + "\"." + "\n" + "Его адрес: " + To.address + ", " + To.index + ".\n"
    if not (isNaN(To.social)):
        text += "Ссылка на социальную сеть: " + To.social + ".\n"
    else:
        text += "К сожалению, твой подопечный не оставил ссылку на социальную сеть :(\n"

    message = MIMEText(text, 'plain', 'utf-8')
    message['Subject'] = Header('Твой подопечный!', 'utf-8')
    message['From'] = your_email

    # print(text)

    # # UNCOMMENT TO SEND
    # # sending the email
    # server.sendmail(your_email, [From.email_address], message.as_string())

# close the smtp server
server.close()
