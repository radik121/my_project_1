import win32com.client
import re
import datetime
import os


def mail_data():
    path = r"Z:/ЛОГИСТИКА/Динамика продаж/ЗАБИВКА/ПРАВ/"

    yesterday = (datetime.datetime.today() - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
    # Обращаемся к Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = outlook.Folders[2]  # Общая папка
    subfldr = folder.Folders[1]  # папка Входящие
    inbox = subfldr.Folders[3]  # подпапка во входящих salesdbf
    messages = inbox.Items
    # date_message = [date.CreationTime.date() for date in messages][-2::]   # даты последних 2-х сообщений
    # if str(date_message[0]) == str(date_message[1]) == yesterday:
    messages.Sort("[ReceivedTime]", True)
    yesterday_message = [i for i in messages if
                         str(i.CreationTime.date()) == yesterday]  # берем только вчерашние сообщения

    # Сохраняем вложения последних 2-х сообщений
    for message in yesterday_message:
        pattern = r'\d+/\d+/\d+.?,?\d+.?/?\d+\s?\w+'
        # pattern = r'\d+/\d+/\d+.?,?\d+/\d+/\d.?,?\d+'
        for attachment in message.Attachments:
            # print(bool(attachment.FileName))
            attachment.SaveAsFile(os.path.join(path, str(attachment.FileName)))
            print(f'Файл {attachment.FileName} загружен!')
        # Данные из тела письма
        if 'новосибирск' in str(message.Sender) and len(re.findall(pattern, message.body)) != 0:
            novomol = re.findall(pattern, message.body)[0]
        if 'Пламида' in str(message.Sender) and len(re.findall(pattern, message.body)) != 0:
            aviapark = re.findall(pattern, message.body)[0]
    # print(novomol, aviapark)

    res = []
    for i in [re.split('[ /]', novomol), re.split('[ /]', aviapark)]:
        if len(i) < 4:
            i.append('0')
        if int(i[0]) < int(i[1]):
            i[0], i[1] = i[1], i[0]
        res.append(i)

    return {'1323': res[0], '1324': res[1]}


if __name__ == '__main__':
    print(mail_data())

    # Файлы загружены!
    # {'1323': ['947', '486', '663600', '2716'], '1324': ['978', '439', '403855', '2612']}
