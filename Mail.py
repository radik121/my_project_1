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

    # Сохраняем вложения последних 2-х сообщений
    for message in messages:
        for attachment in message.Attachments:
            if str(message.ReceivedTime.date()) == yesterday:
                # print(attachment.FileName)
                attachment.SaveAsFile(os.path.join(path, str(attachment.FileName)))
    print('Файлы загружены!')

    # Данные из тела письма
    pattern = r'\d+/\d+/\d+.?,?\d+.?/?\d+\s?\w+'
    # pattern = r'\d+/\d+/\d+.?,?\d+/\d+/\d.?,?\d+'
    novomol = [re.findall(pattern, i.body) for i in messages if 'Якшин' in i.To][0][0]
    aviapark = [re.findall(pattern, i.body) for i in messages if 'Sergey Yakshin' in i.To][0][0]

    res = []
    for i in [re.split('[ /]', novomol), re.split('[ /]', aviapark)]:
        if len(i) < 4:
            i.append('0')
        if int(i[0]) < int(i[1]):
            i[0], i[1] = i[1], i[0]
        res.append(i)

    return {'1323': res[0], '1324': res[1]}
    # else:
    #     return "Дата отчетов не совпадает!"


if __name__ == '__main__':
    print(mail_data())

    # Файлы загружены!
    # {'1323': ['947', '486', '663600', '2716'], '1324': ['978', '439', '403855', '2612']}


