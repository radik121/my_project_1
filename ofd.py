import datetime
import json
import requests
from package.tokens import aviapark, novomol


def ofd_data(token, date: str = None) -> list:  # date = "2022-05-03" !!!

    url = 'https://api.ofd-ya.ru/ofdapi/v1/documents'
    headers = {
        'Ofdapitoken': token,
        'Content-Type': 'application/json',
    }

    # берем вчерашнюю дату, если не задали в параметрах
    yesterday = (datetime.datetime.today() - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
    data = f'"date": "{yesterday}"'
    if date:
        data = f'"date": "{date}"'
    # neg = 1324
    if token == aviapark:
        fiscalDrives = ["9287440301007533", "9287440301007534", "9960440302691467", "9960440302692195"]
    # neg = 1323
    else:
        fiscalDrives = ["9960440300300279", "9960440300301996",
                        "9960440302636573", "9960440302637693",
                        "9960440302640097"]

    total_sum = 0
    total_q = 0
    total_docs = 0

    for driver in fiscalDrives:
        # Делаем запрос в ОФД-Я для выгрузки данных за дату, определенную выше
        res = requests.post(url,
                            f'{{"fiscalDriveNumber":"{driver}", {data}}}',
                            headers=headers)
        docs = json.loads(res.text)

        if docs['count'] > 0:
            # Собираем все цены в список и суммируем
            summa = sum(
                [float(int(i['totalSum']) / 100)
                 if i['operationType'] == int(1)
                 else float(int(i['totalSum']) / -100)
                 for i in docs['items']]
            )

            # Собираем кол-во ед. в каждом чеке в список и суммируем
            q = sum(
                [int(j['quantity'])
                 if i['operationType'] == int(1)
                 else int(j['quantity']) / -1
                 for i in docs['items']
                 for j in i['items']]
            )
            total_sum += summa
            total_q += q
            total_docs += docs['count']

    return [str(int(total_q)), str(total_docs), str(round(total_sum, 2))]


if __name__ == '__main__':
    print(ofd_data(novomol))
    print(ofd_data(aviapark))

    # ['947', '486', '663600.0']
    # ['971', '438', '400162.0']
