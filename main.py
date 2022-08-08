import glob
import os
import pandas as pd
import xlwings as xw
from Mail import mail_data
from ofd import ofd_data
from package.tokens import aviapark, novomol
import datetime


def attach_mail_to_excel():
    yesterday = (datetime.datetime.today() - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
    directory = r"Z:/ЛОГИСТИКА/Динамика продаж/ЗАБИВКА/ПРАВ/"
    all_files = glob.glob(os.path.join(directory, "*.csv"))
    data = []

    for f in all_files:
        d = pd.read_csv(f, sep=';')
        d.NetCurrency = d.SoldPrice.str.replace(',', '.').astype(float) * d.Qty.astype(int)
        d['date'] = yesterday
        d['год'] = ["FA" if i == 1900
                    else ("FF" if i == 2016
                          else ("FG" if i == 2017
                                else ("FH" if i == 2018
                                      else ("FI" if i == 2019
                                            else ("FJ" if i == 2020
                                                  else ("FK" if i == 2021
                                                        else "FL")))))) for i in d.Year]
        d['сезон'] = ["" if i == "Постоянные" else ("1-2" if i == "Весна" or i == "Лето" else "3-4") for i in d.Season]
        d['колл'] = d['год'] + d['сезон']
        d['маг'] = f.split("\\")[1][0:4]
        data.append(d)
        os.remove(f)

    df = pd.concat(data)

    wb = xw.Book(r"Z:/ЛОГИСТИКА/Динамика продаж/ЗАБИВКА/продажи_2022.xlsb")
    # wb = xw.Book(r"test.xlsx")
    ws = wb.sheets['исходник']

    last_row = ws.range('A' + str(ws.cells.last_cell.row)).end('up').row
    ws.range(f'A{last_row + 1}').value = df.values

    wb.save()
    wb.close()
    print('Продажи добавлены!')

    # Делаем сводную таблицу для шаблона
    pivot = df.pivot_table(
        index=['маг', 'date', 'Gender', 'колл'],
        values=['Qty', 'NetCurrency'],
        aggfunc='sum').reset_index()
    pivot['продажи'] = 'продажи'
    pivot['месяц'] = f'=TEXT("{yesterday}","ММММ")'
    pivot['год'] = f'=YEAR("{yesterday}")'
    pivot['Дата2'] = f'=DAY("{yesterday}")'
    pivot['Неделя'] = f'=TRUNC(MOD("{yesterday}"+3-WEEKDAY("{yesterday}",2),365.25)/7+1)'
    pivot['День недели'] = f'=WEEKDAY("{yesterday}",2)'
    pivot['1'] = ["kids" if i == "Для девочки" or i == "Для мальчика" else i for i in pivot.Gender]
    # print(pivot)
    return pivot


def chek_data():
    yesterday = (datetime.datetime.today() - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
    mail = mail_data()  # -> type Dict
    print('Данные с Outlook получены.')
    ofd = {'1323': ofd_data(novomol), '1324': ofd_data(aviapark)}
    print('Данные с ОФД-Я получены.')

    # if type(mail) is str:  # Если даты в модуле mail_data не одиаковые, то выходит текстовое предупреждение
    #     return mail
    # else:
    if float(mail['1323'][2].replace(',', '.')) == float(ofd['1323'][2]) and \
            float(mail['1324'][2].replace(',', '.')) == float(ofd['1324'][2]):
        data_work = mail
        print('Суммы ОФД и TIE верные!')

    else:
        data_work = ofd
        data_work['1323'].append(mail['1323'][3])
        data_work['1324'].append(mail['1324'][3])
        sum_1324 = float(mail['1324'][2].replace(',', '.')) - float(ofd['1324'][2])
        sum_1323 = float(mail['1323'][2].replace(',', '.')) - float(ofd['1323'][2])
        print(f"Суммы ОФД и TIE не сходятся!\n1323: {sum_1323} руб\n1324: {sum_1324} руб")

    res_data = []
    for k, v in data_work.items():
        v[0] = int(v[0])
        v[1] = int(v[1])
        v[2] = float(v[2].replace(',', '.'))
        v[3] = 0 if v[3].isalpha() else int(v[3])
        data = [int(k), yesterday] + v
        res_data.append(data)

    return res_data


def add_to_excel(data):
    # file = r"Z:\ЛОГИСТИКА\Динамика продаж\детал_ШАБЛОН.xlsx"

    wb = xw.Book(r"\\local\Шаблон\детал_ШАБЛОН.xlsb")
    # wb = xw.Book(r"детал_ШАБЛОН.xlsx")
    ws = wb.sheets['общие продажи']

    # Подготавливаем данные для таблицы и добавляем
    for i in data:
        last_row = ws.range('A' + str(ws.cells.last_cell.row)).end('up').row
        i.insert(5,
                 f'=VLOOKUP(TEXT(B{last_row + 1},"ДД.ММ.ГГГГ")&" "&A{last_row + 1},планирование!$C$2:$F$1048576,4,0)')
        i.insert(6, round(float(i[2]) / float(i[3]), 2))
        i.insert(7, round(i[4] / float(i[3]), 2))
        i.insert(8, round(i[4] / float(i[2]), 2))
        i.append(f'=IF(J{last_row + 1}=0,0,D{last_row + 1}/J{last_row + 1})')
        i.append('продажи')
        i.append(f'=TEXT("{i[1]}","ММММ")')
        i.append(f'=YEAR("{i[1]}")')
        i.append(f'=DAY("{i[1]}")')
        i.append(f'=TRUNC(MOD("{i[1]}"+3-WEEKDAY("{i[1]}",2),365.25)/7+1)')
        i.append(f'=WEEKDAY("{i[1]}",2)')
        i.append(f'=E{last_row + 1}-F{last_row + 1}')
        i.append(f'=E{last_row + 1}/F{last_row + 1}*100%')
        i.append('')
        i.append(f'=INT((MONTH("{i[1]}")+2)/3)')
        ws.range(f'A{last_row + 1}').value = i

        print(i[:5])

    # Обновляем сводную "Выполнение плана"
    last_row = ws.range('A' + str(ws.cells.last_cell.row)).end('up').row
    wb.sheets['Выполнение плана'].select()
    wb.api.ActiveSheet.PivotTables('Динамика').PivotCache().Refresh()
    wb.sheets['Выполнение плана'].range('B3').value = wb.sheets['общие продажи'].range(f'O{last_row}').value
    wb.sheets['Выполнение плана'].range('B2').value = wb.sheets['общие продажи'].range(f'M{last_row}').value

    # Добавялем данные по полу
    last_row = wb.sheets['по сексу'].range('A' + str(wb.sheets['по сексу'].cells.last_cell.row)).end('up').row
    wb.sheets['по сексу'].range(f'A{last_row + 1}').value = attach_mail_to_excel().values

    wb.save()
    wb.close()
    print('Шаблон готов!')


if __name__ == '__main__':
    # print(chek_data())
    add_to_excel(chek_data())
    # attach_mail_to_excel()
