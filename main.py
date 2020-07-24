import requests
from bs4 import BeautifulSoup
import pandas as pd
from functools import cmp_to_key
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

url = "https://abit.itmo.ru/bachelor/rating_rank/all/261/"

last_condition = ''


def cmp_items(a, b):
    def convert(l):
        condtion_key = {'без вступительных испытаний': 4,
                        'на бюджетное место в пределах особой квоты': 3,
                        'на бюджетное место в пределах целевой квоты': 2,
                        'по общему конкурсу': 1,
                        'на контрактной основе': 0}
        l[0] = condtion_key[l[0]]
        l[8] = int(l[8] or 0)
        l[11] = 1 if l[11] == 'Да' else 0
        l[12] = 1 if l[12] == 'Да' else 0
        return l
    a = convert(a.copy()); b = convert(b.copy())
    r = 0

    if a[11] > b[11]: r = -1 # Наличие согласия на зачисление
    elif a[11] < b[11]: r = 1
    else:
        if a[0] > b[0]: r = -1 # Условие поступления (бви, контракт ...)
        elif a[0] < b[0]: r = 1
        else:
            if a[12] > b[12]: r = -1 # Преимущественное право
            elif a[12] > b[12]: r = 1
            else:
                if a[8] > b[8]: r = -1 # ЕГЭ+ИД
                elif a[8] < b[8]: r = 1
                else: r = 0

    return r


def parse_row(row):
    global last_condition
    cells = row.find_all('td')

    if len(cells) == 15:
        last_condition = cells[0].getText()
        cells = row.find_all('td', {'rowspan': None})

    condition = last_condition
    number_1 = cells[0].getText()
    number_2 = cells[1].getText()
    full_name = cells[2].getText()

    mode = cells[3].getText()
    m = cells[4].getText()
    r = cells[5].getText()
    i = cells[6].getText()

    exam_and_ia = cells[7].getText()
    exam = cells[8].getText()
    ia = cells[9].getText()

    agreement = cells[10].getText()
    advantage = cells[11].getText()
    olympiad = cells[12].getText()
    status = cells[13].getText()

    res = [condition, number_1, number_2, full_name,
           mode, m, r, i,
           exam_and_ia, exam, ia,
           agreement, advantage, olympiad, status]

    return res


def main():
    print('Скачиваю страницу:',url)
    r = requests.get(url, verify=False)  # получаем страницу
    print('Ищу таблицу')
    soup = BeautifulSoup(r.text, features='html.parser')  # парсим таблицу
    rows = soup.find_all('tr', {'class': None})  # получаем строки
    print('Начинаю парсить таблицу')
    result = []
    for row in rows:
        res = parse_row(row)
        result.append(res)
    print('Ранжирую таблицу')
    result = sorted(result, key=cmp_to_key(cmp_items))
    print('Вывожу таблицу в файл')
    result = pd.DataFrame(result,
                          columns=['condition', 'number_1', 'number_2', 'full_name',
                                   'mode', 'm', 'r', 'i',
                                   'exam_and_ia', 'exam', 'ia',
                                   'agreement', 'advantage', 'olympiad', 'status'])
    result.to_excel('result.xlsx')
    print('Готово!')


if __name__ == '__main__':
    main()