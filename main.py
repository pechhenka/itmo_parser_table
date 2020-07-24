import requests
from bs4 import BeautifulSoup
import pandas as pd
import random

url = "https://abit.itmo.ru/bachelor/rating_rank/all/261/"

last_condition = ''


def parse_row(row):
    global last_condition
    res = pd.DataFrame()
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

    res = res.append(pd.DataFrame([[condition, number_1, number_2, full_name,
                                    mode, m, r, i,
                                    exam_and_ia, exam, ia,
                                    agreement, advantage, olympiad, status]],
                                  columns=['condition', 'number_1', 'number_2', 'full_name',
                                           'mode', 'm', 'r', 'i',
                                           'exam_and_ia', 'exam', 'ia',
                                           'agreement', 'advantage', 'olympiad', 'status']),
                     ignore_index=True)

    return res


def main():
    r = requests.get(url, verify=False)  # отправляем HTTP запрос и получаем результат

    soup = BeautifulSoup(r.text, features='html.parser')  # Отправляем полученную страницу в библиотеку для парсинга
    rows = soup.find_all('tr', {'class': None})  # Получаем все таблицы с вопросами

    result = pd.DataFrame()
    for row in rows:
        res = parse_row(row)
        result = result.append(res, ignore_index=True)

    result.to_excel('result.xlsx')


if __name__ == '__main__':
    main()
