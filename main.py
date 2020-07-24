import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

url = "https://abit.itmo.ru/bachelor/rating_rank/all/261/"
required_name = 'Шараев Павел Ильдарович'


def write_to_file(result):
    global required_name
    import xlsxwriter
    with xlsxwriter.Workbook('result.xlsx') as workbook:
        worksheet = workbook.add_worksheet('Таблица')
        worksheet.write_row(0, 0, ['Номер', 'Номер в конкурсной группе', 'Условие поступления', '№ п/п',
                                   'Номер заявления', 'ФИО', 'Вид', 'М', 'Р', 'И', 'ЕГЭ+ИД', 'ЕГЭ', 'ИД',
                                   'Наличие согласия на зачисление', 'Преимущественное право', 'Олимпиада', 'Статус'])
        data_format1 = workbook.add_format({'bg_color': '#16de69'})

        gray = workbook.add_format({'bg_color': '#dbdbdb'})
        white = workbook.add_format({'bg_color': '#ffffff'})
        current_color = gray
        last_color = white

        last_cond = result[0][0]
        j = 1
        for i in range(len(result)):
            if i > 0 and result[i][0] == last_cond:
                j += 1
            else:
                j = 1
                current_color, last_color = last_color, current_color
                last_cond = result[i][0]
            if required_name == result[i][3]:
                worksheet.write_row(i + 1, 0, [i + 1, j] + result[i], data_format1)
            else:
                worksheet.write_row(i + 1, 0, [i + 1, j] + result[i], current_color)


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

    a = convert(a.copy());
    b = convert(b.copy())
    r = 0

    if a[11] > b[11]:
        r = -1  # Наличие согласия на зачисление
    elif a[11] < b[11]:
        r = 1
    else:
        if a[0] > b[0]:
            r = -1  # Условие поступления (бви, контракт ...)
        elif a[0] < b[0]:
            r = 1
        else:
            if a[12] > b[12]:
                r = -1  # Преимущественное право
            elif a[12] > b[12]:
                r = 1
            else:
                if a[8] > b[8]:
                    r = -1  # ЕГЭ+ИД
                elif a[8] < b[8]:
                    r = 1
                else:
                    r = 0

    return r

last_condition = ''
def parse_row(row):
    def to_int_possible(a):
        try:
            r = int(a)
        except:
            r = ''
        return r
    global last_condition
    cells = row.find_all('td')

    if len(cells) == 15:
        last_condition = cells[0].getText()
        cells = row.find_all('td', {'rowspan': None})

    condition = last_condition
    number_1 = int(cells[0].getText())
    number_2 = int(cells[1].getText())
    full_name = cells[2].getText()

    mode = cells[3].getText()
    m = to_int_possible(cells[4].getText())
    r = to_int_possible(cells[5].getText())
    i = to_int_possible(cells[6].getText())

    exam_and_ia = to_int_possible(cells[7].getText())
    exam = to_int_possible(cells[8].getText())
    ia = to_int_possible(cells[9].getText())

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
    print('Скачиваю страницу:', url)
    import requests
    r = requests.get(url, verify=False)  # получаем страницу

    print('Ищу таблицу')
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(r.text, features='html.parser')  # парсим таблицу
    rows = soup.find_all('tr', {'class': None})  # получаем строки

    print('Начинаю парсить таблицу')
    result = []
    for row in rows:
        res = parse_row(row)
        result.append(res)

    print('Ранжирую таблицу')
    from functools import cmp_to_key
    result = sorted(result, key=cmp_to_key(cmp_items))

    print('Вывожу таблицу в файл')
    write_to_file(result)

    print('Готово!')


if __name__ == '__main__':
    main()
