import requests
from bs4 import BeautifulSoup
import xlrd, xlwt
import string

wb = xlwt.Workbook()
ws = wb.add_sheet('Test')
ws.write(0,0, "Оператор связи")
ws.write(0,1, "Номер заключения")
ws.write(0,2, "Дата заключения")
ws.write(0,3, "Координаты")
ws.write(0,4, "Стандарт связи")
ws.write(0,5, "Адрес размещения")

def get_html(url): #Получение кода страницы на html
    r = requests.get(url)
    return r.text

def get_amount_of_pages(html): #Возвращает общее количество страниц
    soup = BeautifulSoup(html, 'lxml')
    res = str(soup.find('noindex'))
    pattern = "Страницы (всего "
    cur = res.find("Страницы (всего ")
    if cur != -1:
        cur += len(pattern)
        end = res.find(")", cur)
        if end != -1:
            return int(res[cur:end].strip())

def parse_operator(operator_string): # Вынимает из строки оператора, т.е. Владельца, если он явно задан
    txt = ""
    operator_pattern = "Владелец"
    cur = operator_string.find(operator_pattern, 0)
    if cur != -1:
        cur += len(operator_pattern) + 1
        cur = operator_string.find(":", cur)
        if cur != -1:
            cur += 2
            next_q = operator_string.find("\"", cur)
            if next_q != -1:
                next_q = operator_string.find("\"", next_q + 1)
                if next_q != -1:
                    txt = (operator_string[cur : next_q + 1])
    return txt

def parse_date(date_string): #Вынимает из строки дату и номер заключения, основываясь на паттерне *номер* от *дата*
    dxt = ""
    zxt = ""
    date_pattern = " от "
    date_length = 10
    cur = date_string.find(date_pattern, 0)
    if cur != -1:
        if (date_string[cur + len(date_pattern) +1].isdigit()):
            dxt = (date_string[cur + len(date_pattern) : cur + len(date_pattern) + date_length]) #Т.к. длина даты фиксированная, вырезаем необходимое количество символов
            zxt = (date_string[0 : cur])
    return zxt, dxt

def parse_standard(stand_string): # Вынимаем из строки стандарт, если он состоит латинских символов, цифр и разделителей, тоже только если он явно задан
    txt = ""
    stand_pattern = "тандарт"
    standard_alphabet = string.ascii_letters + string.digits + " и-,/+"  #Список символов, которые могут являться частью стандарта
    cur = stand_string.find(stand_pattern, 0)
    if cur != -1:
        cur += len(stand_pattern)
        cur = stand_string.find(" ", cur)
        if cur != -1:
            cur += 1
            end = cur
            while(stand_string[end] in standard_alphabet) & (end < len(stand_string) -1):
                end += 1
            txt = stand_string[cur: end]
    return txt

def go_to_addition(addition): # Переход по ссылке "показать полный текст приложения", чтобы извлечь адрес
    base_url = 'http://fp.crc.ru/service/?oper=s&uinz'
    pattern = "uinz"
    cur = addition.find(pattern, 0)
    if cur != -1:
        cur += len(pattern)
        end = addition.find("\"", cur)
        if end != -1:
            res = BeautifulSoup(get_html(base_url + addition[cur : end]), 'lxml').find_all('tr')
            mass = []
            for item in res:
                smaller_item = item.find_all('td')
                mass.append(smaller_item)
            return str(mass)


def parse_address(add_string): #извлечение адреса из хтмла, полученного при дополнительном переходе по ссылке
    txt = ""
    add_pattern = "г. "
    try:
        cur = add_string.find(add_pattern, 0)
    except:
        cur = -1
    if cur != -1:
        end = add_string.find("<", cur + len(add_pattern))
        txt = add_string[cur: end].strip()
    return txt


def parse_page(html, i): #Функция, в которой вызываются все остальные функции парсинга
    soup = BeautifulSoup(html, 'lxml')
    res = soup.find('noindex').find_all('td')
    for item in res:
        try:
            operator = item.find('b').text
        except:
            operator = ''
        try:
            addition = item.find('a')
        except:
            addition = ''
        addr = item.find_all('td')
        date = item.text.strip()
        z, d = parse_date(operator)
        p = parse_operator(operator)
        s = parse_standard(operator)
        parsed_addition = go_to_addition(str(addition))
        a = parse_address(parsed_addition)
        base_url = 'https://geocode-maps.yandex.ru/1.x/?apikey=fd6de81a-3bf0-4655-a8c3-d02faa88add2&geocode=%D0%B3'
        response = get_html(base_url + a)
        cur = response.find("<pos>", 0)
        if cur != -1:
            cur += len("<pos>")
            end = response.find("</pos>", cur)
            if (end != -1) & (a != ""):
                ws.write(i-1, 3,(response[cur:end]))
        if d != "":
            ws.write(i, 2, d)
        if z != "":
            ws.write(i, 1, z)
            i += 1
        if p != "":
            ws.write(i-1, 0, p)
        if s != "":
            ws.write(i-1, 4, s)
        if a != "":
            try:
                ws.write(i-1, 5, a)
            except:
                ws.write(i-1, 6, a) # Т.к. адресов бывает несколько, записываю лишний в запасное поле
    return i

if __name__ == '__main__':
    pg = 1
    url = 'http://fp.crc.ru/service/?pg=1&oper=s&rpp=25&type=max&text_prodnm=%E1%E0%E7%EE%E2%E0%FF%20%F1%F2%E0%ED%F6%E8%FF&use=0'
    start_url = 'http://fp.crc.ru/service/?pg='
    max_page = get_amount_of_pages(get_html(url))
    end_url = '&oper=s&rpp=25&type=max&text_prodnm=%E1%E0%E7%EE%E2%E0%FF%20%F1%F2%E0%ED%F6%E8%FF&use=0'
    info = 1
    while (pg < max_page):
        new_url = start_url + str(pg) + end_url
        info = parse_page(get_html(new_url), info)
        pg += 1
    wb.save('xl_rec.xls')
