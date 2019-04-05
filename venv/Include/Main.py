import xlrd, xlwt, openpyxl, random, requests


def url_get():
    """
    Создаём url, в который помещаем два значения в виде дат и случайное значение для для параметра t в запросе
    :return:
    """
    input_start = input('Pls input start date like 01.04.2019')
    intput_end = input('Pls input start date like 07.04.2019')
    stas = 'http://www.usue.ru//schedule/?t={0}&action=show&startDate={1}&endDate={2}&group=%D0%AD%D0%9C%D0%90-16-1'.format(
        str(random.random()),
        input_start,
        intput_end)
    return stas


def get_json_list():
    """
    Возвращаем лист оформленный в виде json, полученный после запроса, с входными параметрами headers
    :return:
    """
    url_req = url_get() #получаем url
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.86 Safari/537.36',
        'Accept': '*/*',
        'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
        'Accept-Encoding': 'gzip, deflate',
        'X-Requested-With': 'XMLHttpRequest',
        'Referer': 'http://www.usue.ru/raspisanie/'} #значения, который отправляются вместе с запросам для его успешного завершения (например, ссылка на страницу с которой ранее мы перешли, чтобы сделать запрос)
    newresp2 = requests.get(url_req, headers=headers) #отправляем запрос
    return newresp2.json() #возвращает лист оформленный в виде json


def excel_create():
    excel_create.file_way = r'' + input('''
    Pls enter ur xlsx full path way  
    WARNING: FILE MUST EXIST
    :''') #записываем путь к файлу Excel в атрибут функции
    #file_way = r'C:\SickoMode\python\nova.xlsx'
    wb = openpyxl.load_workbook(filename=excel_create.file_way)
    return wb

def excel_work():
    """

    :param json_list: принимаем json list
    :return: функция является 'процедурой' которая использует готовый xlsx файл и записывает в него полученные значения
    """
    counter = 1 #счётчик номера строки, значени 1 , так как ячейки начинаются с 1
    #file_way = r'' + input()
    json_list = get_json_list() # получаем лист
    work_book = excel_create() #указываем
    actlist = work_book.active # указываем активную ячейку
    actlist.title = "Schedule" #задаём её название
    #Присваиваем зачения из джейсона ячейкам таблицы Excel
    for days in json_list:
        for pairs in days:
            actlist.cell(row=counter, column=3, value=days['date']) #вносим в 3 колонну дату
            if pairs != 'pairs':
                pass
            else:
                for x in days[pairs]:
                    if x['schedulePairs'] != []:
                        cake = x['schedulePairs'][0]
                        for js in cake:
                            print(cake[js])
                            actlist['B{0}'.format(counter)] = cake[js] #записываем информацию о паре
                            actlist['A{0}'.format(counter)] = x['time'] #записываем время пары
                            counter += 1
    work_book.save(excel_create.file_way)
    return "the way to ur file: " + excel_create.file_way



def main():
    excel_work()


if __name__ == '__main__':
    main()