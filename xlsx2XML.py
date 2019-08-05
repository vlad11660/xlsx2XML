import xlrd, sys, re, datetime, os
from bs4 import BeautifulSoup


def init_param(arg):
    global index_id, index_URL, index_price, index_categoryId, index_vendor, index_stock_quantity, index_name, \
    index_description, index_param, index_picture, index_article, index_price_old, index_Currency

    try:
        rb = xlrd.open_workbook('price.xlsx')
        sheet = rb.sheet_by_index(arg)
    except:
        print('Не найден файл с именем price.\n Расширение файла должно быть xlsx (эксель таблица)')
        input()
        if os.path.exists('rozetka.xml'): os.remove("rozetka.xml")
    count_param = len([sheet.row_values(rownum) for rownum in range(sheet.nrows)][0])
    index_picture=[]
    index_param=[]
    index_URL = None
    index_price = None
    index_categoryId = None
    index_vendor = None
    index_stock_quantity = None
    index_name = None
    index_description = None
    index_id = None
    index_price_old = None
    index_article = None
    index_Currency = None
    for row in range(0, count_param):
        val = (sheet.row_values(0)[row])
        if val.find('id') >=0: index_id = row
        if val.find('URL') >= 0: index_URL = row
        if val == 'price': index_price = row
        if val == 'price_old': index_price_old = row
        if val.find('categoryId') >= 0: index_categoryId = row
        if val.find('vendor') >= 0: index_vendor = row
        if val.find('stock_quantity') >= 0: index_stock_quantity = row
        if val.find('product name') >= 0: index_name = row
        if val.find('description') >= 0: index_description = row
        if val.find('article') >= 0: index_article = row
        if val.find('Валюта') >= 0: index_Currency = row
        if val.find('param name') >= 0: index_param.append(row)
        if val.find('picture') >= 0: index_picture.append(row)

    if index_id == None:
        print('Укажите параметр id(уникальный номер в пределах всего XML файла) в таблице')
        input()
        sys.exit()
    if index_URL == None:
        print('Укажите параметр URL(ссылка на товар) в таблице')
        input()
        sys.exit()
    if index_price == None:
        print('Укажите парметр price(стоимость) в таблице')
        input()
        sys.exit()
    if index_price_old == None:
        print('Укажите параметр price_old(цена со скидкой) в таблице. \n Если у Вас отсутствует данное параметр, то оставьте поле со значением пустым.')
        input()
        sys.exit()
    if index_categoryId == None:
        print('Укажите парметр categoryId(номер категории товара) в таблице')
        input()
        sys.exit()
    if index_vendor == None:
        print('Укажите параметр vendor в таблице')
        input()
        sys.exit()
    if index_stock_quantity == None:
        print('Укажите параметр stock_quantity(кол-во товара в наличии) в таблице')
        input()
        sys.exit()
    if index_name == None:
        print('Укажите парметр product name ( имя товара ) в таблице')
        input()
        sys.exit()
    if index_Currency == None:
        print('Укажите параметр Валюта. Возможные значения параметра UAH , USD , EUR ')
        input()
        sys.exit()
    if index_description == None:
        print('Укажите параметр description в таблице, а так же значение - описание товара')
        input()
        sys.exit()
    if len(index_picture) == 0:
        print('Укажите параметр picture(ссылка на фотографию) в таблице. \n Минимум одна фотография.')
        input()
        sys.exit()
    if len(index_param) == 0:
        print('Укажите параметр: param name="здесь ваш парметр".\n\nВозможные варианты параметров смотрите на сайте розетка,\
в категории вашего товара - в левом блоке сайта.\n Пример: param name="Cтрана производитель" или param name="Цвет"')
        input()
        sys.exit()
    if index_article == None:
        print('Укажите параметр article ( артикульный номер ) в таблице. \n Если у Вас отсутствует данное параметр, то оставьте поле со значением пустым.')
        input()
        sys.exit()

def edit(text):
    flag = False
    s = text.replace('&', '&amp;')
    s2 = s.replace('"', '&quot;')
    s5 = s2.replace('>', '&gt;')
    s6 = s5.replace('<', '&lt;')
    s7 = s6.replace("'", '&apos;')
    s8 = s7.replace("&lt;br/&gt;", '<br/>')
    s9 = s8.replace("&lt;br&gt;", '<br/>')

    if s9.find('&') >= 0 :
        return '<![CDATA[' + str(s9) + ']]>'
    if s9.find('<br/>') >= 0 :
        return '<![CDATA[' + str(s9) + ']]>'
    else:
        return str(s9)



def edit_name(text):
    flag = False
    s = text.replace('&', '&amp;')
    s2 = s.replace('"', '&quot;')
    s5 = s2.replace('>', '&gt;')
    s6 = s5.replace('<', '&lt;')
    s7 = s6.replace("'", '&apos;')
    if s7.find('&') >= 0 : flag=True
    return s7, flag


def edit_description(description):
    notag1 = re.sub("<a.*?</a>", "", description)
    notag2 = re.sub("<img.*?>|</img>", "", notag1)
    notag3 = re.sub("<iframe.*?>|</iframe>", "", notag2)
    notag4 = re.sub("www.*?ua", "", notag3)
    notag5 = re.sub("www.*?com", "", notag4)
    notag6 = re.sub("www.*?ru", "", notag5)

    text1 = BeautifulSoup(notag6, 'html.parser')
    text2 = text1.prettify(formatter=lambda t: t.replace('&', '&amp;'))

    text3 = BeautifulSoup(text2, 'html.parser')
    text4 = text3.prettify(formatter=lambda t: t.replace('"', '&quot;'))

    text5 = BeautifulSoup(text4, 'html.parser')
    text6 = text5.prettify(formatter=lambda t: t.replace('>', '&gt;'))
    text7 = BeautifulSoup(text6, 'html.parser')
    text8 = text7.prettify(formatter=lambda t: t.replace('<', '&lt;'))
    text9 = BeautifulSoup(text8, 'html.parser')
    text10 = text9.prettify(formatter=lambda t: t.replace("'", '&apos;'))
    return notag6



def write_text_txt(text):
    f=open('rozetka.xml', 'a', encoding='UTF-8')
    f.write(text)

def write_file(index, param, teg, art, i, stok):
    # print(type(index_id), type(index))
    try:
        q = 2.3
        qw = ''
        if index_id == index and int(stok) <=0:
            print('\n\n<offer id="' + str(int(param)) + '" available="false">')
            write_text_txt('\n\n<offer id="' + str(int(param)) + '" available="false">\n')
        if index_id == index and int(stok) >0:
            print('\n\n<offer id="' + str(int(param)) + '" available="true">')
            write_text_txt('\n\n<offer id="' + str(int(param)) + '" available="true">\n')
        if index_URL == index:
            print('<url>'+ str(param) +'</url>')
            write_text_txt('<url>' + str(param) + '</url>\n')
        if index_price == index:
            qw = ''
            if type(param) == type(qw):
                if len(param) >= 1:
                    print('<price>' + str(float(param)) + '</price>')
                    write_text_txt('<price>' + str(float(param)) + '</price>\n')
            elif type(param) == type(q):
                print('<price>' + str(float(param)) + '</price>')
                write_text_txt('<price>' + str(float(param)) + '</price>\n')
            else:
                print()
        if index_price_old == index:
            if type(param) == type(qw):
                if len(param)>=1:
                    print('<price_old>' + str(float(param)) + '</price_old>')
                    write_text_txt('<price_old>' + str(float(param)) + '</price_old>\n')
            elif type(param) == type(q):
                print('<price_old>' + str(float(param)) + '</price_old>')
                write_text_txt('<price_old>' + str(float(param)) + '</price_old>\n')
            else:
                pass
        if index_Currency == index:
            print('<currencyId>'+ str(param) +'</currencyId>')
            write_text_txt('<currencyId>' + str(param) + '</currencyId>\n')
        if index_categoryId == index:
            print('<categoryId>' + str(int(param)) + '</categoryId>')
            write_text_txt('<categoryId>' + str(int(param)) + '</categoryId>\n')
        if index_vendor == index:
            print('<vendor>'+ edit(str(param)) +'</vendor>')
            write_text_txt('<vendor>' + edit(str(param)) + '</vendor>\n')

        for i in index_picture:
            if int(i) == int(index):
                if len(str(param)) >= 128:
                    print('Ссылка должна быть не больше 128 символов! Запрещенно:', str(param))
                    input()
                    if os.path.exists('rozetka.xml'): os.remove("rozetka.xml")
                    sys.exit()
                if len(str(param))>=3:
                    print('<picture>'+ str(param) +'</picture>')
                    write_text_txt('<picture>' + str(param) + '</picture>\n')
        if index_stock_quantity == index:
            print('<stock_quantity>'+ str(int(param)) +'</stock_quantity>')
            write_text_txt('<stock_quantity>' + str(int(param)) + '</stock_quantity>\n')
        if index_name == index:
            if type(art) == type(q): art=str(int(art))  # проерка типов isinstance(q, float):
            else: art=str(art)
            try:
                if art[0] != '(' and art[-1] !=')':
                    art='('+ art + ')'
            except: pass

            if edit_name(str(param))[1] == True:
                param =edit_name(str(str(param).strip() + ' ' + art.strip()))[0]
                print('<name><![CDATA[' + param + ']]></name>')
                write_text_txt('<name><![CDATA[' + param + ']]></name>\n')
            else:
                print('<name>'+  param.strip() + ' ' + art.strip() + '</name>')
                write_text_txt('<name>' +  param.strip() + ' ' + art.strip() + '</name>\n')

        if index_description == index:
            print('<description><![CDATA['+ edit_description(str(param)) +'</description>')
            write_text_txt('<description><![CDATA[' +  edit_description(str(param)) + ']]></description>\n')


        for i in index_param:
            if int(i) == int(index):
                if len(str(param))>=1:
                    if type(param) == type(q):
                        print('<' + str(teg) + '>' + str(int(param)) + '</param>\n')
                        write_text_txt('<' + str(teg) + '>' + str(int(param)) + '</param>\n')
                    else:
                        print('<'+ teg+ '>'+ edit(str(param)) +'</param>')
                        write_text_txt('<' + teg + '>' + edit(str(param)) + '</param>\n')
    except:
        print('____________________________________________________')
        print( ' Строка: ', i, '\n Столбец: ', index+1, '\nНе верное значение в таблице: ', param)
        input()
        if os.path.exists('rozetka.xml'): os.remove("rozetka.xml")
        sys.exit()











def read_file(arg):
    try:
        rb = xlrd.open_workbook('price.xlsx')
        sheet = rb.sheet_by_index(arg)
    except:
        print('Не найден файл с именем price.\n Расширение файла должно быть xlsx (эксель таблица)')
        input()
        if os.path.exists('rozetka.xml'): os.remove("rozetka.xml")
    count_param = len([sheet.row_values(rownum) for rownum in range(sheet.nrows)][0])
    count_line = len([sheet.row_values(rownum) for rownum in range(sheet.nrows)])
    k=False
    for i in range(1, count_line):
        if(k):
            print('</offer>\n')
            # write_text_txt('</offer>\n')
        k=True
        for row in range(0, count_param):
            val =(sheet.row_values(i)[row])
            teg = sheet.row_values(0)[row]
            art = sheet.row_values(i)[index_article]
            stok = sheet.row_values(i)[index_stock_quantity]
            # print(row)
            write_file(int(row), val, str(teg), art, i, stok)
        print('</offer>\n')
        write_text_txt('</offer>\n')


def categories(arg):
    print(arg)
    rb = xlrd.open_workbook('price.xlsx')
    sheet = rb.sheet_by_index(arg)
    count_line = len([sheet.row_values(rownum) for rownum in range(sheet.nrows)])

    write_text_txt('<categories>\n')
    id_parent = sheet.row_values(0)[2]
    for i in range(count_line):
        try:
            name_category = (sheet.row_values(i)[0])
            id_category = int(sheet.row_values(i)[1])
            id_parent = sheet.row_values(i)[2]
            if type(id_parent) is str:
                category = '<category id="' + str(id_category) + '">' + str(name_category) + '</category>\n'
                write_text_txt(category)
            if len(str(id_parent)) >=1 and int(id_parent)>=0:
                parentId = '<category id="' + str(id_category) + '"' + ' parentId="' + str(int(id_parent)) + '">' + str(name_category) + '</category>\n'
                write_text_txt(parentId)
                print(name_category, id_category, int(id_parent))

            # if (sheet.row_values(i)[1]) == 'parentId':
            #     name_parentId = (sheet.row_values(i)[2])
            #     id_parent = int(sheet.row_values(i)[0])
            #     id_parent2 = int(sheet.row_values(i)[3])
            #
            #     parentId = '<category id="' + str(id_parent) + '"' + ' parentId="' + str(
            #         id_parent2) + '">' + name_parentId + '</category>\n'
            #     print(parentId)
            #     write_text_txt(parentId)
        except:
            print('В описании категорий( последний лист ) найдена ошибка! Исправьте и запустите вновь программу.')
            input()
            if os.path.exists('rozetka.xml'): os.remove("rozetka.xml")
    write_text_txt('</categories>\n<offers>')


def name_magazin():
    try:
        simbol =['+','=','[',']',':',';','«',',','.','/','?','/',':','*','?','«','<','>','|', '\\']

        rb = xlrd.open_workbook('price.xlsx')
        sheet = rb.sheet_by_index(0)
        name = (sheet.row_values(0)[1])
        company = (sheet.row_values(1)[1])
        url = (sheet.row_values(2)[1])
        cursUSD = (sheet.row_values(3)[1])
        cursEUR = (sheet.row_values(4)[1])
        cursRUB = (sheet.row_values(5)[1])
        name_file = (sheet.row_values(6)[1])
        for s in simbol:
            if s in name_file:
                print('Лист №1. Ошибка в имени файла (строка 7) . Запрещенно использовать знак: ', s)
                if os.path.exists('rozetka.xml'): os.remove("rozetka.xml")
                input()
                sys.exit()
        info = '<name>'+name+'</name>\n' + '<company>'+company+'</company>\n' + '<url>'+url+'</url>\n' + '<currencies>\n<currency id="UAH" rate="1"/>\n<currency id="USD" rate="' + str(cursUSD) + '"/>\n<currency id="EUR" rate="' + str(cursEUR) + '"/>\n<currency id="RUB" rate="' + str(cursRUB) + '"/>\n</currencies>\n'
        write_text_txt(info)
        return name_file
    except:
        print('Лист №1 найдена ошибка! Исправьте и запустите вновь программу.')
        input()
        if os.path.exists('rozetka.xml'): os.remove("rozetka.xml")




def start_text():
    now = datetime.datetime.now()
    sys_data = str(now).split('.')[0].split(' ')

    text='''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE yml_catalog SYSTEM "shops.dtd">
<yml_catalog date="''' + str(sys_data[0]) + ' ' +str(sys_data[1][:-3]) +'''">
<shop>
'''

    write_text_txt(text)
    name_file_xml = name_magazin()
    return name_file_xml

def end_text():
    text='''</offers>
</shop>
</yml_catalog>'''
    write_text_txt(text)

def main():
    rb = xlrd.open_workbook('price.xlsx')
    count_sheets = rb.nsheets

    for sheet in range(count_sheets):
        if sheet == 0:
            ffile_name = start_text()
            categories(count_sheets-1)
        if sheet !=0 and sheet!=count_sheets-1:
            init_param(sheet)
            read_file(sheet)
    print('</offer>\n')
    # write_text_txt('</offer>\n')
    end_text()
    if os.path.exists(ffile_name + '.xml'): os.remove(ffile_name + ".xml")
    os.rename('rozetka.xml', str(ffile_name) + '.xml')
    x = input('Файл готов!!')


if __name__ == '__main__':

    if os.path.exists('rozetka.xml'): os.remove("rozetka.xml")
    main()
