import requests

import shutil

import os

from bs4 import BeautifulSoup as bs

from threading import Thread

from xlwt import Workbook

import time

# Function for writing data in Excel file
def write_info(row, url, name, articul, description, price, category, photos, chars):
    global file, filename, sheet

    print(f'Collected: {name}')

    sheet.write(row, 0, articul)
    sheet.write(row, 1, name)
    sheet.write(row, 2, description)
    sheet.write(row, 3, price)
    sheet.write(row, 4, category)
    sheet.write(row, 5, ', '.join(photos))
    sheet.write(row, 6, url)

    for col, prop in enumerate(chars):
        sheet.write(row, 2 * col + 7, prop)
        sheet.write(row, 2 * col + 8, chars[prop])

    try:
        file.save(f'{filename}.xls')
    except:
        pass

# Function for downloading image
def download_images(name, urls):
    name = name.replace('\ '[0], ' ').replace('/', ' ').replace('?', ' ').replace(':', ' ').replace('|', ' ').replace('!', ' ')\
               .replace('<', ' ').replace('>', ' ').replace('"', ' ').replace('*', ' ').replace(".", ' ').strip()
    os.mkdir(f'photos/{name}')

    for image_num, url in enumerate(urls):
        image_response = requests.get(url, stream = True)
        with open(f'photos/{name}/image{image_num + 1}.png', 'wb') as image_file:
            shutil.copyfileobj(image_response.raw, image_file)

# Function for getting html by the URL
def get_html(url):
    global session

    return session.get(url).text

# Function for scraping 1 product
def parse_product(url, row):
    global all_articuls

    html = get_html(url)
    soup = bs(html, 'html.parser')

    try:
        name = soup.find('h1', class_ = 'translate to_lower').text.strip().replace('?', '')
    except:
        name = ''
    try:
        articul = soup.find('div', class_ = 'item-code').strong.text.strip()
    except:
        articul = ''
    try:
        description = '\n'.join(section.text.strip() for section in soup.find_all('section', class_ = 'text-item'))
    except:
        description = ''
    try:
        category = soup.find('ul', class_ = 'breadcrumb').find_all('li')[-2].a.span.text.strip()
    except:
        category = ''
    try:
        price = soup.find('div', class_ = 'price').strong.span.text.strip()
    except:
        price = ''
    try:
        chars = {tr.find_all('td')[0].text.strip(): tr.find_all('td')[1].text.strip() for tr in soup.find('table', class_ = 'datasheet').find_all('tr')}
    except:
        chars = {}
    try:
        photos = [a.get('href') for a in soup.find('div', class_ = 'product-other-images').find_all('a')]
    except:
        photos = [soup.find('div', class_ = 'product-main-image').find_all('img')[-1].get('src')]

    if articul.lower() not in all_articuls:
        all_articuls.append(articul.lower())
        print(name + ' --- ' + str(row))
        write_info(row, url, name, articul, description, price, category, photos, chars)
        download_images(name, photos)
    else:
        print('$' * 120)

# Function for scraping 1 page
def parse_page(url, start_row):
    html = get_html(url)
    soup = bs(html, 'html.parser')
    urls = [prod_block.h3.a.get('href') for prod_block in soup.find_all('div', class_ = 'product-list')[1:]]

    for row, url in enumerate(urls):
        print(url)
        thread = Thread(target = parse_product, args = (url, start_row + row,))
        thread.start()
        time.sleep(0.5)
        #parse_product(url, row + 1)


# Creating folder for photos
try:
    shutil.rmtree('photos')
except:
    pass
time.sleep(1)
os.mkdir('photos')

# Open workbook
file = Workbook()
filename = 'list'

try:
    sheet = file.add_sheet(filename)
    file.save(f'{filename}.xls')
except:
    print(f'Please, close "{filename}.xls"\n\nPress enter to exit')
    a = input()

sheet.write(0, 0, 'Артикул')
sheet.write(0, 1, 'Имя')
sheet.write(0, 2, 'Описание')
sheet.write(0, 3, 'Цена')
sheet.write(0, 4, 'Категория')
sheet.write(0, 5, 'Изображения')
sheet.write(0, 6, 'URL')

for col in range(1, 61):
    sheet.write(0, 2 * col + 5, 'Имя атрибута ' + str(col))
    sheet.write(0, 2 * col + 6, 'Значение(-я) аттрибута(-ов) ' + str(col))

file.save(f'{filename}.xls')

# Define session for scraping
session = requests.Session()

# Count number of pages
try:
    base_url = 'https://docom.com.ua/search?string=al-ko&page=1'
    html = get_html(base_url)
    soup = bs(html, 'html.parser')
    number_of_pages = int(soup.find('ul', class_ = 'pagination').find_all('li')[-2].a.text)
    print(f'Pages: {number_of_pages}')
except:
    number_of_pages = 1

# Start scraping
all_articuls = []
for page_num in range(number_of_pages):
    url = f'https://docom.com.ua/search?string=al-ko&page={page_num + 1}'
    thread = Thread(target = parse_page, args = (url, 12 * page_num + 1,))
    thread.start()
    time.sleep(6)
    #parse_page(url, 12 * page_num + 1)
