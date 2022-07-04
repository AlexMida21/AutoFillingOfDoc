import os
from urllib.request import urlopen
from bs4 import BeautifulSoup


def parse_okved():
    get_link = open("link.txt", encoding='utf-8')
    URL = get_link.read()
    f = urlopen(URL)
    page = f.read().decode('utf-8')
    soup = BeautifulSoup(page, 'lxml')
    osn = soup.find_all('p')
    okv = soup.find_all('table', 'td', class_='tt')
    txt_file = open("okveds.txt", "w+", encoding='utf-8')

    a = 0
    b = 0
    value = 0

    while value == 0:
        if 'Основной (по коду ОКВЭД' in str(osn[b].previous_element):
            txt_file.write(osn[b].text)
            print(osn[b].text)
            value = value + 1
        b = b + 1

    value = 0

    while value == 0:
        if 'Эта группировка включает' in str(okv[a]) or \
                'Эта группировка не включает' in str(okv[a]):
            txt_file.write(okv[a].text)
            print(okv[a].text)
            value = value + 1
        a = a + 1

    txt_file.close()
    get_link.close()
    os.remove("link.txt")


def main():
    parse_okved()


main()
