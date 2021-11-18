import requests
from bs4 import BeautifulSoup

headers = {
    "user-agent": "Mozilla/5.0 (Windows NT 10.0 Win64 x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.69 Safari/537.36"}


url = 'https://goodinfo.tw/StockInfo/StockDetail.asp?STOCK_ID={}'.format(
    3141)
print('url:', url)
response = requests.get(url, headers=headers)
print('response:', response)
response.encoding = 'utf8'
soup = BeautifulSoup(response.text, 'lxml')
print('soup:', soup)
table = soup.find('table', class_='b1 p4_2 r10')
print('table:', table)
tr = table.find('tr', class_='bg_h1 fw_normal')
print('tr:', tr)
td = tr.find('td')
