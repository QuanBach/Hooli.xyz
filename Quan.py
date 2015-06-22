__author__ = 'Quan'

from bs4 import BeautifulSoup as bs
import requests as r

cafef = r.get('http://s.cafef.vn/du-lieu.chn#data') #gui request len sever cafef
cafef = bs(cafef.content)# chuyen du lieu ve dang beautiful
header = cafef.find("div", {"id": "header-time"})# header: kim 1 cai div co id la` "header-time"
xyz = cafef.find_all("div", {"id": "stockindex"})# kiem toan bo cac div co "id"=stockindex", xyz la` 1 list div
for xy in xyz:
    print (xy.prettify().encode('utf-8'))
    print(11111111111111111111111111111111111111111111111)

table = xyz[1].find('table')# ki?m table trong  th? 2 trong xyz


data = []
rows = table.find_all('tr')# kiem trong table nhung cai 'tr'

for i in range(len(rows)):
    cols = rows[i].find_all('td')
    cols = [ele.text.strip() for ele in cols]
    if i % 3 == 1:
        data.append(cols)
print(data)

