__author__ = 'Quan'

from bs4 import BeautifulSoup as bs
import requests as r
import openpyxl as x

fn='D:\LinhTinh\Code\PythonProject\sample.xlsx'
wb=x.Workbook()
wb = x.load_workbook(fn)

cafef= r.get('http://cafef.vn/')
cafef=bs(cafef.content)
cafefnews=cafef.find_all("div",{"class": "right"})

cafefnews=cafefnews[1] # chon div class right thu 2 trong tat cac cach div class right
lis=cafefnews.find_all("li")
li_content=[ele.text.strip() for ele in lis]
ws=wb.worksheets[0]
for i in range(len(li_content)):
    ws.cell(row = i + 1, column = 1).value = li_content[i]
wb.save(fn)


vietstock= r.get('http://vietstock.vn/')
vietstock=bs(vietstock.content)
vietstocknews=vietstock.find_all("ul",{"class": "ui-tabs-nav"})

#vietstocknews=vietstocknews[1] # chon div class right thu 2 trong tat cac cach div class right
lis=vietstocknews.find_all("li")
li_content=[ele.text.strip() for ele in lis]
ws=wb.worksheets[0]
for i in range(len(li_content)):
    ws.cell(row = i + 10, column = 1).value = li_content[i]
wb.save(fn)
