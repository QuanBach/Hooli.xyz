
import tkFileDialog as tkfd
import Tkinter as tk
import requests as r
from lxml import html
from bs4 import BeautifulSoup as bs
import openpyxl as x


fn=''
wb = x.Workbook()

def getFileName():
    global fn
    global wb
    fn = tkfd.askopenfilename(filetypes=[('xlsx files', '.xlsx'), ('all files', '.*')], defaultextension='.xlsx')
    txtBox.delete(0, tk.END)
    txtBox.insert(0, fn)
    wb = x.load_workbook(fn)

def getBasicPrice():
    global wb
    global fn
    print(wb.get_sheet_names())

    ws = wb.worksheets[0]
    try:
        vs = r.get('http://vietstock.vn/')
    except r.ConnectionError:
        return False
    vs_html = html.fromstring(vs.content)

    vnindex = vs_html.get_element_by_id('PRICE_VNIndex')
    vnindex = vnindex.text_content()
    ws['C5'] = vnindex

    hnxindex = vs_html.get_element_by_id('PRICE_HastcIndex')
    hnxindex = hnxindex.text_content()
    ws['C6'] = hnxindex

    oil = vs_html.get_element_by_id('PRICE_CL')
    oil = oil.text_content()
    ws['C7'] = oil

    wb.save(fn)
    return True

def getExchangeRate():
    global wb
    global fn

    # cafef_html = html.fromstring(cafef.content)
    # er = cafef_html.get_element_by_id('div_tygia')
    # print(type(er.text_content()))
    ws = wb.worksheets[0]
    try:
        cafef = r.get('http://cafef.vn/')
    except r.ConnectionError:
        return False
    soup = bs(cafef.content)
    div = soup.find('div', {'id':'div_tygia'})
    table_body = div.find('tbody')
    data = []
    rows = table_body.find_all('tr')
    for i in range(len(rows)):
        cols = rows[i].find_all('td')
        cols = [ele.text.strip() for ele in cols]
        for j in range(len(cols)):
            ws.cell(row=10+i, column=2+j).value = cols[j]
        data.append([ele for ele in cols if ele]) # Get rid of empty values
    wb.save(fn)
    return True

def getFinanceReport():
    try:
        fr = r.get('http://s.cafef.vn/bao-cao-tai-chinh/DPM/IncSta/2050/1/0/0/ket-qua-hoat-dong-kinh-doanh-tong-cong-ty-phan-bon-va-hoa-chat-dau-khictcp.chn')
    except r.ConnectionError:
        return False
    soup = bs(fr.content)
    table_header = soup.find('table', {'id':'tblGridData'})
    rows = table_header.find_all('tr')
    ws = wb.worksheets[1]
    for i in range(len(rows)):
        cols = rows[i].find_all('td')
        cols = [ele.text.strip() for ele in cols]
        for j in range(len(cols)):
            ws.cell(row=10+i, column=2+j).value = cols[j]

    table_content = soup.find('table', {'id':'tableContent'})
    rows = table_content.find_all('tr')
    rowid = 0
    for i in range(len(rows)):
        cols = rows[i].find_all('td')
        cols = [ele.text.strip() for ele in cols]
        if i % 3 == 0:
            for j in range(len(cols)):
                cols[j] = cols[j].replace(',', '')
                try:
                    cols[j]=int(cols[j])
                except ValueError:
                    pass
                ws.cell(row=11+rowid, column=2+j).value=cols[j]
            rowid+=1
    wb.save(fn)
    return True

def getNews():
    try:
        newsPage = r.get('http://s.cafef.vn/hose/DPM-tong-cong-ty-phan-bon-va-hoa-chat-dau-khictcp.chn')
    except r.ConnectionError:
        return False

    soup = bs(newsPage.content)
    div = soup.find('div', {'id':'divTopEvents'})
    lis = div.find_all('li')
    ws = wb.worksheets[1]
    li_content = [ele.text.strip() for ele in lis]
    for i in range(len(li_content)):
        ws.cell(row = i + 1, column = 1).value = li_content[i]
    wb.save(fn)
    return True

def update():
    global fn
    if fn:
        if not getBasicPrice():
            stringAnnouncement.set('Cannot connect! Check your connection!')
            return
        if not getExchangeRate():
            stringAnnouncement.set('Cannot connect! Check your connection!')
            return
        if not getFinanceReport():
            stringAnnouncement.set('Cannot connect! Check your connection!')
            return
        if not getNews():
            stringAnnouncement.set('Cannot connect! Check your connection!')
            return
        stringAnnouncement.set('Update successfully!')

root = tk.Tk()

stringFile = tk.StringVar()
stringFile.set('File: ')
stringAnnouncement = tk.StringVar()
openBtn = tk.Button(root, text='...', command=getFileName)
updateBtn = tk.Button(root, text='Update', command=update)
fileLabel = tk.Label(root, textvariable=stringFile)
txtBox = tk.Entry(root, width=50)
announcementLabel = tk.Label(root, textvariable=stringAnnouncement)

fileLabel.grid(row=0, column=0, padx=5, pady=5)
openBtn.grid(row=0, column=2, padx=5, pady=5)
txtBox.grid(row=0, column=1, padx=5, pady=5)
updateBtn.grid(row=1, column=1, padx=5, pady=5)
announcementLabel.grid(row=2, column=0, columnspan=3)

root.mainloop()

