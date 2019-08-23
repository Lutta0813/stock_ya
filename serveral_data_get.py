import requests
from bs4 import BeautifulSoup
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties as fp
from datetime import datetime
from datetime import timedelta
import time

def geturl(stock_code, start_date):
    url = 'https://www.cnyes.com/twstock/ps_historyprice/' + str(stock_code) + '.htm'
    headers = {'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Safari/537.36'}
    result = requests.Session()
    
    params = {
    'code' : '0056',
    'ctl00$ContentPlaceHolder1$startText' : start_date,
    'ctl00$ContentPlaceHolder1$submitBut:' : '查詢'
    }
    r = result.post(url, params, headers)

    # r = result.get(url, headers = headers, allow_redirects=False)
    print(r.status_code) # 200表示訪問成功

    soup = BeautifulSoup(r.content, 'html.parser')
    return soup

def getdata(soup):
    try:
        price = []
        dateContent = []
        count = 0

        # 表單父級
        trPos = soup.find_all(text = '收盤')[0].parent.parent.parent.select('tr')
        # 股票名稱
        st_name = soup.select('h1')[0].select('a')[0].text
        print(st_name)

        # 算出table data有幾筆我們需要的資料，第0筆為文字描述我們不需要
        for i in trPos[1:]:
            count += 1

        # 從第1筆開始將數據讀取出來
        for i in range(1, count):
            p = float(trPos[i].select('td')[4].text)
            t = trPos[i].select('td')[0].text
            price.append(p)
            dateContent.append(t)

        return  price, dateContent, st_name
    except:
        print('something wrong')

def main():
    st_code = input('請輸入股票代碼：')

    # time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
    # time_now = datetime.now().strftime("%Y-%m-%d")
    time_30days_ago = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")
    time_90days_ago = (datetime.now() - timedelta(days=90)).strftime("%Y-%m-%d")
    time_1yeas_ago = (datetime.now() - timedelta(days=365)).strftime("%Y-%m-%d")

    while True:
        st_date = input('請選擇起始時間：\n0：一個月\n1：三個月\n2：一年\n')
        if st_date == '0':
            price, date, stock_name = getdata(geturl(st_code, time_30days_ago))
            break
        if st_date == '1':
            price, date, stock_name = getdata(geturl(st_code, time_90days_ago))
            break
        if st_date == '2':
            price, date, stock_name = getdata(geturl(st_code, time_1yeas_ago))
            break
        else:
            print('\n---Error---\n請輸入正確數字，三秒後重試\n----------\n')
            time.sleep(3)
            continue

    # price, date, stock_name = getdata(geturl(st_code))
    data = {}

    for p, d in zip(price, date):
        data[d] = p
        print('時間：' + d + ' ' + '成交價：' + str(p))

    # 中文字體路徑
    myfont = fp(fname=r'/System/Library/Fonts/STHeiti Medium.ttc')

    plt.plot(sorted(date), sorted(price))
    plt.xticks(rotation=45)
    plt.title(stock_name, fontproperties=myfont)

    plt.show()

main()

