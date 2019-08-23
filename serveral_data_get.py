import requests
from bs4 import BeautifulSoup
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties as fp

def geturl(stock_code):
    url = 'https://www.cnyes.com/twstock/ps_historyprice/' + str(stock_code) + '.htm'
    print(url)
    headers = {'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Safari/537.36'}
    result = requests.Session()
    r = result.get(url, headers = headers, allow_redirects=False)
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
    price, date, stock_name = getdata(geturl(st_code))
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

