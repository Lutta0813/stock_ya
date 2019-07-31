import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

def geturl(stock_code):
	url = 'https://tw.stock.yahoo.com/q/q?s=' + stock_code
	headers = {'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Safari/537.36'}
	r = requests.get(url, headers = headers, allow_redirects=False)
	# print(r.status_code) # 200表示訪問成功
	# print(r.url)
	soup = BeautifulSoup(r.content, 'html.parser')
	return soup

def getdata(soup):
	stockDic = {}
	table = soup.find_all(text = '成交')[0].parent.parent.parent
	try:
		stockDic['股票名稱'] = table.select('td')[0].select('a')[0].text
		stockDic['時間'] = table.select('td')[1].text
		stockDic['成交價'] = table.select('td')[2].text
		stockDic['買進'] = table.select('td')[3].text
		stockDic['賣出'] = table.select('td')[4].text
		for i in table.select('td')[5].select('font')[0]:
			stockDic['漲跌'] = i.strip()
			break
		stockDic['張數'] = table.select('td')[6].text
		stockDic['昨日收盤價'] = table.select('td')[7].text
		stockDic['開盤價'] = table.select('td')[8].text
		stockDic['最高'] = table.select('td')[9].text
		stockDic['最低'] = table.select('td')[10].text

	except:
		print('cant get table data')

	print(stockDic)
	return stockDic

def write_in_excel(stock_data):
	wb = Workbook()
	ws1 = wb.active
	excel_name = stock_data['股票名稱']
	ws1.title = 'Stock Data'
	stock_title = []
	stock_value =[]
	for key,value in stock_data.items():
		try:
			stock_title.append(key)
			if value.isdigit():
				stock_value.append(int(value))
			else:
				stock_value.append(value)

		except:
			print('Something wrong')

	# 將資料放入xml中
	for row in range(0,1):
		ws1.append(stock_title)
		ws1.append(stock_value)

	# 調整column到適當的大小

	# 儲存
	wb.save(filename = excel_name + '.xlsx')

def main():
	try:
		code = input('請輸入股票代碼：')
	except ValueError:
		print('股票代碼只可以是整數數字')
	stock_data = getdata(geturl(code))
	write_in_excel(stock_data)

main()