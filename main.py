import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment

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

    # print(stockDic)
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

    # 儲存
    wb.save(filename = excel_name + '.xlsx')

    return excel_name, ws1.title

def fixcolumnwidth(excel_name, title):
    wb = load_workbook(filename = excel_name + '.xlsx')
    sheet_range = wb['Stock Data'] # 選擇Stock Data這個Sheet
    col_widths_dict = {}

    # 顯示每個cell的內容
    max_row = sheet_range.max_row
    max_column = sheet_range.max_column
    for row in range(1, max_row+1):
        for column in range(1, max_column+1):
            cell_obj = sheet_range.cell(row=row,column=column)
            print(cell_obj.value,end=' | ')
        print('\n') #斷行
    
    # 尋找每個column的cell有多少個字，以字數最多的cell為基準放大此column的width
    for row in sheet_range.rows:
        for cell in row:
        	#將所有內容置中
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if cell.value:
                # 將每個column的代號column_letter放入dictionary的key，比較完大小後放入value
                col_widths_dict[cell.column_letter] = max((col_widths_dict.get(cell.column_letter, 0), len(str(cell.value))))

    for col, value in col_widths_dict.items():
        sheet_range.column_dimensions[col].width = int(value) + 5

    # 調整完後存入原檔案中
    wb.save(filename = excel_name + '.xlsx')

def main():
    try:
        stockID = input('請輸入股票代碼：')
    except ValueError:
        print('股票代碼只可以是整數數字')
    stock_data = getdata(geturl(stockID))
    excel_data_name, title = write_in_excel(stock_data)
    fixcolumnwidth(excel_data_name, title)

main()