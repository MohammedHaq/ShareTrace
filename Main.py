import requests
from bs4 import BeautifulSoup
import pandas as pd
import win32com.client as win32

df = pd.DataFrame({'A': [10, 20], 'B': [39, 49]})
df.size
tickers = []

NumberOfStocks = int(input("How many stocks would you like to view?"))
for i in range(NumberOfStocks):
    ticker = input(str(i+1) + ". Please input ticker:")
    tickers.append(ticker)


headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.69 Safari/537.36'}

xlApp = win32.Dispatch('Excel.Application')
xlApp.Visible = True
wb = xlApp.Workbooks.Add()

for ticker in tickers:
    url = "https://www.marketwatch.com/investing/stock/{0}/company-profile".format(ticker)
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.content, 'html.parser')

    profile_info = {}

    element_tables = soup.select("div[class='element element--table']")
    for element_table in element_tables:
        valuation_type = element_table.h2.text.strip()
        df = pd.read_html(str(element_table))[0]
        profile_info[valuation_type] = df

    ws = wb.Worksheets.Add()
    ws.name = ticker

    row_spacing = 2

    for table in profile_info.items():
        lastrow = ws.Cells(ws.rows.count, 1).End(-4162).row
        ws.cells(lastrow + row_spacing, 1).value = table[0]
        ws.cells(lastrow + row_spacing, 1).font.bold = True

        ws.Range(
            ws.cells(lastrow + row_spacing + 1, 1),
            ws.cells(lastrow + table[1].shape[0] + row_spacing, table[1].shape[1])
        ).value = table[1].values

    ws.Rows('1:' + str(row_spacing)).delete
    ws.columns(1).columnwidth = 30
    ws.columns(2).columnwidth = 15