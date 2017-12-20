import requests
import xlsxwriter
from lxml import html
import os.path
import xlrd
from xlutils.copy import copy


page = requests.get('http://www.bseindia.com/markets/equity/EQReports/bulk_deals.aspx?expandable=3')
tree = html.fromstring(page.content)

#This will create a list of buyers:
dateandTrade = tree.xpath('//td[@class="TTRow"]/text()')
buyers = tree.xpath('//td[@class="TTRow_left"]/text()')
quantity = tree.xpath('//td[@class="TTRow_right"]/text()')

isFileExists = os.path.exists('C://temp//bulk_deal.xls')

row_count =0
if False == isFileExists:
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook('C://temp//bulk_deal.xls')
    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})
    worksheet = workbook.add_worksheet('BulkDeal')
    worksheet.write(0, 1, 'Security Name', bold)
    worksheet.write(0, 0, 'Client Name', bold)
    worksheet.write(0, 2, 'Price', bold)
    worksheet.write(0, 3, 'Quantity', bold)
    worksheet.write(0, 4, 'Deal Type', bold)
    worksheet.write(0, 5, 'Date', bold)
    worksheet.write(0, 6, 'Security Code', bold)
else:
    book = xlrd.open_workbook("C://temp//bulk_deal.xls")
    sheets = book.sheets()
    row_count = sheets[0].nrows
    wb = copy(book)
    worksheet = wb.get_sheet(0)

row = 1 + row_count
col = 0

for index in range(len(buyers)):
    if index%2 == 0:
        worksheet.write(row, col+1,  buyers[index])
    else:
        worksheet.write(row, col, buyers[index])
        row = row + 1

row = 1 + row_count
col = 3
trade = len(dateandTrade)/3

#print dateandTrade
for tr in range(trade):
    worksheet.write(row, col+1, dateandTrade[3*tr+2])
    worksheet.write(row, col + 3, dateandTrade[3 * tr+1])
    worksheet.write(row, col + 2, dateandTrade[3 * tr])
    row = row + 1

row = 1 + row_count
col = 2
for ind in range(len(quantity)):
    if ind % 2 == 0:
        worksheet.write(row, col + 1, quantity[ind])
    else:
        worksheet.write(row, col, quantity[ind])
        row = row + 1
        # print buyers[index]

if False == isFileExists:
    workbook.close()
else:
    wb.save("C://temp//bulk_deal.xls")

