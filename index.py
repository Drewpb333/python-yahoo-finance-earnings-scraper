from bs4 import BeautifulSoup
import requests
import xlwt
from xlwt import Workbook
import datetime
import calendar

def create_workbook():
    return Workbook()

def create_worksheet():
    return wb.add_sheet("Sheet1")

def create_heading():
    heading = ['Date', 'Symbol', 'Company', 'Call Time', 'EPS Estimate', 'Reported EPS', 'Surpise(%)']
    for i in range(len(heading)):
        sheet1.write(0, i, heading[i])

def save_workbook(file_name):
    wb.save(file_name)

def get_earnings_table(monday_date, friday_date, earnings_date):
    url = 'https://finance.yahoo.com/calendar/earnings?from={}&to={}&day={}'.format(monday_date, friday_date, earnings_date)
    print(url)
    html_get = requests.get(url)
    soup = BeautifulSoup(html_get.text, "html.parser")
    return soup.table