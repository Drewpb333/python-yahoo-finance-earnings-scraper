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

def write_to_sheet(table, earnings_date, starting_row_xl_index):
    earnings_rows = table.find_all('tr')
    #start at 1 to avoid header from yahoo finance
    for i in range(1, len(earnings_rows)):
        columns = earnings_rows[i].find_all("td")
        sheet1.write(starting_row_xl_index + i, 0, earnings_date)
        #offset by one to manually add date                
        print(earnings_date)
        for j in range(len(columns)):
            sheet1.write(starting_row_xl_index + i, j + 1, columns[j].text)
            # print(columns[j].text)

def create_weekdates_2017_list():
    # weekdates_of_2019 = [['2018-12-31','2019-01-01','2019-01-02', '2019-01-03', '2019-1-04']]
    weekdates_of_2017 = []
    first_monday_date = datetime.date(2017, 1, 8)
    for i in range(51):
        week = []
        for j in range(5):
            day_date = first_monday_date + datetime.timedelta((7 * i) + j)
            week.append(day_date)
        weekdates_of_2017.append(week)
    return weekdates_of_2017   