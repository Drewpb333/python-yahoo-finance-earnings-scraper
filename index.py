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
        for j in range(len(columns)):
            sheet1.write(starting_row_xl_index + i, j + 1, columns[j].text)

# Yahoo Earnings currently only provides earnings for a four year period from 3 yars ago to 1 year in the future
def create_weekdates_list():
    weekdates_of_year = []
    first_monday_date = datetime.date(2017, 1, 1)
    # gets number of weeks from beginning of 3 years ago to one year from today's date
    number_of_weeks = round((datetime.date.today() + datetime.timedelta(365) - jan).days / 7)
    for i in range(number_of_weeks):
        week = []
        for j in range(5):
            day_date = first_monday_date + datetime.timedelta((7 * i) + j)
            week.append(day_date)
        weekdates_of_year.append(week)
    return weekdates_of_year   

wb = create_workbook()
sheet1 = create_worksheet()
create_heading()

weekdates_list = create_weekdates_list()

# for preventing overwrite of existing data
starting_row_xl_index = 0

for week in weekdates_of_year:
    monday_date = week[0]
    friday_date = week[4]
    for i in range(5):
        table = get_earnings_table(monday_date, friday_date, week[i])
        try:
            write_to_sheet(table, week[i], starting_row_xl_index)
        except:
            print('Error Occured for week of {} - {}'.format(monday_date, friday_date))
            break
        starting_row_xl_index += len(table.find_all("tr"))

save_workbook('./yahoo-finance-arnings.xls')