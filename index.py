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