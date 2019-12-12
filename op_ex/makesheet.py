#!/usr/bin/python3
#Lane Doyle
#12/12/19

'''create and populate an xlsx file called users.xlsx'''
import openpyxl
import datetime as dt
import time

workbook = openpyxl.Workbook()
sheet = workbook.create_sheet("Users")

#Set the headings row
sheet['A1'].value = "Timestamp"
sheet['B1'].value = "First Name"
sheet['C1'].value = "Last Name"
sheet['D1'].value = "Username"

