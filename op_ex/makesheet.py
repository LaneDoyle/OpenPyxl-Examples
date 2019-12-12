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

for i in range(6):
    new_row = sheet.max_row+1
    
    first = input("Enter first name of new user: ")
    last = input("Enter last name of new user: ")
    u_name = input("Enter username of new user: ")
    
    print("-----------------------")#Breaks up loop
    
    sheet.cell(row = new_row, column = 1).value = dt.datetime.now()
    sheet.cell(row = new_row, column = 2).vale = first
    sheet.cell(row = new_row, column = 3).vale = last
    sheet.cell(row = new_row, column = 4).vale = u_name
    
workbook.save("users.xlsx")