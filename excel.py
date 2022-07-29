from openpyxl import load_workbook
import time
import os

def add(name, localtime):
    max = 1
    for cell in sheet['A']:
        max+=1
    sheet['A' + str(max)] = name
    sheet['B' + str(max)] = localtime
name = input("Enter your name: ")
localtime = time.asctime(time.localtime(time.time()))
wb = load_workbook('kiosk-reports.xlsx')
sheet = wb['sign-in-reports']
add(name, localtime)
wb.save("kiosk-reports.xlsx")
    



