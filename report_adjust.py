import os
import openpyxl
import tkinter
from tkinter import messagebox

def read_settings():
    if os.path.exists("settings.txt"):
        setting = open("settings.txt")
        global report_name
        report_name = setting.readline().strip()
        report_name = report_name.replace("report_name=","")
        report_name = report_name.replace(" ","")
    
    else:
        tkinter.Tk().withdraw()
        messagebox.showerror("File not found","settings.txt not found")

    if os.path.exists(report_name):
        excel_adjust()
    else:
        tkinter.Tk().withdraw()
        messagebox.showerror("File not found",report_name+" not found")


def excel_adjust():
    wb = openpyxl.load_workbook(report_name)
    sheet = wb.active
    sheet.insert_cols(1)
    current_sheet = wb.worksheets[0]
    row_count = current_sheet.max_row
    for i in range(0,row_count):
        current_sheet.cell(row=2*i+1,column=1).value = i+1
    wb.save(report_name)


read_settings()