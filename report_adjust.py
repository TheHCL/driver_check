import os
import openpyxl
import tkinter
from tkinter import messagebox

def run():
    if os.path.exists("settings.txt"):
        setting = open("settings.txt")
        global report_name
        global tool_name
        tool_name = setting.readline().strip()
        tool_name = tool_name.replace("tool_name=","")
        tool_name = tool_name.replace(" ","")
        report_name = setting.readline().strip()
        report_name = report_name.replace("report_name=","")
        report_name = report_name.replace(" ","")
    
    else:
        tkinter.Tk().withdraw()
        messagebox.showerror("File not found","settings.txt not found")

    if os.path.exists(report_name):
        excel_adjust()  # run method
    else:
        tkinter.Tk().withdraw()
        messagebox.showerror("File not found",report_name+" not found")


def excel_adjust():
    wb = openpyxl.load_workbook(report_name)
    sheet = wb.active
    sheet.insert_cols(1)
    current_sheet = wb.worksheets[0]
    row_count = current_sheet.max_row
    for i in range(0,int(row_count/2)):
        current_sheet.cell(row=2*i+1,column=1).value = i+1
    wb.save(report_name)


run()