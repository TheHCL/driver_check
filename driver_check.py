import os
import openpyxl
from tkinter import messagebox
import tkinter

def read_settings():
    if os.path.exists("settings.txt"):
        setting = open("settings.txt")
        global report_name
        global hwid_list
        report_name = setting.readline().strip()
        report_name = report_name.replace("report_name=","")
        report_name = report_name.replace(" ","")
        hwid_list = setting.readline().strip()
        hwid_list = hwid_list.replace("HWID_list=","")
        hwid_list = hwid_list.replace(" ","")
        
    else:
        tkinter.Tk().withdraw()
        messagebox.showerror("File not found","settings.txt not found")
    

def check_files():
    if os.path.exists("devcon_amd64.exe"):
        pass
    else:
        tkinter.Tk().withdraw()
        messagebox.showerror("File not found","devcon_amd64.exe not found")
    if os.path.exists(hwid_list):
        pass
    else:
        tkinter.Tk().withdraw()
        messagebox.showerror("File not found",hwid_list+" not found")
    


def read_id():
    f = open(hwid_list)
    count =0
    line = f.readline().strip("\n")
    global name_list
    name_list=[]
    while line:
        print(line)
        name = "output"+str(count)+".txt"
        name_list.append(name)
        os.system("devcon_amd64.exe drivernodes "+line+" > "+name)
        line = f.readline().strip("\n")
        count+=1
    f.close()

def get_device_version():
    global drv_description
    global drv_version
    drv_description =[]
    drv_version=[]
    for x in name_list:
        f = open(x)
        line =f.readline().strip("\n")
        
        while line:
            if "Driver description" in line:
                line = line.replace("Driver description is ","")
                drv_description.append(line)
            if "Driver version" in line:
                line = line.replace("Driver version is ","")
                drv_version.append(line)
            line =f.readline().strip("\n")
        f.close()

def to_excel():   
    if os.path.exists(report_name):
        wb = openpyxl.load_workbook(report_name)
        sheet = wb.active
        # load excel and active (ready to write data)
        get_device_version()
        sheet.append(drv_description)
        sheet.append(drv_version)
        wb.save(report_name)
    else :
        wb = openpyxl.Workbook()
        wb.save(report_name)
        to_excel()

def excel_adjust():
    wb = openpyxl.load_workbook(report_name)
    sheet = wb.active
    sheet.insert_cols(1)
    wb.save(report_name)

def clear_temp_data():
    for x in name_list:
        os.remove(x)
    print("done")




read_settings()
check_files()
read_id()
to_excel()
clear_temp_data()
