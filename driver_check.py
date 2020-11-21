import os
import openpyxl
from tkinter import messagebox
import tkinter


def check_files():
    if os.path.exists("devcon_amd64.exe"):
        pass
    else:
        tkinter.Tk().withdraw()
        messagebox.showerror("File not found","devcon_amd64.exe not found")
    if os.path.exists("drv_list.txt"):
        pass
    else:
        tkinter.Tk().withdraw()
        messagebox.showerror("File not found","drv_list.txt not found")
    


def get_device_info():
    os.system("devcon_amd64.exe drivernodes * > drivernodes.txt")


def read_id():
    f = open("drv_list.txt")
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
    if os.path.exists("stress.xlsx"):
        wb = openpyxl.load_workbook("stress.xlsx")
        sheet = wb.active
        # load excel and active (ready to write data)
        get_device_version()
        sheet.append(drv_description)
        sheet.append(drv_version)
        wb.save("stress.xlsx")
    else :
        wb = openpyxl.Workbook()
        wb.save("stress.xlsx")
        to_excel()

def clear_temp_data():
    for x in name_list:
        os.remove(x)
    print("done")

def test():
    print(drv_description)
    print(drv_version)



check_files()
read_id()
to_excel()
clear_temp_data()