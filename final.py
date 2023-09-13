import tkinter
from tkinter import ttk
import os
import openpyxl

window=tkinter.Tk()
window.title("data entry form")
frame=tkinter.Frame(window)
frame.pack()

def enter_data():
   firstname=first_name_entry.get()
   lastname=last_name_entry.get()
   apptype=app_type_entry.get()
   appname=app_name_entry.get()
   apppassword=app_pass_entry.get()
   title=title_combobox.get()

   print('First Name: ',firstname,'Last Name: ',lastname)
   print("Entry Type: ",title, 'App/Website Name:',apptype)
   print("Username: ",appname, 'Password :',apppassword)

   filepath=(r"C:/Users/badou/Desktop/Nethakrek_project/Data.xlsx")
   
   if os.path.exists(filepath):
       workbook=openpyxl.Workbook()
       sheet =workbook.active
       heading =["First Name","Last Name",'App/Website Name:',"Username","Password", "Entry Type"]
       sheet.append(heading)
       #using lists to add excel rows
       workbook.save(filepath)
   workbook=openpyxl.load_workbook(filepath)
   sheet=workbook.active
   sheet.append([firstname,lastname,apptype,appname,apppassword,title])
   workbook.save(filepath)

   

user_info_frame=tkinter.LabelFrame(frame, text= "User information")
user_info_frame.grid(row=0,column=0,padx=20,pady=2,sticky="news")
first_name_label=tkinter.Label(user_info_frame,text="First Name :")
first_name_label.grid(row=0,column=0)
last_name_label=tkinter.Label(user_info_frame,text="Last Name :")
last_name_label.grid(row=0,column=1)

first_name_entry=tkinter.Entry(user_info_frame)
first_name_entry.grid(row=1, column=0)
last_name_entry=tkinter.Entry(user_info_frame)
last_name_entry.grid(row=1, column=1)

for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=20, pady=5)
    
apps_frame=tkinter.LabelFrame(frame,text="Entry")
apps_frame.grid(row=1,column=0,sticky="news",padx=20,pady=20)

app_type_label=tkinter.Label(apps_frame, text="name of app/website :")
app_type_label.grid(row=0,column=0)
app_type_entry=tkinter.Entry(apps_frame)
app_type_entry.grid(row=1,column=0)

app_name_label=tkinter.Label(apps_frame, text="Username :")
app_name_label.grid(row=0,column=1)
app_name_entry=tkinter.Entry(apps_frame)
app_name_entry.grid(row=1,column=1)

app_pass_label=tkinter.Label(apps_frame, text="Password :")
app_pass_label.grid(row=0,column=2)
app_pass_entry=tkinter.Entry(apps_frame)
app_pass_entry.grid(row=1,column=2)

title_label=tkinter.Label(apps_frame,text="Chose Type :")
title_label.grid(row=2,column=0)
title_combobox=ttk.Combobox(apps_frame,values=["","Website","Application"])
title_combobox.grid(row=3,column=0)

for widget in apps_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

#entry button
button = tkinter.Button(frame,text="Enter data",command=enter_data)
button.grid(row=2,column=0,sticky="news",padx=20,pady=20)



window.mainloop()