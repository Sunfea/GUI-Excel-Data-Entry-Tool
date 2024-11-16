import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import openpyxl
import os

root = tk.Tk()
root.title('GUI Excel')

frame = ttk.Frame(root)
frame.pack()
user_info_labelframe = tk.LabelFrame(frame, text='User Information')
user_info_labelframe.grid(row=0, column=0, padx=20, pady=10)

def data_entered():
    first_name = first_name_entry.get()
    last_name = last_name_entry.get()
    title = title_label_combobox.get()
    age = age_spinbox.get()
    nationality = nationality_label_combobox.get()
    courses = numcourses_spin.get()
    Semester = numsemester_spin.get()
    register = checkbox_var.get()
    print(f'first name: {first_name} | Last name: {last_name} | Title: {title}')
    print(f'Registration Status: {register}')

    filepath = 'D:\\Excel\\data.xlsx'

    if not os.path.exists(filepath):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        heading = ['First Name', 'Last Name', 'Title', 'Registration status']
        sheet.append(heading) 
        workbook.save(filepath)
    workbook = openpyxl.load_workbook(filepath)
    sheet = workbook.active
    sheet.append([first_name, last_name, title, register])
    workbook.save(filepath)

first_name = ttk.Label(user_info_labelframe, text='First Name')
first_name.grid(row=0, column=0)
last_name = ttk.Label(user_info_labelframe, text='Last Name')
last_name.grid(row=0, column=1)

first_name_entry = tk.Entry(user_info_labelframe)
first_name_entry.grid(row=1, column=0)
last_name_entry = tk.Entry(user_info_labelframe)
last_name_entry.grid(row=1, column=1)

title_label = ttk.Label(user_info_labelframe, text='Title')
title_label.grid(row=0, column=2)
title_label_combobox = ttk.Combobox(user_info_labelframe, values=['Mr.', 'Miss.'])
title_label_combobox.grid(row=1, column=2)

age_label = ttk.Label(user_info_labelframe, text='Age')
age_label.grid(row=2, column=0)
age_spinbox = ttk.Spinbox(user_info_labelframe, from_= 18, to=60)
age_spinbox.grid(row=3, column=0)

nationality_label = ttk.Label(user_info_labelframe, text='Nationality')
nationality_label.grid(row=2, column=1)
nationality_label_combobox = ttk.Combobox(user_info_labelframe, values=['India','Nepal', 'Bhutan'])
nationality_label_combobox.grid(row=3, column=1)

for widget in user_info_labelframe.winfo_children():
    widget.grid_configure(padx=10, pady=10)

courses_frame = tk.LabelFrame(frame, text='Course Details')
courses_frame.grid(row=1, column=0, sticky='news', padx=20, pady=10)

checkbox_var = tk.StringVar(value='Not registered')
registered_label = ttk.Label(courses_frame, text='Registration Status')
registered_label.grid(row=0, column=0)
registered_check = ttk.Checkbutton(courses_frame, text='Currently Registered', variable=checkbox_var, onvalue='Registered', offvalue='Not Registred')
registered_check.grid(row=1, column=0)

numcourses_label = ttk.Label(courses_frame, text='Completed Courses')
numcourses_spin = tk.Spinbox(courses_frame, from_=1, to='infinity')
numcourses_label.grid(row=0, column=1)
numcourses_spin.grid(row=1, column=1)

numsemester_label = ttk.Label(courses_frame, text='Semester')
numsemester_spin = tk.Spinbox(courses_frame, from_=1, to='infinity')
numsemester_label.grid(row=0, column=2)
numsemester_spin.grid(row=1, column=2)

for widget in courses_frame.winfo_children():
    widget.grid_configure(padx=10, pady=10)

terms_frame = tk.LabelFrame(frame, text='Terms and co9ndition')
terms_frame.grid(row=2, column=0, sticky='news', padx=20, pady=10)

terms_checkbox = tk.Checkbutton(terms_frame, text='I accept terms and condition')
terms_checkbox.grid(row=0, column=0)

edit_btn = ttk.Button(frame, text='Enter Data', command=data_entered)
edit_btn.grid(row=3, column=0, sticky='news', padx=20, pady=10)

root.mainloop()