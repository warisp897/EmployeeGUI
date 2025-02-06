# -*- coding: utf-8 -*-
"""
Created on Wed May 15 11:28:36 2024

@author: Waris Popal
"""

from tkinter import *
from tkinter import ttk
#from tkinter import Toplevel
from PIL import Image, ImageTk
from pdf2image import convert_from_path
from functools import partial
import os
#import sv_ttk
#import re
import math as m
import datetime
from tkcalendar import DateEntry
#import tkinter as tk
#import tktimepicker

import pandas as pd
#import numpy as np


#Creates the main GUI
master = Tk()
master.title("Student Employee Query")
master.minsize(800, 500)
#master.maxsize(1920, 1080)

#image_lab = ttk.Label(master, image = perry_logo)

#Calls the theme
master.call('source', "C:\\Perry Files\\Scripts\\forest-light.tcl")
ttk.Style().theme_use("forest-light")
#ttk.Style().theme_use("clam")
master.configure(background="#ffffff")


#Grab spreadsheets (back-end data)
onedrive_path = "C:\\Users\\" + os.getlogin() + "\\OneDrive - Virginia Tech\\Student Tracking\\Student Sheets\\"
emp_df = pd.read_excel(onedrive_path + "Student Overview.xlsx")
#emp_df = pd.read_excel("C:\\Perry Files\\Test Files\\Student Overview.xlsx")
#memo_df = pd.read_excel("C:\Perry Place\Test Files\Memo Sheet.xlsx")
#write_df = pd.read_excel("C:\Perry Place\Test Files\Write-Up Sheet.xlsx")
call_df = pd.read_excel(onedrive_path + "Call Out Sheet.xlsx")
ncns_df = pd.read_excel(onedrive_path + "NCNS sheet.xlsx")

list_of_employees = emp_df["Name"].tolist()

#Empty list for labels, used to clear labels when needed
mem_write_lab = []

#Generate the dates for the memos and enable buttons to click
def mem_disp(frame_i):
    global name
    
    matching_pdfs = []
    for file_name in os.listdir("C:\\Perry Files\\Student Disc"):
        if file_name.endswith('.pdf'):
            parts = file_name.split('_')
            name_i = (parts[0][0].upper() + parts[0][1:]) + " " + (parts[1][0].upper() + parts[1][1:])
            level = parts[2] == "memo"
            if name_i == name and level:
                matching_pdfs.append(parts)
                
    memo_tab = notebook.add(frame_i, text = 'Memos: ' + str(len(matching_pdfs)))
    
    curr_line = 0
    if (len(matching_pdfs) == 0):
        memo_i = ttk.Label(frame_i, text = "No Memos",
              font = ("Arial", 16)).place(x = 0, y = 30 * curr_line)
    else:
        for i in range(0, len(matching_pdfs)):
            date_str = matching_pdfs[i][3].replace(".pdf", "")
            #format the date from 05142024 to 05/14/2024
            text_date = date_str[0:2] + "/" + date_str[2:4] + "/" + date_str[4:]
            
            memo_i = ttk.Label(frame_i, text = text_date,
                                font = ("Arial", 16, "underline"), 
                                cursor = "hand2",
                                foreground = "#0000EE")
            memo_i.place(x = 0, y = 30 * curr_line)
            curr_line += 1
            
            #Setting up file name
            file_name =  "_".join(matching_pdfs[i])
            
            #mini function to call disp_file with the filename
            def call_disp_file(event, fn = file_name):
                display_file(fn)
            
            #binds the mini function to the label
            memo_i.bind("<Button-1>", call_disp_file)
    
    #add = ttk.Button(frame_i, text = "Add", command = new_memo)
    #rem = ttk.Button(frame_i, text = "Remove", command = rem_memo)
    #add.place(x = tab_w / 2.05, y = tab_h / 1.35)
    #rem.place(x = tab_w / 2.05, y = tab_h / 1.2)

#Generate dates for write ups and enable buttons to click
def write_disp(frame_i):
    global name 
    
    matching_pdfs = []
    for file_name in os.listdir("C:\\Perry Files\\Student Disc"):
        if file_name.endswith('.pdf'):
            parts = file_name.split('_')
            name_i = (parts[0][0].upper() + parts[0][1:]) + " " + (parts[1][0].upper() + parts[1][1:])
            level = parts[2] == "writeup"
            if name_i == name and level:
                matching_pdfs.append(parts)
                
    write_tab = notebook.add(frame_i, text = 'Write Ups: ' + str(len(matching_pdfs)))
    
    curr_line = 0
    if (len(matching_pdfs) == 0):
        write_i = ttk.Label(frame_i, text = "No Write Ups",
              font = ("Arial", 16)).place(x = 0, y = 30 * curr_line)
    else:
        for i in range(0, len(matching_pdfs)):
            date_str = matching_pdfs[i][3].replace(".pdf", "")
            #format the date from 05142024 to 05/14/2024
            text_date = date_str[0:2] + "/" + date_str[2:4] + "/" + date_str[4:]
            
            write_i = ttk.Label(frame_i, text = text_date,
                                font = ("Arial", 16, "underline"), 
                                cursor = "hand2",
                                foreground = "#0000EE")
            write_i.place(x = 0, y = 30 * curr_line)
            curr_line += 1
            
            #Setting up file name
            file_name =  "_".join(matching_pdfs[i])
            #print(matching_pdfs)
            #print(matching_pdfs[i])

            def call_disp_file(event, fn = file_name):
                display_file(fn)
            
            write_i.bind("<Button-1>", call_disp_file)
        
    #add = ttk.Button(frame_i, text = "Add", command = new_write)
    #rem = ttk.Button(frame_i, text = "Remove", command = rem_write)
    #add.place(x = tab_w / 2.05, y = tab_h / 1.35)
    #rem.place(x = tab_w / 2.05, y = tab_h / 1.2)

#Display the trainings that need to be completed
def train_disp(frame_i):
    global name
    global train_list
    
    if "train_list" in locals():
        train_list.clear()
    #Trainings
    train_ind = emp_df.index[emp_df["Name"] == name][0]
    train_list = emp_df.loc[train_ind, "Trainings Completed"]
    print("this is the list generated:" + str(train_list) + "_")
    print(train_list)
    print(type(train_list))
    
    if (type(train_list) != str or train_list == ""):
        train_list = []
        train_len = 0
        ttk.Label(frame_i, text = "No Trainings Completed", 
                              font = ("Arial", 14)).place(x = 0, y = 0)
    else:
        train_list = train_list.split(", ")
        
        #Fixes the issue of an empty string populating a student
        #training list when it's empty
        if "" in train_list:
            train_list.remove("")
            
        train_len = len(train_list)
        ttk.Label(frame_i, text = "The following trainings are complete:",
                  font = ("Arial", 16)).place(x = 0, y = 0)
        
        training_lab = []
        for i in range(0, train_len):
            training_lab.append(ttk.Label(frame_i, text = train_list[i],
                      font = ("Arial", 14)))
            training_lab[i].place(x = 0, y = 35 * (0.75 * (i + 1)) + 10)
        
    train_tab = notebook.add(frame_i, text = "Trainings: " + str(train_len))
    
    add = ttk.Button(frame_i, text = "Add", command = new_train)
    rem = ttk.Button(frame_i, text = "Remove", command = rem_train)
    add.place(x = tab_w / 2.05, y = tab_h / 1.35)
    rem.place(x = tab_w / 2.05, y = tab_h / 1.2)
    
            
def call_disp(frame_i):
    global name
    
    #copy of call_df that only has the callouts from the submitted name
    call_name_df = call_df[call_df["Name"] == name].copy()
    call_name_df.reset_index(drop = True, inplace = True)
    
    #Get callout amount
    call_val = len(call_name_df)

    #generate the tab with the title Callouts: (num of callouts)
    call_tab = notebook.add(frame_i, text = 'Call Outs: ' + str(call_val))
    
    #No callouts from this person
    if (call_val == 0):
        ttk.Label(frame_i, text = "No Callouts", 
                              font = ("Arial", 14)).place(x = 0, y = 0)
        
    else:
        curr_line = 0
        for i in range(0, call_val):
            print("row", curr_line)
        #Display the callout date, time, reason, and then comments           
            call_i = ttk.Label(frame_i, text = str(call_name_df["Date"][i])[5:10].replace("-","/")
                                   + "/" + str(call_name_df["Date"][i])[0:4]
                                   + " | "
                                   + str(call_name_df["Time"][i])[0:5]
                                   + " | " 
                                   + call_name_df["Reason"][i], 
                                   font = ("Arial", 16))
            call_i.place(x = 0, y = 30 * curr_line)
            call_c = call_name_df["Comments"][i]
            curr_line += 1
                
            #End here if theres no comments written in the callout
            if type(call_c) == type(1.1):
                curr_line += 1
                continue
                
            #In case the comments part is too long, split the string
            #over multiple lines
            call_c = call_c.split()
                
            #Index counters
            prev_ind = 0
            curr_ind = 0
            len_counter = 0
            
            #the size variable affects the number of characters before the
            #comments moves on to the next line
            size = master.winfo_width() / 20
            
            words_checked = 1
            for word in call_c:
                len_counter += len(word)
                curr_ind += 1
                words_checked += 1
                #newline if number of chars> size or its the last word
                if len_counter > size or words_checked == len(call_c) + 1:
                    text_c = " ".join(call_c[prev_ind:curr_ind])
                    prev_ind = curr_ind
                    ttk.Label(frame_i, text = text_c, 
                              font = ("Arial", 14)).place(
                                  x = 0, y = 30 * curr_line)
                    curr_line += 1
                    len_counter = 0
                    words_checked = 0
                    
            curr_line += 1
    
    #Generate add/remove buttons on the bottom right
    add = ttk.Button(frame_i, text = "Add", command = new_call)
    rem = ttk.Button(frame_i, text = "Remove", command = rem_call)
    add.place(x = tab_w / 2.05, y = tab_h / 1.35)
    rem.place(x = tab_w / 2.05, y = tab_h / 1.2)
        
          
def ncns_disp(frame_i):
    global name
    global ncns_name_df
    
    #get number of NCNS
    ncns_name_df = ncns_df[ncns_df["Name"] == name].copy()
    ncns_name_df.reset_index(drop = True, inplace = True)
    
    ncns_val = len(ncns_name_df)
    ncns_tab = notebook.add(frame_i, text = 'NCNS: ' + str(ncns_val))
    
    #if none, dislplay no ncns
    if ncns_val == 0:
        ttk.Label(frame_i, text = "No NCNS", 
                  font = ("Arial", 16)).place(x = 0, y = 0)
        
    #if there are ncns, display the date. If comments, display that as well?
    curr_line = 0
    for i in range(0, ncns_val):
        
        #Pull the date
        ncns_i = str(ncns_name_df["Date"][i])[0:10].replace("-", "/")
        
        #2024/05/03 -> 05/03/2024
        ncns_i = ncns_i[5:8] + ncns_i[8:10] + "/" + ncns_i[0:4]
        
        #05/03/2024 -> 05/03/2024 16:01
        ncns_i = ncns_i + " " + str(ncns_name_df["Date"][i])[10:16]
        
        #create label
        ttk.Label(frame_i, text = ncns_i, font = 
                  ("Arial", 16)).place(x = 0, y = 30 * curr_line)
        
        curr_line += 2

    add = ttk.Button(frame_i, text = "Add", command = new_ncns)
    rem = ttk.Button(frame_i, text = "Remove", command = rem_ncns)
    add.place(x = tab_w / 2.05, y = tab_h / 1.35)
    rem.place(x = tab_w / 2.05, y = tab_h / 1.2)
        
        
        
#Function to search for the name after submit is clicked
#If found, generated the text.

global_labels = []

def search_name():
    global name
    
    #Delete memo/write labels
    for lab in mem_write_lab:
        lab.destroy()
    mem_write_lab.clear()
    
    #Check if both boxes filled
    first = e1.get().lower().replace(" ", "")
    last = e2.get().lower().replace(" ", "")
    if (len(first) == 0 or len(last) == 0):
        error_label.config(text = "Error: Fill in Both Boxes",  
                           foreground='red')
        #raise Exception("Fill in both boxes") 
        return 0
       
    #create name
    first = (first[0].upper() + first[1:])
    last = (last[0].upper() + last[1:])
    name = first + " " + last
    
    #destroy old labels
    for label in global_labels:
        label.destroy()
    global_labels.clear()
    
    #check if name exists
    if name not in list_of_employees:
        print(name, "not found")
        error_label.config(text = "Error: " + name + " not found", 
                           foreground='red')
    else:
        error_label.config(text="")
        print("found", name)
        
        create_tabs(name)

def add_rem_button(frame_i, tab_h, tab_w):
    global add_rem_buttons
    
    add = ttk.Button(frame_i, text = "Add", command = new_train)
    rem = ttk.Button(frame_i, text = "Remove", command = rem_train)
    add.place(x = tab_w / 2.05, y = tab_h / 1.35)
    rem.place(x = tab_w / 2.05, y = tab_h / 1.2)
    
    add_rem_buttons.append(add)
    add_rem_buttons.append(rem)
    
    #add_rem goes on infinitely!! AAHHH
        
def create_tabs(name):
    global add_rem_buttons 
    global global_labels
    global tab_h
    global tab_w
    
    #Will destroy the old frames so they dont take up space
    for label in global_labels:
        label.destroy()
    
    #clears the list so they are no longer referenced
    global_labels.clear()
    
    #for button in add_rem_buttons:
        #button.destroy()
        
    #Create frames:
    tab_h = master.winfo_height()
    tab_w = master.winfo_width()
    
    frame1 = ttk.Frame(notebook, width = tab_w, height = tab_h)
    frame2 = ttk.Frame(notebook, width = tab_w, height = tab_h)
    frame3 = ttk.Frame(notebook, width = tab_w, height = tab_h)
    frame4 = ttk.Frame(notebook, width = tab_w, height = tab_h)
    frame5 = ttk.Frame(notebook, width = tab_w, height = tab_h)
    
    train_disp(frame1)
    global_labels.append(frame1)
    
    call_disp(frame2)
    global_labels.append(frame2)
                
    ncns_disp(frame3)
    global_labels.append(frame3)
    
    mem_disp(frame4)
    global_labels.append(frame4)
    
    write_disp(frame5)
    global_labels.append(frame5)
    
    notebook.place(x = 300, y = 0)
    
    add_rem_buttons = []


#Popup to add a training
def new_train():
    global train_add_popup
    global new_t
    
    train_add_popup = Toplevel(master)
    train_add_popup.title("New Training Completed")
    train_add_popup.geometry("400x300")
    
    train_title = ttk.Label(train_add_popup, text = "Enter Complete Training:",
          font = ("Arial", 16, "underline"))
    train_title.place(x = 30, y = 30)
    
    date_lab = ttk.Label(train_add_popup, text = "Date Completed:",
                           font = ("Arial", 16))
    date_lab.place(x = 30, y = 100)
    
    train_date = DateEntry(train_add_popup)
    train_date.place(x = 220, y = 95)
    
    train_text = ttk.Label(train_add_popup, text = "Training Completed:",
                                font = ("Arial", 16))
    train_text.place(x = 30, y = 150)
    
    new_t = ttk.Entry(train_add_popup, width = 20)
    new_t.place(x = 220, y = 145)
    
    add_submit_button = ttk.Button(train_add_popup, text = "Submit", width = 31,
               command = train_add)
    add_submit_button.place(x = 74, y = 200)
    
    #makes the submit button green if text is written in the entry
    def button_pressed(event):
        if len(new_t.get()) > 0:
            add_submit_button.configure(style="Accent.TButton")
        else:
            add_submit_button.configure(style="TButton")
    
    #submits if enter is pressed
    def hitEnter(event):
        train_add()

    train_add_popup.bind("<Return>", hitEnter)
    
    train_add_popup.bind("<KeyPress>", button_pressed)

#Popup to remove from the list of trainings
def rem_train():
    global train_remove_popup
    global button_list
    
    train_remove_popup = Toplevel(master)
    train_remove_popup.title("Remove Trainings")
    train_remove_popup.geometry("400x400")
    Label(train_remove_popup, text = "Select Trainings to Remove:",
          font = ("Arial", 16)).place(x = 70, y = 50)
    button_list = []
    train_buttons()
   
#Creates the checkboxes to remove the trainings
def train_buttons():
    global popup
    global popup_submit
    global name 
    global button_list
    global train_list
    global variable_list
    #global train_title
    
    for button in button_list:
        #print("button text:", button.cget("text"))
        button.destroy()
    
    #print("\n the amount of buttons is", len(button_list))
    button_list.clear()
    #print("\n now cleared is", len(button_list))
    
    train_ind = emp_df.index[emp_df["Name"] == name][0]
    train_str = emp_df.loc[train_ind, "Trainings Completed"]
    
    #train_str returns a float value is it is empty
    if type(train_str) == float:
        for widget in train_remove_popup.winfo_children():
            widget.destroy()
        ttk.Label(train_remove_popup, text = "No Trainings to Remove", 
                  font = ("Arial", 16)).place(x = 100, y = 100)
        return 0
    
    train_list = train_str.split(", ")
    print("\n training list: ", train_list, "\n")
    if "" in train_list:
        train_list.remove("")
        
    variable_list = []
    
    #all trainings were removed
    
    if len(train_list) == 0 or train_list[0] == "":
        
        #should destroy every object
        for widget in train_remove_popup.winfo_children():
            widget.destroy()
            
        ttk.Label(train_remove_popup, text = "No Trainings to Remove", 
                  font = ("Arial", 16)).place(x = 100, y = 50)
        return 0
    
    #some trainings remain
    for i in range(0, len(train_list)):
        intvar_i = IntVar(value = 0)
        print("button text:", train_list[i])
        check_i = ttk.Checkbutton(train_remove_popup, text = train_list[i],
                                  variable = intvar_i,
                                  onvalue = 1, 
                                  offvalue = 0, 
                                  width = 20,
                                  command = submit_color)
        #check_i.state(None)
        check_i.place(x = 50, y = 30 * i + 100)
        
        variable_list.append(intvar_i)
        button_list.append(check_i)
    
    popup_submit = ttk.Button(train_remove_popup, text = "Submit",
               width = 15, command = lambda: refresh_train(button_list))
    
    popup_submit.place(x = 50, y = 30 * len(train_list) + 120)

#Changes the submit button color when a box is checked
def submit_color():
    global variable_list

    for var in variable_list:
        if var.get() == 1:
            popup_submit.configure(style="Accent.TButton")
            return 0

    popup_submit.configure(style="TButton")

#Removes or adds a training to the list, refreshes the main tab
def refresh_train(submission):
    global name
    global popup_submit
    global global_labels
    
    if type(submission) == list:
        if len(button_list) == 0:
            return 0
        
        for button in button_list:
            print(button.state())
            
            #button.state returns a tuple of the state it is in
            popup_submit.destroy()
            print(train_list, "\n")
            
            if (button.state() == ("selected",)):
                print("Identified selected button")
                print("_" + button.cget("text") + "_" "\n")
                
                train_list.remove(button.cget("text"))

    #print("replacing with this", ','.join(train_list))
    
    train_ind = emp_df.index[emp_df["Name"] == name][0]
    emp_df.loc[train_ind, "Trainings Completed"] = ', '.join(train_list)
    
    train_buttons()
    for tab in global_labels:
        tab.destroy()
        
    create_tabs(name)
      
#The function that is run when submit is clicked in the add tab for training
def train_add():
    global new_t
    global train_list
    global name
    global train_add_popup
    
    #add it to the list
    train_list.append(new_t.get())
    
    #add the new list to the df
    train_ind = emp_df.index[emp_df["Name"] == name][0]
    emp_df.loc[train_ind, "Trainings Completed"] = ', '.join(train_list)
    
    #create the frames again
    create_tabs(name)
    train_add_popup.destroy()

def new_call():
    global call_add_popup
    global new_hour
    global new_min
    global new_date
    global new_time
    global am_pm_lab
    global new_reason
    global new_comments
    global error_text
    global intvar_hour
    #global late_bool
    
    call_add_popup = Toplevel(master)
    call_add_popup.title("New Callout")
    call_add_popup.geometry("550x400")
    
    #global train_title
    
    call_title = Label(call_add_popup, text = "Enter Call-Out:",
          font = ("Arial", 16, "underline"))
    
    call_title.place(x = 70, y = 50)
    
    #DATE
    new_date = DateEntry(call_add_popup, width = 10, borderwidth = 1,
                         background = "white", foreground = "green")
    new_date.place(x = 190, y = 110)
    
    date_lab = Label(call_add_popup, text = "Date:", font = ("Arial", 16))
    date_lab.place(x = 70, y = 110)

    #TIME
    date_lab = Label(call_add_popup, text = "Time:", font = ("Arial", 16))
    date_lab.place(x = 70, y = 150)
    
    #new_time = tktimepicker.SpinTimePickerModern(call_add_popup)
    #new_time.place(x = 70, y = 20)
    
    #format_lab = Label(call_add_popup, text = "HR : MN", font = ("Arial", 14))
    #format_lab.place(x = 400, y = 200)
    colon_lab = Label(call_add_popup, text = ":", font = ("Arial", 20))
    colon_lab.place(x = 270, y = 145)
    
    #hour
    
    #This function sets the min/max value that can be entered into the hour
    #box (1 - 12)
    def hour_func(num):
        if num.isdigit():
            if 1 <= int(num) <= 12:
                return TRUE
        return False
    
    #This tuple makes it so that the function runs at every new entry
    hour_check = (master.register(hour_func), "%P")
    
    #Constructor for the hour spinbox entry
    new_hour = ttk.Spinbox(call_add_popup, values = [i for i in range(1, 13)], 
                            width = 2, validate = "key", 
                            validatecommand = hour_check)
    
    #Set the initial text to --
    new_hour.set("--")
    new_hour.place(x = 190, y = 150)
    
    #minute
    
    #Function sets the min/max value that can be entered in the min box (1-60)
    def min_func(num):
        if num.isdigit():
            if 0 <= int(num) <= 59:
                return TRUE
        return False
    
    #This tuple makes it so that the function runs at every new entry
    min_check = (master.register(min_func), "%P")
    
    #Constructor for the min spinbox entry
    new_min = ttk.Spinbox(call_add_popup, values = [i for i in range(0, 60)], 
                          width = 2, validate = "key", 
                          validatecommand = min_check)
    
    new_min.set("--")
    new_min.place(x = 290, y = 150)
    
    #AM/PM
    am_pm_lab = ttk.Combobox(call_add_popup, values = ["AM", "PM"], 
                             width = 3)
    am_pm_lab.set("--")
    am_pm_lab.place(x = 370, y = 150)
    
    #reason text
    reason_lab = ttk.Label(call_add_popup, text = "Reason:", font = ("Arial", 16))
    reason_lab.place(x = 70, y = 190)
    
    #reason entry
    new_reason = ttk.Entry(call_add_popup, width = 40)
    new_reason.place(x = 190, y = 190)
    
    #comments text
    comments_lab = ttk.Label(call_add_popup, text = "Comments:", font = ("Arial", 16))
    comments_lab.place(x = 70, y = 230)
    
    #comments entry
    new_comments = ttk.Entry(call_add_popup, width = 40)
    new_comments.place(x = 190, y = 230)
    
    #within 1 hour label
    #hour_label = ttk.Label(call_add_popup, text = "Submitted Late?", font = 
                           #("Arial", 16))
    #hour_label.place(x = 70, y = 270)
    
    #within 1 hour checkbox
    #intvar_hour has to be made global because otherwise
    #garbage collection deletes it and sets the state of the checkbox to ()
    #late_bool = False
    #intvar_hour = IntVar(value = 0)
    #hour_box = ttk.Checkbutton(call_add_popup, variable = intvar_hour, 
                               #onvalue = 1, offvalue = 0)
    
    #Find a way to set the checkbox to be unchecked by default
    #hour_box.set("0")
    #hour_box.place(x = 340, y = 270)
    
    #Error label (yells at user if submit is pressed with blank boxes)
    error_text = ttk.Label(call_add_popup, text = "", foreground = "red", 
                           width = 40, font = ("Arial", 14))
    error_text.place(x = 160, y = 345)
    
    def hitEnter(event):
        call_add()

    call_add_popup.bind("<Return>", hitEnter)
    
    submit_call = ttk.Button(call_add_popup, text = "Submit", command = call_add,
                        width = 39)
    submit_call.place(x = 140, y = 290)

#This is the function that is run when submit is pressed in callout -> add
def call_add():
    #Check if a time was entered
    if new_hour.get() == "--" or new_min.get() == "--":
        error_text.config(text = "Error: Enter Time Submitted")
        return 0
    
    #Check if a reason was entered
    if new_reason.get() == "":
        error_text.config(text = "Error: Enter Reason")
        return 0
    
    error_text.config(text = "")
    new_call_row = [name, new_date.get_date(), 
                    datetime.time(int(new_hour.get()), int(new_min.get())), 
                    new_reason.get(), new_comments.get()]
    
    call_df.loc[len(call_df)] = new_call_row
    
    #Recreate tabs with new callout
    create_tabs(name)
    
    #Set the selected tab to be '1' (the second tab)
    notebook.select(1)
    call_add_popup.destroy()
    
def rem_call():
    global call_rem_popup
    global call_buttons
    global popup_submit
    global name 
    global button_list
    global train_list
    global variable_list
    
    call_rem_popup = Toplevel(master)
    call_rem_popup.title("Remove Callouts")
    call_rem_popup.geometry("400x400")
    Label(call_rem_popup, text = "Select Callouts to Remove:",
          font = ("Arial", 16)).place(x = 70, y = 100)
    call_buttons = []
    call_boxes()
    
def call_boxes():
    global call_rem_popup
    global call_buttons
    global popup_submit
    global name 
    global button_list
    global train_list
    global variable_list
    global call_name_df
    
    for button in call_buttons:
        print("button text:", button.cget("text"))
        button.destroy()
    
    #print("\n the amount of buttons is", len(button_list))
    call_buttons.clear()
    #print("\n now cleared is", len(button_list))

    variable_list = []
    
    call_name_df = call_df[call_df["Name"] == name].copy()
    call_name_df.reset_index(drop = True, inplace = True)
    
    call_val = len(call_name_df)
    
    #all callouts were removed
    if call_val == 0:
        
        #should destroy every object
        for widget in call_rem_popup.winfo_children():
            widget.destroy()
            
        ttk.Label(call_rem_popup, text = "No Callouts to Remove", 
                  font = ("Arial", 16)).place(x = 100, y = 100)
        return 0
    
    #some callouts remain
    for i in range(0, len(call_name_df)):
        if call_name_df["Name"][i] == name:
            intvar_i = IntVar(value = 0)
            date_text = (str(call_name_df["Date"][i])[5:10] + "/" + str(call_name_df["Date"][i])[0:4]).replace("-", "/")
            
            check_i = ttk.Checkbutton(call_rem_popup, text = date_text +
                                      " " + str(call_name_df["Reason"][i]),
                                      variable = intvar_i,
                                      onvalue = 1, 
                                      offvalue = 0, 
                                      width = 70,
                                      command = submit_color)
            #check_i.state(None)
            check_i.place(x = 50, y = 30 * i + 150)
            
            variable_list.append(intvar_i)
            call_buttons.append(check_i)
            print("appended", check_i.cget("text"))
    
    popup_submit = ttk.Button(call_rem_popup, text = "Submit",
               width = 15, command = lambda: refresh_calls(call_buttons))
    
    popup_submit.place(x = 50, y = 30 * len(call_name_df) + 170)


def refresh_calls(submission):
    global name
    global popup_submit
    global global_labels
    global call_df
    
    #
    if type(submission) == list:
        if len(submission) == 0:
            print("ended :(")
            return 0
        
        for button in submission:
            print(button.state())
            
            #button.state returns a tuple of the state it is in
            #popup_submit.destroy()
            
            if (button.state() == ("selected",)):
                print(button.cget("text")[0:10])
                remove_date = button.cget("text")[0:10].replace("/", "-")
                remove_date = remove_date[6:10] + "-" + remove_date[0:5]
                #remove_time = button.cget("text")[10:16]
                print("\n", button.cget("text"), "\n")
                
                #search the entire original df for the callout, remove it,
                #fix the ordering
                i = 0
                while i <= len(call_df):
                    if call_df["Name"][i] == name and str(call_df["Date"][i])[0:10] == remove_date:
                        popup_submit.destroy()
                        
                        #Remove the row from the df
                        call_df.drop(index = i, inplace = True)
                        
                        #fix the indexing
                        call_df.reset_index(drop = True, inplace = True)
                        #done with the while loop
                        break
                    i += 1
    
    #Reset tabs
    for tab in global_labels:
        tab.destroy()
    
    #create tabs
    create_tabs(name)
    
    #set selected tab as callout
    notebook.select(1)
    
    #regenerate checkboxes
    call_boxes()

def new_ncns():
    global ncns_add_popup
    global ncns_new_hour
    global ncns_new_min
    global ncns_am_pm_lab
    global ncns_new_date
    global submit_ncns_add
    global ncns_error_text
    
    ncns_add_popup = Toplevel(master)
    ncns_add_popup.title("Add NCNS")
    ncns_add_popup.geometry("500x350")

    #Title
    title = ttk.Label(ncns_add_popup, text = "Add No-Call No-Show:",
          font = ("Arial", 16, "underline"))
    title.place(x = 50, y = 50)
    
    #Date label
    date_label = ttk.Label(ncns_add_popup, text = "Date of Missed Shift:", 
                           font = ("Arial", 16))
    date_label.place(x = 50, y = 130)
    
    #Date entry
    ncns_new_date = DateEntry(ncns_add_popup, width = 10, borderwidth = 1,
                         background = "white", foreground = "green")
    ncns_new_date.place(x = 255, y = 127)
    
    #Time label
    time_label = ttk.Label(ncns_add_popup, text = "Time of Missed Shift:", 
                          font = ("Arial", 16))
    
    time_label.place(x = 50, y = 170)
    
    #submit button
    submit_ncns_add = ttk.Button(ncns_add_popup, text = "Submit", width = 34,
                                 command = add_ncns)
    submit_ncns_add.place(x = 120, y = 225)
    
    #Time entry
    def hour_func(num):
        if num.isdigit():
            if 1 <= int(num) <= 12:
                return TRUE
        return False
    
    #This tuple makes it so that the function runs at every new entry
    ncns_hour_check = (master.register(hour_func), "%P")
    
    #Constructor for the hour spinbox entry
    ncns_new_hour = ttk.Spinbox(ncns_add_popup, values = [i for i in range(1, 13)], 
                            width = 2, validate = "key", 
                            validatecommand = ncns_hour_check)
    
    #Set the initial text to --
    ncns_new_hour.set("--")
    ncns_new_hour.place(x = 200, y = 170)
    
    #minute
    
    #Function sets the min/max value that can be entered in the min box (1-60)
    def min_func(num):
        if num.isdigit():
            if 0 <= int(num) <= 59:
                return TRUE
        return False
    
    #This tuple makes it so that the function runs at every new entry
    ncns_min_check = (master.register(min_func), "%P")
    
    #Constructor for the min spinbox entry
    ncns_new_min = ttk.Spinbox(ncns_add_popup, values = [i for i in range(0, 60)], width = 2, 
                          validate = "key", validatecommand = ncns_min_check)
    
    ncns_new_min.set("--")
    ncns_new_min.place(x = 290, y = 170)
    
    #AM/PM
    ncns_am_pm_lab = ttk.Combobox(ncns_add_popup, values = ["AM", "PM"], 
                             width = 3)
    ncns_am_pm_lab.set("--")
    ncns_am_pm_lab.place(x = 380, y = 170)
    
    #Error text block
    ncns_error_text = ttk.Label(ncns_add_popup, text = "", foreground = "red",
                           width = 40, font = ("Arial", 16))
    ncns_error_text.place(x = 120, y = 270)
    
def add_ncns():
    global ncns_new_date
    
    if ncns_new_hour.get() == "--" or ncns_new_min.get(
            ) == "--" or ncns_am_pm_lab.get() == "--":
        ncns_error_text.config(text = "Error: Enter Time Submitted")
        return 0
    
    ncns_day = ncns_new_date.get_date()
    ncns_complete = datetime.datetime.combine(ncns_day, 
                                              datetime.datetime.min.time()
                                              ).replace(hour = 
                                                        int(ncns_new_hour.get()), 
                                                        minute = 
                                                        int(ncns_new_min.get()))
    
    ncns_error_text.config(text = "")
    new_ncns_row = [name, ncns_complete] 
    
    ncns_df.loc[len(ncns_df)] = new_ncns_row
    
    #Recreate tabs with new callout
    create_tabs(name)
    
    #Set the selected tab to be '1' (the second tab)
    notebook.select(2)
    ncns_add_popup.destroy()


def rem_ncns():
    global ncns_rem_popup
    global ncns_buttons
    global popup_submit
    global name 
    global ncns_button_list
    global variable_list
    
    ncns_rem_popup = Toplevel(master)
    ncns_rem_popup.title("Remove NCNS")
    ncns_rem_popup.geometry("400x400")
    Label(ncns_rem_popup, text = "Select NCNS to Remove:",
          font = ("Arial", 16)).place(x = 70, y = 100)
    ncns_buttons = []
    ncns_boxes()

def ncns_boxes():
    global ncns_rem_popup
    global ncns_buttons
    global popup_submit
    global name 
    global button_list
    global variable_list
    
    for button in ncns_buttons:
        print("button text:", button.cget("text"))
        button.destroy()
    
    variable_list = []
    #print("\n the amount of buttons is", len(button_list))
    ncns_buttons.clear()
    #print("\n now cleared is", len(button_list))
    
    #ncns_name_df = ncns_df["Name"] == name
    ncns_val = len(ncns_name_df)
    
    #all ncnsouts were removed
    if ncns_val == 0:
        
        #should destroy every object
        for widget in ncns_rem_popup.winfo_children():
            widget.destroy()
            
        ttk.Label(ncns_rem_popup, text = "No NCNS to Remove", 
                  font = ("Arial", 16)).place(x = 100, y = 100)
        return 0
    
    #some ncnsouts remain
    #print(ncns_name_df)
    for i in range(0, len(ncns_name_df)):
        intvar_i = IntVar(value = 0)
        date_text = ((str(ncns_name_df["Date"][i])[5:10] 
                     + "/" 
                     + str(ncns_name_df["Date"][i])[0:4]).replace("-", "/") 
                     + str(ncns_name_df["Date"][i])[10:16])
        
        
        check_i = ttk.Checkbutton(ncns_rem_popup, text = date_text,
                                  variable = intvar_i,
                                  onvalue = 1, 
                                  offvalue = 0, 
                                  width = 70,
                                  command = submit_color)
        #check_i.state(None)
        check_i.place(x = 50, y = 30 * i + 150)
        
        variable_list.append(intvar_i)
        ncns_buttons.append(check_i)
        #print("appended", check_i.cget("text"))
    
    popup_submit = ttk.Button(ncns_rem_popup, text = "Submit",
               width = 15, command = lambda: ncns_buttons(ncns_buttons))
    
    popup_submit.place(x = 50, y = 30 * len(ncns_name_df) + 170)

def ncns_buttons(submission):
    global name
    global ncns_popup
    global global_labels
    global call_df
    
    print("\n", type(ncns_buttons), "\n")
    
    if type(submission) == list:
        if len(submission) == 0:
            print("ended :(")
            return 0
        
        for button in submission:
            print(button.state())
            
            #button.state returns a tuple of the state it is in
            #popup_submit.destroy()
            
            if (button.state() == ("selected",)):
                print(button.cget("text")[0:10])
                remove_date = button.cget("text")[0:10]
                #remove_time = button.cget("text")[10:16]
                print("\n", button.cget("text"), "\n")
                
                i = 0
                while i <= len(call_df):
                    print("run", i)
                    if ncns_df["Name"][i] == name and str(ncns_df["Date"][i])[0:10] == remove_date:
                        popup_submit.destroy()
                        
                        #Remove the row from the df
                        ncns_df.drop(index = i, inplace = True)
                        #fix the indexing
                        ncns_df.reset_index(drop = True, inplace = True)
                        #done with the while loop
                        break
                    i += 1
    
    #Reset tabs
    for tab in global_labels:
        tab.destroy()
    
    #create tabs
    create_tabs(name)
    
    #set selected tab as callout
    notebook.select(2)
    
    #regenerate checkboxes
    ncns_boxes()


def new_memo():
    #REACH
    pass

def rem_memo():
    #REACH
    pass

def new_write():
    #REACH
    pass

def rem_write():
    #REACH
    pass

#This function is run when the program is closed. It saves all changes
#made to the original excel spreadsheet.
def exit_save():
    try:
        emp_df.to_excel(onedrive_path + "Student Overview.xlsx", 
                        index = False)
        
        ncns_df.to_excel(onedrive_path + "NCNS sheet.xlsx",
                         index = False)
        
        call_df.to_excel(onedrive_path + "Call Out Sheet.xlsx",
                         index = False)
        
        #Memo and Write up may not need to be saved since add/remove is REACH
        #memo_df.to_excel("C:\Perry Place\Test Files\Memo Sheet.xlsx", 
                         #index = False)
        
        #write_df.to_excel("C:\Perry Place\Test Files\Write-Up Sheet.xlsx",
                          #index = False)
        
    finally:
        #shuts the GUI down
        master.destroy()
        
#Function to display files, should work for all files
#Writeups and memos should go in the same folder for this to work

def display_file(file_name):
    newfile = Toplevel()
    newfile.title("Employee File")
    #print('Function called with file name:', file_name)
    
    file_path = os.path.join("C:\\Perry Files\\Student Disc", file_name)
    print('File path:', file_path)
    
    #throw error if file not found
    if not os.path.exists(file_path):
        print('File does not exist:', file_path)
        return
    
    #convert pdf -> png
    try:
        new_image = convert_from_path(file_path)
    except Exception as e:
        print('Error converting PDF:', e)
        return

    
    #get first image (in case it goes over 1 page)
    new_image = new_image[0]
    
    #A4 size at 96 PPI
    new_image = new_image.resize((794, 1123), Image.LANCZOS)
    #archive: Image.Resampling.LANCZOS

    image_tk = ImageTk.PhotoImage(new_image)
    label = ttk.Label(newfile, image=image_tk)
    label.image = image_tk
    label.pack()
    
#Part 1: Set up the search query for employee names

ttk.Label(master, text = "First Name").place(x = 10, y = 10)
ttk.Label(master, text = "Last Name").place(x = 10, y = 50)

#Submit button
ttk.Label(master, text= "Submit").place(x = 20, y = 90)

#Add entry blocks to the GUI
e1 = ttk.Entry(master)
e2 = ttk.Entry(master)

#Buttons to have (have after entry so it can reference entry)
submit_button = ttk.Button(master, text = "Submit", width = 19, command = 
                       search_name)

#Add entry/buttons to the window
e1.place(x = 90, y = 3)
e2.place(x = 90, y = 43)
submit_button.place(x = 90, y = 85)

#Error label, is empty by default
error_label = ttk.Label(master, text= "", foreground= "red")
error_label.place(x = 90, y = 123)

notebook = ttk.Notebook(master)

def hitEnter(event):
    if "focus" in e1.state() or "focus" in e2.state():
        search_name()

master.bind("<Return>", hitEnter)

#Perry logo:
perry_logo = Image.open("C:\\Perry Files\\Scripts\\perry logo.png")
perry_logo = perry_logo.resize((306, 344), Image.LANCZOS)
logo_photo = ImageTk.PhotoImage(perry_logo)


perry_label = Label(master, image = logo_photo)
perry_label.place(x = -10, y = 140)
#master.logo_photo = logo_photo

#Run the function "exit_save" when the application is closed
master.protocol("WM_DELETE_WINDOW", exit_save)

master.mainloop()