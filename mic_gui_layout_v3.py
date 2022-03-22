'''
PROMA

One window (App)
Four frames (MainPage, UpdateSummary, UpdateConfirmation, UpdateCancel)

'''



# %% Imports

## Tkinter imports
import tkinter as tk
from tkinter import messagebox
import tkinter.ttk as ttk
import tkinter.font as tkFont

## other imports + encryption
import smtplib
import os
#import pandas as pd
from datetime import date
from datetime import datetime
import requests
from cryptography.fernet import Fernet

## gspread imports
import gspread
from oauth2client.service_account import ServiceAccountCredentials

name = os.environ.get("USERNAME")
current_date = date.today()

#%% Connect to gsheet

scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"] 
creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope) 
client = gspread.authorize(creds)

# existing spreadsheet
spreadsheet = client.open("Database_v1")
pic_assignment = client.open("Database_v1").worksheet("pic_assignment")
mic_update = client.open("Database_v1").worksheet("mic_update")
email_update = client.open("Database_v1").worksheet("email_update")
list_of_staff = client.open("Database_v1").worksheet("list_of_staff")
activity_log = client.open("Database_v1").worksheet("gui_activity_log")

# list of dictionaries
pic_assignment_list = pic_assignment.get_all_records() #returns a list of dictionaries
mic_update_list = mic_update.get_all_records()

# dataframes
#pic_df = pd.DataFrame(pic_assignment_list)

#mic_df = pd.DataFrame(mic_update_list)
#mic_df['date'] = pd.to_datetime(mic_df['date'], format='%d/%m/%Y').dt.date #formatting

#email_df = pd.DataFrame(email_update)
#staff_df = pd.DataFrame(list_of_staff)

#%% Encryption and decryption

def generate_master_key():
    my_key_file = "encryption_key.csv"
    if os.path.exists(my_key_file):
        with open(my_key_file, 'rb') as myfile:
            master_key = myfile.read()
    else:
        master_key = Fernet.generate_key()
        with open(my_key_file, 'wb') as myfile:
            myfile.write(master_key)
    return master_key
    del master_key

key = generate_master_key()
fernet = Fernet(key)

def encrypt(data):
    return fernet.encrypt(str(data).encode()).decode()

def decrypt(data):
    return fernet.decrypt(str(data).encode()).decode()

def decrypt_dict_list(dict_list):
    for i in dict_list:
        for k,v in i.items():
            i[k] = decrypt(v)
    return dict_list

#%% Decrypted dictionaries
    
#pic_assignment_list = decrypt_dict_list(pic_assignment_list)
#mic_update_list = decrypt_dict_list(mic_update_list)

# convert date string to date
for i in mic_update_list:
    i["date"] =  datetime.strptime(i["date"], "%d/%m/%Y").date()
    
#%% Controller class

class App(tk.Tk):
    
    def __init__(self):
        tk.Tk.__init__(self) # create main window
        self.title("Prospect Tracker") # title of window
        self.resizable(False, False) # not resizable
        
        container = tk.Frame(self) # create main frame
        container.pack(side='top', fill='both', expand=True)
        
        # set common GUI style
        self.frames = {} # dict to store all frames
        self.framesize = {} # dict to store all frame sizes
        
        # loop to initialize all frames 
        for F in (MainPage, UpdateSummary, UpdateConfirmation, UpdateCancel):
            page_name = F.__name__
            frame = F(parent=container, controller=self) # new frames will be created in container frame which is in controller window
            self.frames[page_name] = frame
            self.framesize[page_name] = frame.size
            frame.grid(row=0, column=0, sticky='nsew')

        self.show_frame('MainPage') # set MainPage as the first page
 
    
    def show_frame(self, page_name):
        '''show a frame for the given page name'''
        # get frame size of interest
        frame_size = self.framesize[page_name]
        
        # change the window size to fit the frame size
        self.geometry(frame_size)
       
        # get frame of interest
        frame = self.frames[page_name]
       
        # show frame
        frame.tkraise()
        
    def get_page(self, page_class):
        return self.frames[page_class]
    
    def print_log(self, object_name, action):
        log = [datetime.now(), name, object_name]
        encrypted_log = [encrypt(i) for i in log]
        activity_log.append_row(encrypted_log)
        return
                
#%% Main Page
        
class MainPage(tk.Frame):
    
    def __init__(self, parent, controller):
        self.controller = controller 
        
        self.size = '450x556'
        
        tk.Frame.__init__(self, parent, bg='#ffffff', width=450, height=556)
        self.grid_propagate(0)
        
        self.option_add("*TCombobox*Listbox*selectBackground", '#009cde')

        # font style
        self.fstyle = tkFont.Font(family='Roboto', size=10, weight="bold")
        
        options = {'font':self.fstyle, 'bg':'#ffffff'}
        btn_opt = {'fg':'#ffffff', 'activebackground':'#878787',
                   'borderwidth':0, 'font':self.fstyle}
        
        self.temp_status = tk.StringVar()
        self.temp_current_update = tk.StringVar()
        self.temp_past_update = tk.StringVar()
        self.temp_prospect = tk.StringVar()
        
        self.past_updates_df = []
        self.name = os.environ.get("USERNAME")
        self.mic_email = self.name.lower() + "@rsmsingapore.sg" 
        
        #button functions
        def get_list():
            self.controller.print_log("main pg - get prospects button", "click")  
            prospects_options = []
            #prospects_options = pic_df.loc[(pic_df["mic"] == self.mic_email) & (pic_df["latest_status"] == "Working")]["prospect"].values.tolist()
            
            for i in pic_assignment_list:
                if ((i["mic"] == self.mic_email) and (i["latest_status"] == "Working")): #can change this to whatever
                    prospects_options.append(i["prospect"])
            
            self.cbb_prospect = ttk.Combobox(self, textvariable=self.temp_prospect, values=prospects_options)
            self.cbb_prospect.place(x=98, y=92)
            return self.cbb_prospect
        
            # # do we need it
            # except requests.exceptions.ConnectionError:
            #     self.controller.print_log("main pg - get prospects button", "click")  
            #     notif = tk.Label(self, text = "Connection Error", fg = "red")
            #     notif.grid(row = 2, sticky =tk.W , padx = 5)
            #     self.controller.print_log("main pg - connection error message", "impression")  
            #     return notif
        
        def get_friday(input):
            d = input.toordinal()
            sunday = d - (d % 7)
            friday = sunday - 2
            friday = date.fromordinal(friday)
            return friday
        
        def get_last_weeks_update():
            # get temp_pic_assignment_id and latest status of prospects
            try:
                self.temp_pic_assignment_id = ""
                latest_status = ""
                for i in pic_assignment_list:
                    if ((i["mic"] == self.mic_email) and (i["prospect"] == self.temp_prospect.get())): 
                        self.temp_pic_assignment_id = i["pic_assignment_id"]
                        latest_status = i["latest_status"]
                
                if self.temp_pic_assignment_id == "":
                    messagebox.showerror("Error", "Prospect not selected. Please select and refresh again.")
                    self.controller.print_log("main pg - get updates button - no prospect selected error message", "impression")  


                # get updates before this week
                self.cut_off = get_friday(current_date)
                past_update_list = []
                for i in mic_update_list:
                    if ((i["pic_assignment_id"] == self.temp_pic_assignment_id) and (i["date"] < self.cut_off)):
                        past_update_list.append(i)

                latest_past_update = ''
                if len(past_update_list) == 0:
                    latest_past_update = "No past updates available."
                
                else:
                    for i in past_update_list:
                        if i["date"] == max(i["date"] for i in past_update_list):
                            latest_past_update = i["remark"]
                                            
                self.txt_lupdate.config(state = 'normal')
                self.txt_lupdate.delete(1.0, 'end')
                self.txt_lupdate.insert('end', latest_past_update)
                self.txt_lupdate.config(state = 'disabled')
                
                return self.txt_lupdate, self.temp_pic_assignment_id, self.cut_off ,self.cbb_status.set(latest_status)
            except: #error message doesnt appear cos gui will just say "no past updates"
                messagebox.showerror("Error", "Prospect not selected. Please select and refresh again.")

        def get_current_update():
            current_update_list = []
            for i in mic_update_list:
                if ((i["pic_assignment_id"] == self.temp_pic_assignment_id) and (i["date"] >= self.cut_off)):
                    current_update_list.append(i)
            
            latest_current_update = ""
            for i in current_update_list:
                if i["date"] == max(i["date"] for i in current_update_list):
                    latest_current_update = i["remark"]
            
            if len(latest_current_update) == 0:
                latest_current_update = "NIL/no updates submitted this week"
                        
            self.txt_cupdate.delete(1.0, "end")
            # except: 
            #     latest_current_update = "no updates from this week."
            #     self.txt_cupdate.delete(1.0, "end")
            return self.txt_cupdate.insert(1.0, latest_current_update)
        
        def get_updates():
            self.controller.print_log("main pg - get updates button", "click")  
            return get_last_weeks_update(), get_current_update()
        
        def update_summary():
            self.controller.print_log("main pg - update button", "click")  
            
            if self.temp_prospect.get() == "" or self.temp_status.get() == "" or self.temp_update.get() == "":
                messagebox.showerror("Error", "Please fill in all required fields before updating.")
                self.controller.print_log("main pg - update button - fill in all required fields error message", "impression")  
                
            else:
                self.controller.show_frame('UpdateSummary')
                self.controller.print_log("update summary pg", "open page")  
                
                summary_pg = self.controller.get_page('UpdateSummary')
                summary_pg.lbl_name.config(text = "User: " + self.name)
                summary_pg.lbl_prospect.config(text = "Prospect: " + self.temp_prospect.get())
                summary_pg.lbl_status.config(text = "Status: " + self.temp_status.get())
                summary_pg.lbl_cupdate.config(text = "Update: ")
                
                ## clear all text, re-update edits and disable editing for confirmation
                summary_pg.txt_cupdate.config(state = 'normal')
                summary_pg.txt_cupdate.delete('1.0', 'end')
                summary_pg.txt_cupdate.insert('end', self.txt_cupdate.get("1.0", 'end-1c'))
                summary_pg.txt_cupdate.config(state = 'disabled')
                
                # run function so we get temp_pic_assignment_id even if user doesn't click "refresh"
                get_last_weeks_update()
                return
        
        def cancel_update():
            self.controller.print_log("main pg - cancel update button", "click")  
            self.controller.show_frame('UpdateCancel')
            self.controller.print_log("update cancel pg", "open page") 
            return
        
        #load prospects once prospect tracker is opened
        get_list()
        
        # Layout 1. Label
        self.lbl_line = tk.Label(self, fg='#878787', bg='#ffffff', 
                    text='________________________________________________________________________________________________')
        self.lbl_line.place(x=-10, y=30)

        self.lbl_info = tk.Label(self, text='Welcome to PROMA, ' + self.name + '!', 
                                 **options)
        self.lbl_info.place(x=26, y=56)

        self.lbl_prospect = tk.Label(self, text='Prospect:', **options)
        self.lbl_prospect.place(x=26, y=92)

        self.lbl_status = tk.Label(self, text='Status:', **options)
        self.lbl_status.place(x=26, y=122)

        self.lbl_lupdate = tk.Label(self, text="Last Week's Update:", 
                                    **options)
        self.lbl_lupdate.place(x=26, y=163)

        self.lbl_cupdate = tk.Label(self, text="Current Week's Update:",
                                    **options)
        self.lbl_cupdate.place(x=26, y=302)
        
        # Layout 2. Combobox
        # default_values is just an example
        self.status_values = ["Working", "Won", "Lost", "Disqualified", "Not a Target", "Nurture"]

        self.cbb_status = ttk.Combobox(self, textvariable=self.temp_status, values=self.status_values)
        self.cbb_status.place(x=98, y=122)
        

        # Layout 3. Entry
        self.txt_lupdate = tk.Text(self, borderwidth=0, bg='#DEDEDE')
        self.txt_lupdate.place(x=29, y=181, width=393, height=105)
        self.txt_lupdate.config(state = 'disabled')

        self.txt_cupdate = tk.Text(self, wrap=tk.WORD, borderwidth=0, bg='#DEDEDE')
        self.txt_cupdate.place(x=29, y=320, width=393, height=129)


        # Layout 4. Button
        self.btn_read = tk.Button(self, text='Read List', bg='#009cde', 
                                  **btn_opt, command = get_list)
        self.btn_refresh = tk.Button(self, text='Refresh', bg='#009cde', 
                                     **btn_opt, command = get_updates)

        self.btn_read.place(x=10, y=10, width=85, height=25)
        self.btn_refresh.place(x=355, y=10, width=85, height=25)
        self.btn_cancel = tk.Button(self, text='CANCEL', 
                                    bg='#D02D2D',
                                    command=cancel_update,
                                    **btn_opt)
        self.btn_update = tk.Button(self, text='UPDATE',
                                    bg='#4F9A35',
                                    command=update_summary,
                                    **btn_opt)
        
        self.btn_cancel.place(x=29, y=510, width=83, height=25)
        self.btn_update.place(x=339, y=510, width=83, height=25) 
        

    def __name__():
        return 'MainPage'


#%% Update Summary Page
        
class UpdateSummary(tk.Frame):
    
    def __init__(self, parent, controller):
        self.size = '335x363'
        
        tk.Frame.__init__(self, parent, bg='#ffffff', width=335, height=363)

        self.controller = controller
        
        # set font style
        self.fstyle = tkFont.Font(family='Roboto', size=10, weight="bold")

        options = {'font':self.fstyle, 'bg':'#ffffff'}
        btn_opt = {'activebackground':'#878787',
                   'borderwidth':0, 'font':self.fstyle}
        
        def encrypt_upload():
            self.controller.print_log("update summary pg - confirm upload button", "click") 

            main_pg = self.controller.get_page('MainPage')
            confirmation_pg = self.controller.get_page('UpdateConfirmation')

            # get variables; why cant we just make a list directly
            new_update = [{"pic_assignment_id": encrypt(str(main_pg.temp_pic_assignment_id)),  #sth problematic here
                          "mic_update_id": encrypt(str( max([i["mic_update_id"] for i in mic_update_list])+1)),
                          "remark": encrypt(main_pg.temp_current_update.get()), 
                          "date": encrypt(current_date.strftime("%m/%d/%Y")), 
                          "status": encrypt(main_pg.temp_status.get()),
                          "update_by": encrypt(name)
                          }]
            
            header_to_key = {
                    "pic_assignment_id": "pic_assignment_id", 
                    "mic_update_id": "mic_update_id", 
                    "remark": "remark", 
                    "date": "date", 
                    "status":"status",
                    "update_by":"update_by"
            }
            
            headers = mic_update.row_values(1)
            put_values = []
            for v in new_update:
                temp = []
                for h in headers:
                    temp.append(v[header_to_key[h]])
                put_values.append(temp)
            spreadsheet.values_append("mic_update_e", {'valueInputOption': 'USER_ENTERED'}, {'values': put_values})    
            
            # update confirmation message'
            confirmation_pg.lbl_info.config(text='Prospect ' + main_pg.temp_prospect.get() + ' updated!')
            
            # show page
            self.controller.show_frame('UpdateConfirmation')
            self.controller.print_log("update confirmation pg", "open page")  
        
        def go_back():
            self.controller.print_log("update summary pg - go back", "click")  
            self.controller.show_frame('MainPage')
            self.controller.print_log("main page", "open page")  

            main_pg = self.controller.get_page('MainPage')
            
            self.txt_cupdate.config(state = 'normal')
            self.txt_cupdate.insert('end', self.txt_cupdate.get("1.0", 'end-1c'))
            return
        
        # layout - label 
        self.lbl_info = tk.Label(self,
                                 text='Please confirm the information below:',
                                 **options)   
        self.lbl_info.place(x=13, y=21)
        self.lbl_name = tk.Label(self, text='User: <username>',
                                 **options)
        self.lbl_name.place(x=92, y=57)
        self.lbl_prospect = tk.Label(self, text='Prospect: <pros name>',
                                     **options)
        self.lbl_prospect.place(x=92, y=77)
        self.lbl_status = tk.Label(self, text='Status: <>',
                                     **options)
        self.lbl_status.place(x=92, y=97)
        self.lbl_cupdate = tk.Label(self, text='Current Update:',
                                    **options)
        self.lbl_cupdate.place(x=36, y=135)
        
        # layout - entry
        self.txt_cupdate = tk.Text(self, wrap=tk.WORD, borderwidth=0, bg='#DEDEDE')
        self.txt_cupdate.place(x=36, y=155, width=263, height=129)
        
        # layout - button
        self.btn_goback = tk.Button(self, text='GO BACK', bg='#C4C4C4',
                                    command=go_back,
                                    **btn_opt)
        self.btn_goback.place(x=89, y=314, width=70, height=25)
        self.btn_confirm = tk.Button(self, text='CONFIRM', fg='#ffffff',
                                     command=encrypt_upload,
                                     bg='#009cde',
                                     **btn_opt)
        self.btn_confirm.place(x=175, y=314, width=70, height=25)
    

    def __name__():
        return 'UpdateSummary'


#%% Update Confirmation Page
        
class UpdateConfirmation(tk.Frame):
    
    def __init__(self, parent, controller):
        self.size = '276x167'
        tk.Frame.__init__(self, parent, bg='#ffffff', width=276, height=167)
        self.controller = controller
        
        # set font style
        self.fstyle = tkFont.Font(family='Roboto', size=10, weight="bold")

        options = {'font':self.fstyle, 'bg':'#ffffff'}
        btn_opt = {'activebackground':'#878787',
                   'borderwidth':0, 'font':self.fstyle}
        
        # connect to main page
        main_pg = self.controller.get_page('MainPage')
        
        # button functions
        def new_update(): #same as "confirm cancel" button
            self.controller.print_log("update confirmation pg - input new update button", "click")  
            
            start_pg = self.controller.get_page('MainPage')
            start_pg.temp_prospect.set("")
            
            start_pg.txt_lupdate.config(state = 'normal')
            start_pg.txt_lupdate.delete(1.0, 'end')
            start_pg.txt_lupdate.config(state = 'disabled')
                
            start_pg.txt_cupdate.delete(1.0, "end")
            start_pg.temp_status.set("")
            
            #show page
            self.controller.show_frame('MainPage')
            self.controller.print_log("main pg", "open pg")  

        # def close_proma():
        #     self.controller.print_log("update confirmation pg - close PROMA button", "click")   
        #     root.destroy()
        
        # labels
        self.lbl_info = tk.Label(self,
                                 text='Prospect updated!', # message changes when encrypt upload runs
                                 **options)
        self.lbl_info.place(x=20, y=16)
        self.lbl_updateby = tk.Label(self,
                                 text='Updated By: ' + name,
                                 **options)
        self.lbl_updateby.place(x=70, y=56)
        self.lbl_date = tk.Label(self,
                                 text='Date: ' + str(date.today()),
                                 **options)
        self.lbl_date.place(x=70, y=76)
        
        # buttons
        self.btn_ok = tk.Button(self, text='OK', fg='#ffffff',
                                     command=new_update,
                                     bg='#009cde',
                                     **btn_opt)
        self.btn_ok.place(x=103, y=119, width=70, height=25)
        

    def __name__():
        return 'UpdateConfirmation'


#%% Cancel Warning Page
class UpdateCancel(tk.Frame):
    
    def __init__(self, parent, controller):
        self.size = '276x167'
        tk.Frame.__init__(self, parent, bg='#ffffff', width=276, height=167)
        self.controller = controller
        
        # set font style
        self.fstyle = tkFont.Font(family='Roboto', size=10, weight="bold")

        options = {'font':self.fstyle, 'bg':'#ffffff'}
        btn_opt = {'activebackground':'#878787',
                   'borderwidth':0, 'font':self.fstyle}
        
        # button functions
        def dont_cancel():
            self.controller.print_log("update cancel pg - dont cancel button", "click")   
            
            self.controller.show_frame('MainPage')
            self.controller.print_log("start pg", "open pg")  


        def confirm_cancel(): # same as new update button above
            self.controller.print_log("update cancel pg - confirm cancel button", "click")   

            start_pg = self.controller.get_page('MainPage')
            start_pg.temp_prospect.set("")
            
            start_pg.txt_lupdate.config(state = 'normal')
            start_pg.txt_lupdate.delete(1.0, 'end')
            start_pg.txt_lupdate.config(state = 'disabled')
                
            start_pg.txt_cupdate.delete(1.0, "end")
            start_pg.temp_status.set("")

            #show page
            self.controller.show_frame('MainPage')
            self.controller.print_log("start pg", "open pg")  
            return
        
        # labels
        self.lbl_warning = tk.Label(self,
                                 text='Are you sure you want to cancel?\n\nAll information will be lost.',
                                 **options)
        self.lbl_warning.configure(anchor='center')
        self.lbl_warning.place(x=40, y=32)
        
        # buttons
        self.btn_no = tk.Button(self, text='NO', fg='#ffffff',
                                     command=dont_cancel,
                                     bg='#009cde',
                                     **btn_opt)
        self.btn_no.place(x=75, y=105, width=60, height=25)
        self.btn_yes = tk.Button(self, text='YES', fg='#ffffff',
                                     command=confirm_cancel,
                                     bg='#009cde',
                                     **btn_opt)
        self.btn_yes.place(x=141, y=105, width=60, height=25)
        
        
    def __name__():
        return 'UpdateCancel'


#%% Run App
root = App()
root.mainloop()





