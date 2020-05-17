__version__ = '3.7.5'
__author__ = 'Dusan Petronijevic'
__name__ = 'DuMail SENDER with Attahcment GUI'

__check_email_options__ = 'https://myaccount.google.com/lesssecureapps' # ON

import time

try:
    import win32gui, win32con
except:
    print('Install PyWin32 Module')
    time.sleep(5)
    quit()
    
try:
    import pygetwindow as gw
except:
    print('Install PyGetWindow module')
    time.sleep(5)
    quit()

import os

def _hide_console():
    The_program_to_hide = win32gui.GetForegroundWindow()
    win32gui.ShowWindow(The_program_to_hide , win32con.SW_HIDE)
_hide_console() # hide console method

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase
from email import encoders
import time
from tkinter import *
import tkinter as tk
from tkinter.filedialog import askopenfilename


class DuMail_Client(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)

        self.__login__()

    def __send_mail__(self):
        
        self.__sent_from__ = self.__gmail_user__

        self.__email_send_to__ = self.var_txt_send_to.get()

        self.__sent_to__ = [self.__email_send_to__, self.__email_send_to__]

        self.__sent_subject__ = self.var_txt_subject.get()

        self.__sent_body__ = self.__message__.get('0.0',END)
        
        self.__message__ = MIMEMultipart() 
           
        self.__message__['From'] = self.__sent_from__ 
          
        self.__message__['To'] = self.__email_send_to__
          
        self.__message__['Subject'] = self.__sent_subject__
          
        self.body = '' + self.__sent_body__

        self.__message__.attach(MIMEText(self.body, 'plain'))

        try:
            self.filename = self.__file_path__
            
            self.attachment = open(self.filename,'rb')
                  
            self.__payload__ = MIMEBase('application', 'octet-stream') 

            self.__payload__.set_payload((self.attachment).read())

            encoders.encode_base64(self.__payload__) 
               
            self.__payload__.add_header('Content-Disposition', "attachment; filename= %s" % self.filename) 

            self.__message__.attach(self.__payload__)


            # SEND EMAIL - only text & attachment

            self.txt = self.__message__.as_string() 
              
            self.__mail_server__.sendmail(self.__sent_from__, self.__sent_to__, self.txt) 

            self.__mail_server__.quit()
            
        except:
            # SEND EMAIL - only text
            self.txt = self.__message__.as_string()
            
            self.__mail_server__.sendmail(self.__sent_from__, self.__sent_to__, self.txt) 
              
            self.__mail_server__.quit()
            

        self.root.destroy() # exit window
        

    def __login_window__(self):
        
        try:
            self.__gmail_user__ = self.var_txt_email.get() # get gmail email
        except:
            print('Niste ukucali email ili isti nije tacan !!!')
            quit()
            
        try:
            self.__gmail_app_password__ = self.var_txt_password.get() # get password from email
        except:
            print('Niste ukucali sifru ili ista nije tacna !!!')
            quit()
        
        try:
            self.__mail_server__ = smtplib.SMTP('smtp.gmail.com', 587) 
              
            self.__mail_server__.starttls() 
              
            self.__mail_server__.login(self.__gmail_user__,self.__gmail_app_password__)
        except:
            print('Ne moze da se uloguje !!!')
            quit()

        self.destroy() # Destroy login window

        self.__send_mail_window__() # Mail sender window


    def __send_mail_window__(self):
        
        self.root = Tk()

        self.__frame__ = Frame(self.root)
        self.__frame__.grid(row = 0,column = 0,padx = 10,pady = 10)
        self.__frame__.configure(bg = 'black')
        
        #Send to - kome
        self.lbl_send_to = Label(self.__frame__,text = 'Send to:')
        self.lbl_send_to.grid(row = 0,column = 0,padx = 10,pady = 10)
        self.lbl_send_to.config(bg = 'black')
        self.lbl_send_to.config(fg = 'white')
        self.lbl_send_to['font'] = ('Calibri','12')

        self.var_txt_send_to = StringVar()
        self.txt_send_to = Entry(self.__frame__,
                                 textvariable = self.var_txt_send_to,
                                 width=70)
        self.txt_send_to.grid(row = 0,column = 1,padx = 10,pady = 10)
        self.txt_send_to.focus()
        self.txt_send_to.config(bg = 'black')
        self.txt_send_to.config(fg = 'white')
        self.txt_send_to['font'] = ('Calibri','12')

        
        #Subject - naslov email-a
        self.lbl_subject = Label(self.__frame__,text='Subject:')
        self.lbl_subject.grid(row = 1,column = 0,padx = 10,pady = 10)
        self.lbl_subject.config(bg = 'black')
        self.lbl_subject.config(fg = 'white')
        self.lbl_subject['font'] = ('Calibri','12')
        
        self.var_txt_subject = StringVar()
        self.txt_subject = Entry(self.__frame__,
                                 textvariable = self.var_txt_subject,
                                 width = 70)
        self.txt_subject.grid(row = 1,column = 1,padx = 10,pady = 10)
        self.txt_subject.config(bg = 'black')
        self.txt_subject.config(fg = 'white')
        self.txt_subject['font'] = ('Calibri','12')
        

        #Attachment message - Label
        self.lbl_attachment = Label(self.__frame__)
        self.lbl_attachment.grid(row = 2,column = 1,padx = 10,pady = 10)
        self.lbl_attachment.config(bg = 'black')
        self.lbl_attachment.config(fg = 'white')
        self.lbl_attachment['font'] = ('Calibri','20','bold')

        def __filename_path__():
            self.__file_path__ = askopenfilename(title = 'Filename?',initialdir = r'./').replace('/','\\')
            self.lbl_attachment['text'] = str(os.path.basename(self.__file_path__))

        #Attachment message - Button
        self.btn_attachment = Button(self.__frame__,text = 'Attachment')
        self.btn_attachment.grid(row = 2,column = 0,padx = 10,pady = 10)
        self.btn_attachment['font'] = ('Calibri',14,'bold')
        self.btn_attachment.config(bg = 'lawn green')
        self.btn_attachment.config(fg = 'black')
        self.btn_attachment['command'] = __filename_path__
        
        #Messaage - poruka
        self.__message__ = Text(self.root)
        self.__message__.grid(row = 1,column = 0,padx = 10,pady = 10)
        self.__message__.config(bg = 'gray13')
        self.__message__.config(fg = 'white')
        self.__message__['font'] = ('Calibri','12')
            
        #Send message
        self.btn_send_msg = Button(self.root,text = 'Send message')
        self.btn_send_msg.grid(row = 2,column = 0,padx = 10,pady = 10)
        self.btn_send_msg.config(bg = 'dark turquoise')
        self.btn_send_msg.config(fg = 'black')
        self.btn_send_msg['font'] = ('Calibri',16,'bold')

        #Button send message - SEND MESSAGE
        self.btn_send_msg['command'] = self.__send_mail__
        
        try:
            self.root.title('Login:' + str(self.var_txt_email.get()))
        except:
            self.root.title('Login:')

        self.x = int(740)
        self.y = int(720)
            
        try:
            self.root.resizable(False,False)
            self.root.geometry(str(self.x)+'x'+str(self.y))
            self.root.configure(bg = 'black')
            self.root.mainloop()
        except:
            print(end='')


    def __login__(self):
        # StringVar
        self.var_txt_email = StringVar()
        self.var_txt_password = StringVar()
        
        # Email - Label
        self.lbl_email = Label(self,text = 'Email:')
        self.lbl_email['width'] = 10
        self.lbl_email.config(bg = 'black')
        self.lbl_email.config(fg = 'white')
        self.lbl_email['font'] = ('Arial', '15','bold')

        # passsword - Label
        self.lbl_password = Label(self,text = 'Password:')
        self.lbl_password['width'] = 10
        self.lbl_password.config(bg='black')
        self.lbl_password.config(fg='white')
        self.lbl_password['font'] = ('Arial', '15','bold')

        # Email - Text
        self.txt_email = Entry(self,textvariable = self.var_txt_email)
        self.txt_email.focus()
        self.txt_email.config(bg = 'gray15')
        self.txt_email.config(fg = 'white')
        self.txt_email['width'] = 40
        self.txt_email['font'] = ('Calibri', '15')

        # Password - Text
        self.txt_password = Entry(self,show = "*",textvariable = self.var_txt_password)
        self.txt_password.config(bg = 'gray15')
        self.txt_password.config(fg = 'white')
        self.txt_password['width'] = 40
        self.txt_password['font'] = ('Calibri', '15')

        # Login - Button
        self.btn_login = Button(self,text='Login',command = self.__login_window__)
        self.btn_login['width'] = 10
        #self.btn_login['height'] = 1
        self.btn_login.config(bg = 'forest green')
        self.btn_login.config(fg = 'white')
        self.btn_login['font'] = ('Verdana', '15','bold')

        self.lbl_email.grid(row = 0,column = 0,padx = 10,pady = 10)
        self.txt_email.grid(row = 0,column = 1,padx = 10,pady = 10)
        self.lbl_password.grid(row = 1,column = 0,padx = 10,pady = 10)
        self.txt_password.grid(row = 1,column = 1,padx=10,pady = 10)
        self.btn_login.grid(row = 2,column = 1,padx = 10,pady = 10)

class Run_App():
    def __init__(self):

        try:
            self.__gui__()
        except:
            print(end='')

    def __gui__(self):
        app = DuMail_Client()
        app.title('DuMail Sender')
        app.resizable(False,False)
        app.configure(bg='black')
        app.geometry('600x200')
        app.mainloop()

app = Run_App()
