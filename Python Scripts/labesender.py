from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
import smtplib
from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import datetime

def array_maker():
  global excel_doc
  if excel_doc.endswith( ".csv" ):
    panda = pd.read_csv( excel_doc )
  else:
    panda = pd.read_excel( excel_doc )
  return panda.get_values()

def get_excel_doc():
  global excel_doc
  excel_doc = filedialog.askopenfilename( filetypes = ( ( "Excel Document",
        "*.xlsx" ), ( "CSV Document", "*.csv" ) ) )
  if excel_doc.endswith( ".csv" ):
    panda = pd.read_csv( doc )
  else:
    panda = pd.read_excel(excel_doc)
  temp = panda.get_values()
  headers = panda.columns
  columns = len( temp )
  rows = len( temp[0] )
  for col in range(columns + 1):
    for rowa in range(rows):
      if rowa == 0:
        textz = headers[col]
      else:
        textz = temp[rowa-1][col]
      label = Label( tab_spreadsheet, text = textz )
      label.grid( row = rowa, column = col )
      
def get_text_document():
  global doc
  doc = filedialog.askopenfilename( filetypes = ( ( "Word Document",
        "*.docx" ), ( "Text file Document", "*.txt" ) ) )
  text = ""
  if doc.endswith( ".docx" ):
    print("whoops")
  if doc.endswith( ".txt" ):
    with open(doc, 'r', encoding='utf-8') as template_file:
        text = template_file.read()
  text_box.config(state=NORMAL)
  text_box.insert(END, text)
  text_box.config(state=DISABLED)

def email(message_object, doctor_name):
  global doc
  if doc.endswith( ".docx" ):
    print("whoops")
  if doc.endswith( ".txt" ):
    with open(doc, 'r', encoding='utf-8') as template_file:
        text = template_file.read()
  message_template = Template( text )
  message = message_template.substitute(NAME=doctor_name)
  message_object.attach(MIMEText(message, 'plain'))
            
def read_template(filename):
  with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
  return Template(template_file_content)

def sender(to, name):

  HOST = 'smtp.gmail.com'#need to get all the email providers
  PORT = 587

  
  s = smtplib.SMTP(HOST, PORT)
  s.starttls()
  s.login(Email_Adress_Button.get(), Email_Password.get())

  msg = MIMEMultipart()

  msg['From']=Email_Adress_Button.get()
  msg['To']=to
  msg['Subject']=Email_Subject_Title_Button.get()
    
  email(msg,name)    
  s.send_message(msg)
  del msg

  s.quit()
    
def itorator():
  info = array_maker()
  for i in info:
      sender(i[1],i[0])
  popup = Tk()
  msg = "All emails have been send as of at {}!".format( \
        datetime.datetime.now() ) 
  label = ttk.Label(popup, text= msg)
  label.pack()
  B1 = ttk.Button(popup, text="Okay", command = popup.destroy)
  B1.pack()
  popup.mainloop()
       

#runnning it all
tk = Tk()
tk.geometry("500x400")

#Title
tk.title( "LabESender" )

#Making the tabs
tab_controller = ttk.Notebook( tk )

tab_instructions = ttk.Frame( tab_controller )
tab_controller.add( tab_instructions, text = 'Instructions' )
tab_controller.pack( expand=1, fill='both' )


tab_variables = ttk.Frame( tab_controller )
tab_controller.add( tab_variables, text = 'Required Input' )
tab_controller.pack( expand=1, fill='both' )

tab_email_template = ttk.Frame( tab_controller )
tab_controller.add( tab_email_template, text = 'Email Template' )
tab_controller.pack( expand=1, fill='both' )

tab_spreadsheet = ttk.Frame( tab_controller )
tab_controller.add( tab_spreadsheet, text = 'Spreadsheet' )
tab_controller.pack( expand=1, fill='both' )

#Grid on the INSTRUCTION PAGE
Label(tab_variables, text='Your Email Address:').grid(row=0)
Label(tab_variables, text='Your Email Password:').grid(row=1)
Label(tab_variables, text='The Email Subject Title:').grid(row=2)

Email_Adress_Button = Entry(tab_variables)
Email_Password = Entry(tab_variables) 
Email_Subject_Title_Button = Entry(tab_variables) 
Excel_Doc_Button = Entry(tab_variables)
Email_Template_Button = Entry(tab_variables) 
Email_Adress_Button.grid(row=0, column=1)
Email_Password.grid(row=1, column=1) 
Email_Subject_Title_Button.grid(row=2, column=1) 

#Text box on the EMAIL TEMPLATE PAGE
text_box = Text( tab_email_template, state=DISABLED )
text_box.pack( expand=True, fill='both' )


#Menu's
menu = Menu( tk ) 
tk.config( menu = menu )
  
filemenu = Menu( menu ) 

importmenu = Menu( menu )
menu.add_cascade( label = 'Import', menu = importmenu )
importmenu.add_command( label = 'Import Spreadsheet or .csv File',
    command = get_excel_doc)#TODO
importmenu.add_command( label = 'Import Word or .txt File', command =
    get_text_document )#TODO

sendmenu = Menu( menu ) 
menu.add_cascade( label='Send', menu = sendmenu,  )
sendmenu.add_command( label = 'To All', command = itorator ) #TO DO

excel_doc = ''
doc = ''

tk.mainloop()
