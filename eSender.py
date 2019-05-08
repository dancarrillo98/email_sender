#-------------------------------------------------------------------
# imports
#-------------------------------------------------------------------

from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import smtplib
import datetime

#--------------------------------------------------------------------
# constants
#--------------------------------------------------------------------

GMAIL_HOST = 'smtp.gmail.com' #need all of the providers
GMAIL_PORT = 587

#--------------------------------------------------------------------
# global variables
#--------------------------------------------------------------------

# stores the email and name extracted from the excel doc
info = []

# stores the email template extracted from the text doc
template = ''

# text loaded boolean
text_loaded = False

# excel loaded boolean
excel_loaded = False

# subject boolean
subject_entered = False

# email address boolean
address_entered = False

# email address password boolean
password_entered = False

#---------------------------------------------------------------------
# making the email server
#--------------------------------------------------------------------

# itorating through the contacts
def itorate_through_emails():
  
  if subject.get() != "":
    subject_entered = True

  if email_address.get() != "":
    address_entered = True

  if password.get() != "":
    password_entered = True

  
  if excel_loaded and text_loaded and subject_entered and\
  address_entered and password_entered:
    for contact in info:
      email_sender( contact[1], contact[0] )

    message = "All emails have been sent as of {}!".format(
        datetime.datetime.now() )
    popup( message )
  
  else:
    message = "Not All Required Fields Have Been Filled Out" 
    popup( message )


# sending the email
def email_sender( contact_email, contact_name ):
  
  # starting the email server
  sender = smtplib.SMTP( GMAIL_HOST, GMAIL_PORT )
  sender.starttls()
  try:
    sender.login( email_address.get(), password.get() )
  except:
    message = "Invalid Email Address or Email Password" 
    popup( message )
    return

  # making the message object
  message = MIMEMultipart()
  message[ 'From' ] = email_address.get()
  message[ 'To' ] = contact_email
  message[ 'Subject' ] = subject.get()
  
  try:
    msg = template.substitute( NAME = contact_name )
    message.attach( MIMEText( msg, 'plain' ) )
  except:
    message = "There is no attribute ${NAME} in the text document, no\
 emails have been sent"
    popup( message )
    return

  # sending the email
  try:
    sender.send_message( message )
  except:
    message = "{} is not a valid email, other emails will continue to\
 be sent".format( contact_email )
    popup( message )
    return

  # closing the email server
  sender.quit()

# constructing the pop-up that appears after sending all the emails
def popup( message ):
  popup = Tk()

  label = ttk.Label( popup, text = message )
  label.pack()

  okay_button = ttk.Button( popup, text = "Okay", command =
      popup.destroy )
  okay_button.pack()

  popup.mainloop()

#--------------------------------------------------------------------
# getting the Documents
#--------------------------------------------------------------------

# getting the document, making a panda, and setting info
def excel_document():
  excel_doc = filedialog.askopenfilename( filetypes = [ ( "Excel\
          Documents", "*.xlsx" ) ] )

  if excel_doc == ():
    message = "No Excel Document Uploaded"
    popup( message )
    return

  table = pd.read_excel( excel_doc )
  
  global info
  info = table.get_values()
  excel_doc_check.config( text = excel_doc )

  global excel_loaded
  excel_loaded = True

#getting the text document, reading it, and setting template
def text_document():
  text_doc = filedialog.askopenfilename( filetypes = [ ( "Text\
 Documents", "*.txt" ) ] )

  if text_doc == ():
    message = "No text file uploaded"
    text_doc_check.config( text = "" )
    popup( message )
    return

  with open( text_doc, 'r', encoding = 'utf-8' ) as template_file:
    text = template_file.read()

  global template
  template = Template( text )
  text_doc_check.config( text = text_doc )

  global text_loaded
  text_loaded = True

#--------------------------------------------------------------------
# making the gui
#--------------------------------------------------------------------

# making the main pane, setting its geometry, and setting the title
tk = Tk()
tk.geometry( "500x400" )
tk.title( "e-Sender" )

# making the tab manager and calling the different tab constructors
tab_manager = ttk.Notebook( tk )

instructions_tab = ttk.Frame( tab_manager )
tab_manager.add( instructions_tab, text = "Instructions" )
tab_manager.pack( expand = 1, fill = "both" )

input_tab = ttk.Frame( tab_manager )
tab_manager.add( input_tab, text = "Email, Documents, and Send Button" )
tab_manager.pack( expand = 1, fill = "both" )

# populating the instructions tab
title = "e-Sender"
setup = "Welcome to Lab e-Sender, a program designed to help\
 you get into labs faster by using a python program to read a text file\
 and an excel document to send mass emails from your school email\
 account to professors you're interested in working for.\n\
First, you're going to need to change some email settings,\
 namely your email security settings, this is just so the program can\
 send emails from your account. Once your done, the settings can easily\
 be changed back so your account is safe. What your going to need\
 to do is the following: \n\
(1) Log into your school email account\n\
(2) Go to the upper right corner with the circle symbol with your \
initials (or your picture if you changed it) and click 'Account\
 Settings' \n\
(3) Look at the left column panel and click 'Security' \n\
(4) Scroll down to 'Less Secure App Access' and turn it on. This will\
 allow the program to access your email account and send emails from\
 your address \n\
(5) Import your text document in .txt format, make sure to put ${NAME}\
 where the professors name is going to go\n\
(6) Import your excel document in .xlsx format. The program will read\
 the first two columns and expects that the first column has 'Names:'\
 as its header and the second column has 'Emails:' as its header\n\
(7) press the send button and watch your emails get sent!"

title_label = Label( instructions_tab, text = title, fg = "blue", font
    = ( "System", 25 ) )
title_label.pack()

setup_label = Label( instructions_tab, text = setup, wraplength =
    400, justify = LEFT )
setup_label.pack()

# populating the input tab
whitespace0 = Label( input_tab, text = "  " )
email_address_prompt = Label( input_tab, text = "Email Address:",
    justify = LEFT )
password_prompt = Label( input_tab, text = "Email Password:", justify
    = LEFT )
subject_prompt = Label( input_tab, text = "Email Subject Line:",
    justify = LEFT )
whitespace1 = Label( input_tab, text = "  " )

whitespace0.grid( rowspan = 2, columnspan = 2, row = 0, column = 0 )
email_address_prompt.grid( row = 2, column = 0 )
password_prompt.grid( row = 3, column = 0 )
subject_prompt.grid( row = 4, column =  0 )
whitespace1.grid( rowspan = 2, columnspan = 6, row = 5, column = 0 )



email_address = Entry( input_tab )
password = Entry( input_tab )
subject = Entry( input_tab )

email_address.grid( row = 2, column = 1 )
password.grid( row = 3, column = 1 )
subject.grid( row = 4, column = 1 )



text_doc_check = Label( input_tab )
text_doc_button = Button( input_tab, text = "Upload Text Document",
    command = text_document, pady = 25 )
excel_doc_check = Label( input_tab )
excel_doc_button = Button( input_tab, text = "Upload Excel Document",
    command = excel_document, pady = 25 )

text_doc_button.grid( row = 10, column = 0 )
text_doc_check.grid( row = 10, column = 1 )
excel_doc_button.grid( row = 11, column = 0 )
excel_doc_check.grid( row = 11, column = 1 )



whitespace2 = Label( input_tab, text = "  " )
send = Button( input_tab, text = "Send Emails!", command =
    itorate_through_emails )

whitespace2.grid( columnspan = 2, rowspan = 8, row = 12, column = 0 )
send.grid( row = 21, column = 0 )

#--------------------------------------------------------------------
# running the program
#--------------------------------------------------------------------

tk.mainloop()




