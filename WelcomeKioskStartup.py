# Touch Screen https://www.amazon.com/Portable-Security-Touchscreen-1280x800-Raspberry/dp/B0B3DJY8Z1/ref=sr_1_4?crid=3QW6WB6Y3QAOK&keywords=touchscreen%2Bmonitor&qid=1656430782&refinements=p_36%3A-10000&rnid=386442011&s=electronics&sprefix=touchscreen%2Bmonitor%2Caps%2C95&sr=1-4&th=1
# Import module
import time
import os
import smtplib
import webbrowser
from openpyxl.workbook.protection import WorkbookProtection
from openpyxl import load_workbook
from email.message import EmailMessage
from tkinter import *
from tkinter.ttk import *

# excel vars
wb = load_workbook(r"Z:\00 Front Desk Staff\Data\Welcome Kiosk\.xlsx files\kiosk-reports.xlsx")
sheet = wb['walk-in-reports']
sheet.column_dimensions.group(start='A', end='H', hidden=True)
sheet.protection.sheet = True
sheet.protection.password = 'Scott102'
sheet = wb['sign-in-reports']
sheet.column_dimensions.group(start='A', end='D', hidden=True)
sheet.protection.sheet = True
sheet.protection.password = 'Scott102'


# Globals ---> initailization
options = {
           "Karan Venayagamoorthy": "Karan.Venayagamoorthy@colostate.edu",
           "Susan Benzel": "Susan.Benzel@colostate.edu", 
           "Emily Valerioti": "Emily.Valerioti@colostate.edu", 
           "Nickey Pietila": "Nickey.Pietila@colostate.edu", 
           "Bert Vermeulen": "Bert.Vermeulen@colostate.edu",
           "Heather Hall": "Hall,Heather Heather.Hall@colostate.edu", 
           "Claudia Hernandez": "Claudia.Hernandez@colostate.edu", 
           "Jacqui Goldring": "Jacqui.Goldring@colostate.edu",
           "Jose Arce": "Jose.Arce@colostate.edu",
           "Shannon Wagner": "Shannon.Wagner@colostate.edu",
           "Estelle McCabe": "EstelleDavalos.McCabe@colostate.edu",
           "April Undy": "April.Undy@colostate.edu",
           "Teresa Simske": "Teresa.Simske@colostate.edu",
           "Chase Jackson": "Chasen.Jackson@colostate.edu",
           "Other": "EngineerSuccess@colostate.edu"
           }

# String variables
recipient = ""
path = r"Z:\00 Front Desk Staff\Data\Welcome Kiosk\.png files\Karan Venayagamoorthy.png"

# Integers
max_len = -1

def open_walk_in():

    # method to add to excel sheet
    def add_to_walk_in():
        
        sheet = wb['walk-in-reports']
        
        # get and set
        max_row = get_maximum_rows(sheet_object = sheet) + 1
        sheet['A' + str(max_row)] = name_var.get()
        sheet['B' + str(max_row)] = time.asctime(time.localtime(time.time()))
        sheet['C' + str(max_row)] = email_var.get()
        sheet['D' + str(max_row)] = phone_var.get()
        sheet['E' + str(max_row)] = major_var.get()
        sheet['F' + str(max_row)] = who_var.get()
        sheet['G' + str(max_row)] = id_var.get()
        sheet['H' + str(max_row)] = reason_entry.get('1.0','end')
        name_var.set("")
        email_var.set("")
        phone_var.set("")
        major_var.set("")
        who_var.set("")
        id_var.set("")
        reason_entry.delete('1.0','end')
        
        wb.save(r"Z:\00 Front Desk Staff\Data\Welcome Kiosk\.xlsx files\kiosk-reports.xlsx")
        
    # top level assign
    walk_in = Toplevel(master)
    
    # Adjust size
    walk_in.geometry("648x365")
    
    # Title
    walk_in.title('Walk In')
    walk_in.iconphoto(False, logo)
    
    # initialization
    canvas2 = Canvas(walk_in)
    
    # images
    image1 = canvas2.create_image(325, 180, image = bg)
    
    # Initializing Labels
    # name_var = StringVar()
    # clicked = StringVar()
    # location_var = StringVar()
    # email_var = StringVar()
    # phone_var = StringVar()
    # major_var = StringVar()
    # who_var = StringVar()
    # id_var = StringVar()
    # reason_var = StringVar()
    name_entry = Entry(canvas2, textvariable = name_var)
    email_entry = Entry(canvas2, textvariable = email_var)
    phone_entry = Entry(canvas2, textvariable = phone_var)
    major_entry = Entry(canvas2, textvariable = major_var)
    who_entry = Entry(canvas2, textvariable = who_var)
    id_entry = Entry(canvas2, textvariable = id_var)
    reason_entry = Text(canvas2, height = 7, width = 40, font = ("Times New Roman", 10))
    
    # tkinter buttons 
    sub_btn_2 = Button(canvas2, text = 'Submit', compound = LEFT, image = arrow, command = add_to_walk_in)
    
    # Text text2 = canvas2.create_text(10, 20, text = "", font = font3, anchor = "w", fill = "#1E4D2B")
    text1 = canvas2.create_text(190, 20, text = "Name:", font = font3, fill = "#1E4D2B", anchor = "e") 
    text2 = canvas2.create_text(190, 50, text = "Email:", font = font3, fill = "#1E4D2B", anchor = "e")
    text3 = canvas2.create_text(190, 80, text = "Phone #:", font = font3, fill = "#1E4D2B", anchor = "e")
    text4 = canvas2.create_text(190, 110, text = "Major:", font = font3, fill = "#1E4D2B", anchor = "e")
    text5 = canvas2.create_text(190, 145, text = "Student?(Yes/No):", font = font3, fill = "#1E4D2B", anchor = "e")
    text6 = canvas2.create_text(190, 175, text = "CSU ID:", font = font3, fill = "#1E4D2B", anchor = "e")
    text7 = canvas2.create_text(190, 290, text = "Reason for\nyour visit today:", font = font3, fill = "#1E4D2B", anchor = "e")
    text8 = canvas2.create_text(640, 290, text = "*By submitting this form\nyou agree that your data\nwill be collected", font = font4, anchor = "e", fill = "#1E4D2B")
    
    # Packing
    name_entry.grid(row = 0, column = 0, padx = (200, 0), pady = (10, 0))
    email_entry.grid(row = 1, column = 0, padx = (200, 0), pady = (10, 0))
    phone_entry.grid(row = 2, column = 0, padx = (200, 0), pady = (10, 0))
    major_entry.grid(row = 3, column = 0, padx = (200, 0), pady = (10, 0))
    who_entry.grid(row = 4, column = 0, padx = (200, 0), pady = (10, 0))
    id_entry.grid(row = 5, column = 0, padx = (200, 0), pady = (10, 0))
    reason_entry.place(x = 200, y = 200)
    sub_btn_2.place(x = 527, y = 322) # submit
    canvas2.pack(fill = BOTH, expand = YES)
    
    

# function to open a new window
# on a button click
def open_sign_in():
	
    # Image Display Command
    def display_image(choice):
        """gets the choice of the dropdown menu and displays the picture"""
        choice = clicked.get()
        counter = -1
        for person in options.keys():
            counter+=1
            if choice == person:
                canvas1.itemconfig(image2, image = photo_paths[counter])

	# Toplevel object which will
	# be treated as a new window
    sign_in = Toplevel(master)
    # Adjust size
    sign_in.geometry("648x365")

    # Add title
    sign_in.title("Sign In")
    sign_in.iconphoto(False, logo)
    
    # Create Canvas
    canvas1 = Canvas(sign_in)
    
    # Display image
    image1 = canvas1.create_image(325, 180, image = bg)
    image2 = canvas1.create_image(587, 260, image = blank_img)
    
    # Text
    text2 = canvas1.create_text(10, 45, text = "Please enter your name and select\nwho you have a meeting with\nfrom the dropdown menu", font = font2, anchor = "w", fill = "#1E4D2B") # name entry label
    text3 = canvas1.create_text(10, 155, text = "Please enter the location\nof the meeting if known", font = font2, anchor = "w", fill = "#1E4D2B") # name entry label
    text3 = canvas1.create_text(10, 300, text = "*By submitting this form you agree that your data will be collected", font = font4, anchor = "w", fill = "#1E4D2B")
   
    # Initializing Labels
    # selection_label = Label(textvariable = recipient_var)
    drop = OptionMenu(canvas1, clicked, list(options.keys())[0], *list(options.keys()), command = display_image)
    name_entry = Entry(canvas1, textvariable = name_var)
    location_entry = Entry(canvas1, textvariable = location_var)
    sub_btn_1 = Button(canvas1, text = 'Submit', compound = LEFT, image = arrow, command = submit1)

    # Packing Heirarchy
    #  padx=(padding_from_left_side, padding_from_right_side), 
    #  pady=(padding_from_top, padding_from_bottom))
    name_entry.grid(row = 0, column = 0, pady = (90,0), padx = (10, 0), ipadx = 100) # name entry text
    drop.grid(row = 0, column = 1, pady = (90,0), padx = (10, 0)) # dropdown
    location_entry.grid(row = 2, column = 0, pady = (75, 0), padx = (10, 0), ipadx = 100)
    sub_btn_1.place(x = 527, y = 322) # submit
    canvas1.pack(fill = BOTH, expand = YES) # background
    
# Email Alert Function
def email_alert(subject, body, to):
    """paramaters are (subject, body, to) and sets and sends an email message through gmail smtp"""
    msg = EmailMessage() # create a msg object
    msg.set_content(body) # set the content with the passed in body paramater
    msg['subject'] = subject # add a subject to the object
    msg['to'] = to # this is the recipient

    # This section may need to be updated
    user = "walter.alerts.asa@gmail.com" # current email adress used to send messages
    msg['from'] = user # same as above
    password = "gniskpmrasnqzgka" # Two-Factor auth APP PASSWORD

    server = smtplib.SMTP("smtp.gmail.com", 587) # gmail smtp server
    server.starttls() # use tls
    server.login(user, password) # login with the app password and the user email address
    server.send_message(msg) # send the message

    server.quit() # quit the gmail server

# Submit Button Command
def submit1():
    """submit button command from button"""
    recipient = options[clicked.get()]
    print("Email was sent to: " + recipient + " | from: " + name_var.get() + " |")
    add_to_sign_in()
    email_alert("Front Desk Alert", "{} is here to see you!".format(name_var.get()), recipient)
    location_var.set("")
    name_var.set("")

def esc_clicked():
    webbrowser.open("https://www.engr.colostate.edu/engineering-success-center/")

def web_map_clicked():
    webbrowser.open("https://map.concept3d.com/?id=748#!ct/53279,46630,20377,13646,13645,13644,9554")

def current_events_clicked():
    webbrowser.open("https://calendar.colostate.edu/")

def about_clicked():
    webbrowser.open(r"Z:\00 Front Desk Staff\Data\Welcome Kiosk\.docx files\About.docx")

# the following functions are for excel
def add_to_sign_in():
    sheet = wb['sign-in-reports']
    max_row = get_maximum_rows(sheet_object = sheet) + 1
    sheet['A' + str(max_row)] = name_var.get()
    sheet['B' + str(max_row)] = time.asctime(time.localtime(time.time()))
    sheet['C' + str(max_row)] = clicked.get()
    sheet['D' + str(max_row)] = location_var.get()
    wb.save(r"Z:\00 Front Desk Staff\Data\Welcome Kiosk\.xlsx files\kiosk-reports.xlsx")

def get_maximum_rows(*, sheet_object):
    rows = 0 # row counter
    for max_row, row in enumerate(sheet_object, 1): # run through each row in the sheet
        if not all(col.value is None for col in row): # check if there's no value in the row cells
            rows += 1
    return rows # finally return max row

# creates a Tk() object
master = Tk()

# sets the geometry of main window
master.geometry("200x200")

# master title
master.title("Welcome Kiosk Main")

# style

# fonts
font1 = ("Helvetica", 20, "bold") # heading 1
font2 = ("Helvetica", 16, "bold") # heading 2
font3 = ("Helvetica", 12, "bold") # heading 3
font4 = ("Helvetica", 8, "italic") # subscript

# tkinter vars
name_var = StringVar()
clicked = StringVar()
location_var = StringVar()
email_var = StringVar()
phone_var = StringVar()
major_var = StringVar()
who_var = StringVar()
id_var = StringVar()
reason_var = StringVar()

# Add logo
logo = PhotoImage(file = r"Z:\00 Front Desk Staff\Data\Welcome Kiosk\.png files\Ram Logo.png").subsample(50, 50)
master.iconphoto(False, logo)

# Add background
bg = PhotoImage(file = r"Z:\00 Front Desk Staff\Data\Welcome Kiosk\.png files\Background.png")

# Initalize staff pictures
blank_img = PhotoImage(file = r"Z:\00 Front Desk Staff\Data\Welcome Kiosk\.png files\White.png")

# Add image files
arrow = PhotoImage(file = r"Z:\00 Front Desk Staff\Data\Welcome Kiosk\.png files\Arrow.png").subsample(100, 100)
esc_logo = PhotoImage(file = r"Z:\00 Front Desk Staff\Data\Welcome Kiosk\.png files\Ram.png").subsample(5, 5)
map_logo = PhotoImage(file = r"Z:\00 Front Desk Staff\Data\Welcome Kiosk\.png files\Map.png").subsample(5, 5)
events_logo = PhotoImage(file = r"Z:\00 Front Desk Staff\Data\Welcome Kiosk\.png files\Events.png").subsample(5, 5)
about_logo = PhotoImage(file = r"Z:\00 Front Desk Staff\Data\Welcome Kiosk\.png files\Question.png").subsample(5, 5)
photo_paths = []
for person in options.keys():
    photo_paths.append(PhotoImage(file = r"Z:\00 Front Desk Staff\Data\Welcome Kiosk\.png files" + "\\" + person + ".png"))

# initializing 
btn1 = Button(master, text = "Sign In", command = open_sign_in)
btn2 = Button(master, text = "Walk In", command = open_walk_in)
esc_btn = Button(master, image = esc_logo, command = esc_clicked)
map_btn = Button(master, image = map_logo, command = web_map_clicked)
events_btn = Button(master, image = events_logo, command = current_events_clicked)
about_btn = Button(master, image = about_logo, command = about_clicked)

# packing
btn1.pack(pady = 20)
btn2.pack(pady = 20)
map_btn.place(x = 25, y = 150) # interactive map
events_btn.place(x = 65, y = 150) # events 
esc_btn.place(x = 105, y = 150)
about_btn.place(x = 145, y = 150) # about


# mainloop, runs infinitely
mainloop()
