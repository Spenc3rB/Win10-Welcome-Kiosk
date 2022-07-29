# How to set up guide for teams https://www.youtube.com/watch?v=yd1_AUYi6Fg
# Touch Screen https://www.amazon.com/Portable-Security-Touchscreen-1280x800-Raspberry/dp/B0B3DJY8Z1/ref=sr_1_4?crid=3QW6WB6Y3QAOK&keywords=touchscreen%2Bmonitor&qid=1656430782&refinements=p_36%3A-10000&rnid=386442011&s=electronics&sprefix=touchscreen%2Bmonitor%2Caps%2C95&sr=1-4&th=1
# Import module
import time
import os
import smtplib
import webbrowser
from openpyxl import load_workbook
from email.message import EmailMessage
from tkinter import *
from tkinter.ttk import *

# excel vars
wb = load_workbook('kiosk-reports.xlsx')

# Globals ---> initailization
# Key : Value dict variables
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
path = r"C:\Users\beersc\Desktop\gui_images_2\Karan Venayagamoorthy.png"

# Integers
max_len = -1

# Email Alert Function
def email_alert(subject, body, to):
    """paramaters are (subject, body, to) and sets and sends an email message through gmail smtp"""
    msg = EmailMessage()
    msg.set_content(body)
    msg['subject'] = subject
    msg['to'] = to

    user = "walter.alerts.asa@gmail.com"
    msg['from'] = user
    password = "gniskpmrasnqzgka"

    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(user, password)
    server.send_message(msg)

    server.quit()

# Submit Button Command
def submit1():
    """submit button command from button"""
    recipient = options[clicked.get()]
    print("Email was sent to: " + recipient + "from :" + name_var.get())
    add_to_sign_in()
    email_alert("Front Desk Alert", "{} is here to see you!".format(name_var.get()), recipient)
    name_var.set("")

#def submit2():
#    print("Walk in report: Name = " + name_var + 

# Image Display Command
def display_image(choice):
    """gets the choice of the dropdown menu and displays the picture"""
    choice = clicked.get()
    counter = -1
    for person in options.keys():
        counter+=1
        if choice == person:
            canvas1.itemconfig(image2, image = photo_paths[counter])

def google_clicked():
    webbrowser.open("https://www.google.com/")

def esc_clicked():
    webbrowser.open("https://www.engr.colostate.edu/engineering-success-center/")

def web_map_clicked():
    webbrowser.open("https://map.concept3d.com/?id=748#!ct/53279,46630,20377,13646,13645,13644,9554")

def current_events_clicked():
    webbrowser.open("https://calendar.colostate.edu/")

def about_clicked():
    webbrowser.open(r"Z:\00 Front Desk Staff\Personnel-Student Staff\Spencer\Welcome Kiosk\About.docx")

# the following functions are for excel
def add_to_sign_in():
    sheet = wb['sign-in-reports']
    max = 1
    for cell in sheet['A']:
        max+=1
    sheet['A' + str(max)] = name_var.get()
    sheet['B' + str(max)] = time.asctime(time.localtime(time.time()))
    sheet['C' + str(max)] = clicked.get()
    wb.save("kiosk-reports.xlsx")

for person in options.keys():
    if len(person) > max_len:
        max_len = len(person)

# Tkinter object initilization below
# Create object
root = Tk()

# tkinter vars
name_var = StringVar()
clicked = StringVar()

# Styling 
style = Style()
style.configure('TButton', font = ('calibri', 10, 'bold', 'underline'), foreground = 'black')

# Adjust size
root.geometry("960x1080")

# Add title
root.title('Welcome Kiosk v0.3')

# Add image files
arrow = PhotoImage(file = r"C:\Users\beersc\Desktop\gui_images_2\Arrow.png").subsample(100, 100)
google_logo = PhotoImage(file = r"C:\Users\beersc\Desktop\gui_images_2\Google.png")
esc_logo = PhotoImage(file = r"C:\Users\beersc\Desktop\gui_images_2\ESC Logo.png")
map_logo = PhotoImage(file = r"C:\Users\beersc\Desktop\gui_images_2\Map.png")
events_logo = PhotoImage(file = r"C:\Users\beersc\Desktop\gui_images_2\Events.png")
about_logo = PhotoImage(file = r"C:\Users\beersc\Desktop\gui_images_2\Question.png")
photo_paths = []
for person in options.keys():
    photo_paths.append(PhotoImage(file = r"C:\Users\beersc\Desktop\gui_images_2" + "\\" + person + ".png"))

# Add logo
logo = PhotoImage(file = r"Z:\00 Front Desk Staff\Logos\Ram Logo.png").subsample(50, 50)
root.iconphoto(False, logo)

# Add background
bg = PhotoImage(file = r"C:\Users\beersc\Desktop\gui_images_2\Background.png")

# Initalize staff pictures
blank_img = PhotoImage(file = r"C:\Users\beersc\Desktop\gui_images_2\White.png")

# Create Canvas
canvas1 = Canvas(root)

# Display image
image1 = canvas1.create_image(0, 540, image = bg)
image2 = canvas1.create_image(385, 265, image = blank_img)

# Add Text
text1 = canvas1.create_text(385, 70, text = "Welcome, Sign in here", fill = "white")
text2 = canvas1.create_text(385, 90, text = "Please enter your name:", fill = "white") # name entry label
text3 = canvas1.create_text(385, 130, text = "I have a meeting with (select from dropdown):", fill = "white") # dropdown menu label
text4 = canvas1.create_text(600, 900, text = "Hours of operation:\n8am - 4pm", font = ('Helvetica','50','bold'), fill = "white", justify = "right")

# Initializing Labels
# selection_label = Label(textvariable = recipient_var)
drop = OptionMenu(canvas1, clicked, list(options.keys())[0], *list(options.keys()), command = display_image)
name_entry = Entry(canvas1, textvariable = name_var)
google_button = Button(canvas1, image = google_logo, command = google_clicked)
esc_btn = Button(canvas1, image = esc_logo, command = esc_clicked)
map_btn = Button(canvas1, image = map_logo, command = web_map_clicked)
events_btn = Button(canvas1, image = events_logo, command = current_events_clicked)
about_btn = Button(canvas1, image = about_logo, command = about_clicked)
sub_btn_1 = Button(canvas1, text = 'Submit', compound = LEFT, image = arrow, command = submit1)

# Packing Heirarchy
#  padx=(padding_from_left_side, padding_from_right_side), 
#  pady=(padding_from_top, padding_from_bottom))
name_entry.grid(row = 0, column = 0, pady = (100,0), padx = (220,0), ipadx = 100) # name entry text
drop.grid(row = 1, column = 0, pady = (20,0), padx = (220,0)) # dropdown
sub_btn_1.grid(row = 2, column = 0, pady = (5,0), padx = (220,0)) # submit
map_btn.place(x = 800, y = 20) # interactive map
events_btn.place(x = 800, y = 140) # events 
google_button.place(x = 800, y = 260) # google
about_btn.place(x = 800, y = 380) # about
esc_btn.grid(row = 4, column = 0, pady = (120,0), padx = (220,0)) # our website
canvas1.pack(fill = BOTH, expand = YES) # background

# Execute tkinter
mainloop()