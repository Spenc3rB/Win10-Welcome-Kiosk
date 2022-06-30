import smtplib
import time
from email.message import EmailMessage
from tkinter import *
from tkinter.ttk import *


# globals ------------------------------------------------------
recipient = ""
# --------------------------------------------------------------

def email_alert(subject, body, to):
    msg = EmailMessage()
    msg.set_content(body)
    msg['subject'] = subject
    msg['to'] = to

    user = "esc.email.alert@gmail.com"
    msg['from'] = user
    password = "dbgtrcyascggumry"

    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(user, password)
    server.send_message(msg)

    server.quit()

def submit():
    print(clicked.get())
    recipient = options[clicked.get()]
    print(recipient)
    email_alert("Front Desk Alert", "{} is here to see you!".format(name_var.get()), recipient)
    name_var.set("")
   
def display_image(choice):
    choice = clicked.get()
    for person in options.keys():
        if choice == person:
            recipient_var.set("You've Selected {}".format(person))
            path = r"C:\Users\beersc\Desktop\gui_images" + "\\" + person + ".png"
            print(path)
            staff_image = PhotoImage(file = r"C:\Users\beersc\Desktop\gui_images" + "\\" + person + ".png").subsample(50,50)
            Label(image = staff_image, text= "Hello?").pack()
            
     
    

# key : value pairs 
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

# initialization and styling
root = Tk()
root.geometry("600x400")
root.title("Welcome Kiosk v0.2")
logo = PhotoImage(file = r"Z:\00 Front Desk Staff\Logos\Ram Logo.png")
logo_adjusted = logo.subsample(50, 50)
root.iconphoto(False, logo)
style = Style()
style.configure('TButton', font = ('calibri', 10, 'bold', 'underline'), foreground = '#1E4D2B')

# tkinter vars
name_var = StringVar()
clicked = StringVar()
recipient_var = StringVar()
recipient_var.set("You've Selected None")

# tkinter labels / buttons
selection_label = Label(textvariable = recipient_var)
drop = OptionMenu(root, clicked, list(options.keys())[0], *list(options.keys()), command = display_image)
name_label = Label(root, text = 'Please enter your name', font=('calibre',10, 'bold'))
other_label = Label(root, text = "Who is the person you wish to see?", font=('calibre',10, 'bold'))
name_entry = Entry(root, textvariable = name_var, font=('calibre', 10, 'normal'))
sub_btn = Button(root, text = 'Submit', image = logo_adjusted, compound = LEFT, style = 'TButton', command = submit)

# packing to screen
selection_label.place(x = 200, y = 200)
name_label.place(x = 300, y = 0)
name_entry.place(x = 300, y = 22)
other_label.place(x = 0, y = 0)
drop.place(x = 70, y = 20)
sub_btn.place(x = 315, y = 44)
root.mainloop()
