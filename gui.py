import smtplib
import time
from email.message import EmailMessage
from tkinter import *
from tkinter.ttk import *


# globals ------------------------------------------------------
recipient = ""
path = r"C:\Users\beersc\Desktop\gui_images_2\Karan Venayagamoorthy.png"
# --------------------------------------------------------------

def email_alert(subject, body, to):
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

def submit():
    print(clicked.get())
    recipient = options[clicked.get()]
    print(recipient)
    email_alert("Front Desk Alert", "{} is here to see you!".format(name_var.get()), recipient)
    name_var.set("")
   
def display_image(choice):
    choice = clicked.get()
    counter = -1
    for person in options.keys():
        counter+=1
        if choice == person:
            recipient_var.set("You've Selected {}".format(person))
            print(person)
            staff.configure(image = photo_paths[counter])
            

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
root.geometry("600x600")
root.title("Welcome Kiosk v0.2")

# images 
logo = PhotoImage(file = r"Z:\00 Front Desk Staff\Logos\Ram Logo.png").subsample(50, 50)
root.iconphoto(False, logo)
photo_paths = []
for person in options.keys():
    photo_paths.append(PhotoImage(file = r"C:\Users\beersc\Desktop\gui_images_2" + "\\" + person + ".png"))
    
print(path)
# styling
style = Style()
style.configure('TButton', font = ('calibri', 10, 'bold', 'underline'), foreground = '#1E4D2B')

# tkinter vars
name_var = StringVar()
clicked = StringVar()
recipient_var = StringVar()
recipient_var.set("You've Selected None")

# tkinter labels / buttons
background_image = PhotoImage(file = r"C:\Users\beersc\Desktop\gui_images_2\Background.png")
background = Label(image = background_image)
staff = Label(image = photo_paths[0])
selection_label = Label(textvariable = recipient_var)
drop = OptionMenu(root, clicked, list(options.keys())[0], *list(options.keys()), command = display_image)
name_label = Label(root, text = 'Please enter your name', font=('calibre',10, 'bold'))
other_label = Label(root, text = "Who is the person you wish to see?", font=('calibre',10, 'bold'))
name_entry = Entry(root, textvariable = name_var, font=('calibre', 10, 'normal'))
sub_btn = Button(root, text = 'Submit', image = logo, compound = LEFT, style = 'TButton', command = submit)

# packing to screen
background.place(x=0,y=0)
name_label.pack()
name_entry.pack()
other_label.pack()
drop.pack()
selection_label.pack()
staff.pack()
sub_btn.pack()
root.mainloop()
