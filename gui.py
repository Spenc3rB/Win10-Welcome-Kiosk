import smtplib
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
	print(radio_var.get())
	email_alert("Front Desk Alert", "{} is here to see you!".format(name_var.get()), radio_var.get())
	name_var.set("")

# initialization and styling
root = Tk()
root.geometry("600x400")
root.title("Welcome Kiosk v0.1")
logo = PhotoImage(file = r"C:\Users\spenc\OneDrive\Documents\Notes - Python\Colorado_State_Rams_logo.svg.png")
logo_adjusted = logo.subsample(50, 50)
root.iconphoto(False, logo)
style = Style()
style.configure('TButton', font = ('calibri', 10, 'bold', 'underline'), foreground = '#1E4D2B')

# tkinter vars
name_var = StringVar()
radio_var = StringVar(root, "")
radios = {"Spencer" : "spencer@beerfamily.us",
        "Spencer via sms" : "3038566085@vtext.com",
        "Maya" : "mayap116@gmail.com",
        "Maya via sms" : "6506498277@txt.att.net"}


name_label = Label(root, text = 'Please enter your name', font=('calibre',10, 'bold'))
name_entry = Entry(root, textvariable = name_var, font=('calibre', 10, 'normal'))
for (text, value) in radios.items():
    Radiobutton(root, text = text, variable = radio_var,
        value = value).pack()

sub_btn = Button(root, text = 'Submit', image = logo_adjusted, compound = LEFT, style = 'TButton', command = submit)

name_label.pack()
name_entry.pack()
sub_btn.pack()

root.mainloop()
