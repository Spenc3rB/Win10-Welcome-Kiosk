import smtplib
from email.message import EmailMessage

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

if __name__ == '__main__':
    x = input("Enter your name: \n")
    email_alert("Front Desk Alert", "{} is here to see you!".format(x), "spencer@beerfamily.us")
