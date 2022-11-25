import ssl
import smtplib
import email
from email.message import EmailMessage
import time
import maskpass
import pandas as pd

# get the user inputs
print("Working on it ...")
time.sleep(0)
password = maskpass.askpass(prompt='Password: ', mask="*")
sender = input('Sender: ')
receiver = input('Receiver: ')
gmail = "@gmail.com"

sender = sender + gmail
receiver = receiver + gmail
subject = "Automated Message"


def createEmailMessage(body):
    """creates an EmailMessage object then returns it

    Args:
        body (string): contents of the email message

    Returns:
        EmailMessage: object to be returned
    """
    EM = EmailMessage()
    EM["subject"] = subject
    EM.set_content(body)
    EM["To"] = receiver
    EM["From"] = sender
    EM["Date"] = email.utils.formatdate(localtime=True)
    return EM


def connectTo(EM):
    """ Connect to the gmail server with the specified context then send it to the users

    Args:
        EM (EmailMessage): Email Message that is to be sent to the users email.
    """

    context = ssl.create_default_context()

    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as smtp:
        smtp.login(sender, password)
        smtp.sendmail(sender, receiver, EM.as_string())
        print("It is done")


if __name__ == "__main__":
    # EM = createEmailMessage()
    # connectTo(EM)

    # get mail from excel_file["dene.xlsx"]
    excel_file = pd.read_excel("dene.xlsx")

    number = int(input("How many emails do you want to send: "))

    for j in range(number):

        for i in range(len(excel_file["ad"].tolist())):
            ad = excel_file["ad"].tolist()
            soyad = excel_file["soyad"].tolist()
            mail = excel_file["mail"].tolist()

            body = f"""Sayın {ad[i]} {soyad[i]} bu mesaj otomatik olarak tarafınıza iletilmiştir. Lütfen yanıtlamayın.\n
            You don't need to do anything, just enjoy your ride. In the meantime we will keep you busy with these mails.
            """
            receiver = mail[i]

            print(f"sending mail to {ad[i]} {soyad[i]}.")
            EM = createEmailMessage(body)
            connectTo(EM)
            print("Sleeping, continue process")
