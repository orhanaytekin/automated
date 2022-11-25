import ssl
import smtplib
import email
from email.message import EmailMessage
import time
import pandas as pd
from main import Ui_MainWindow
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMainWindow, QApplication, QMessageBox
import sys

def createEmailMessage(body, subject, receiver, sender):
    """creates an EmailMessage object then returns it

    Args:
        body (string): contents of the email message
        subject (string): contents of the subject
        receiver (string): receiver
        sender (string): sender

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


def connectTo(EM, sender, password, receiver):
    """ Connect to the gmail server with the specified context then send it to the users

    Args:
        EM (EmailMessage): Email Message that is to be sent to the users email.
    """

    context = ssl.create_default_context()

    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as smtp:
        smtp.login(sender, password)
        smtp.sendmail(sender, receiver, EM.as_string())
        print("It is done")

class MyMainWindow(QMainWindow):
    def __init__(self):
        super(MyMainWindow, self).__init__()
        self.main_ui = Ui_MainWindow()
        self.main_ui.setupUi(self)
        
        self.setWindowTitle("Automated Mail Sender")
        self.select_button = self.main_ui.select_button
        self.send_mail = self.main_ui.send_mail
        self.is_visible = self.main_ui.is_visible
        
        self.passedit = self.main_ui.passedit
        
        self.file_name = self.main_ui.file_name
        self.file_path = ""
        self.file = ""
        
        self.is_visible.stateChanged.connect(self.change_visible)
        self.select_button.clicked.connect(self.open_file)
        self.send_mail.clicked.connect(self.save)
        
        
        
    def change_visible(self):
        if self.is_visible.isChecked:
            self.passedit.setEchoMode(QtWidgets.QLineEdit.Normal)
        else:
            self.passedit.setEchoMode(QtWidgets.QLineEdit.Password)
             
        
    def open_file(self):
        self.file_path = QFileDialog.getOpenFileName(caption = "select file", directory= "", filter= "*.xlsx",
                                                     initialFilter="")
        
        self.absolute_path = self.file_path[0]
        self.file = self.file_path[0].split("/")[-1]
        self.show_file()
    
    def show_file(self):
        file = self.file
        if file:
            self.file_name.setText(file)
        else:
            self.file_name.setText("I/O Error")
            
    def save(self):
        
        self.sender = self.main_ui.mailedit.text()
        self.password = self.main_ui.passedit.text()
        self.subject = self.main_ui.subject.text()
        self.body = self.main_ui.body.toPlainText()
        
        message = f"{self.sender} {self.subject}\n {self.body}"
        print(message)
        
        excel_file = pd.read_excel(window.absolute_path)

        for i in range(len(excel_file["ad"].tolist())):
            ad = excel_file["ad"].tolist()
            soyad = excel_file["soyad"].tolist() 
            mail = excel_file["mail"].tolist()

            receiver = mail[i]

            print(f"sending mail to {ad[i]} {soyad[i]}.")
            EM = createEmailMessage(window.body, window.subject, window.sender, receiver)
            connectTo(EM, window.sender, window.password, receiver)
            print("Sleeping")
            time.sleep(1) 
        
        

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    window = MyMainWindow()
    window.show()
    app.exit(app.exec_())
    