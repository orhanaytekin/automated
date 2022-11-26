import ssl
import smtplib
import email
from email.message import EmailMessage
import pandas as pd
from main import Ui_MainWindow
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMainWindow, QApplication, QMessageBox
from PyQt5.QtCore import QTimer
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


def connectTo(EM, sender, password, receiver, ad, soyad):
    """ Connect to the gmail server with the specified context then send it to the users

    Args:
        EM (EmailMessage): Email Message that is to be sent to the users email.
    """

    context = ssl.create_default_context()

    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as smtp:
        error = ""
        try:
            smtp.login(sender, password)
            smtp.sendmail(sender, receiver, EM.as_string())
            error = f"Mail sent to user {ad} {soyad}."
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText(error)
            msg.setWindowTitle("Mail Sent!")
            time_milliseconds = 1000
            QTimer.singleShot(time_milliseconds, lambda : msg.done(0))
            msg.exec_()
        except Exception as e:
            error = f"Unexpected {e=}, {type(e)=}"
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText(error)
            msg.setWindowTitle("Error")
            msg.exec_()
        
        return error

class MyMainWindow(QMainWindow):
  
    def __init__(self):
        """Init MyMainWindow Object 
        """
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
        self.err = ""
        
        
        
    def change_visible(self):
        """Changes the visibility of the password based on the check mark.
        """
        if self.is_visible.isChecked:
            self.passedit.setEchoMode(QtWidgets.QLineEdit.Normal)
        else:
            self.passedit.setEchoMode(QtWidgets.QLineEdit.Password)
    
    def error_message(self):
        """Error message displayed if the attribute is not empty."""
        if self.err:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText(self.err)
            msg.setWindowTitle("Error")
            msg.exec_()
             
        
    def open_file(self):
        """Open the file menu to select an excel file as specified.
        """
        self.file_path = QFileDialog.getOpenFileName(caption = "select file", directory= "", filter= "*.xlsx",
                                                     initialFilter="")
        
        self.absolute_path = self.file_path[0]
        self.file = self.file_path[0].split("/")[-1]
        self.excel_file = ""
        try:    
            self.excel_file = pd.read_excel(self.absolute_path)
        except Exception as e:
            msgn = f"Unexpected {e=}, {type(e)=}"
            # error_dialog = QtWidgets.QErrorMessage()
            # error_dialog.showMessage(msgn)
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText(msgn)
            msg.setWindowTitle("Error")
            msg.exec_()
        self.show_file()
    
    def show_file(self):
        """Show the file in the file_name attribute, if not selected then show 'File not selected!.'
        """
        file = self.file
        if file:
            self.file_name.setText(file)
        else:
            self.file_name.setText("File not selected!")
            
    def is_valid_mail(self):
        """Check if the mail is valid

        Returns:
            boolean: is_valid or not
        """
        if "." in self.sender and "@" in self.sender and " " not in self.sender:
            return True
        return False
                    
    def save(self):
        """Save the entered sender, subject, and body then send it to the users in the excel_file.
        """
        error_message = ""
        
        if self.file_name.text() == "File not selected!": 
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Please select a valid file.")
            msg.setWindowTitle("Error")
            msg.exec_()
            return
        
        self.sender = self.main_ui.mailedit.text().lower()
        if not self.is_valid_mail():
            error_message = "Sender is not a valid email address."
            self.err = error_message
            self.error_message()
            return
        
        self.password = self.main_ui.passedit.text()
        
        self.subject = self.main_ui.subject.text()
        self.body = self.main_ui.body.toPlainText()
        
        if self.subject.strip() == "" or self.body.strip() == "": 
            self.err = "Body or Subject of the mail should not be empty!"
            self.error_message()
            return
        
        message = f"{self.sender} {self.subject}\n {self.body}"
        print(message)

        for i in range(len(self.excel_file["ad"].tolist())):
            if not error_message.startswith("Unexpected"):
                ad = self.excel_file["ad"].tolist()
                soyad = self.excel_file["soyad"].tolist() 
                mail = self.excel_file["mail"].tolist()

                receiver = mail[i]

                print(f"sending mail to {ad[i]} {soyad[i]}.")
                EM = createEmailMessage(window.body, window.subject, window.sender, receiver)
                error_message = connectTo(EM, window.sender, window.password, receiver, ad[i], soyad[i])
                print("Sleeping")
        
        self.main_ui.subject.setText("")
        self.main_ui.body.setText("")
       

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    window = MyMainWindow()
    window.show()
    app.exit(app.exec_())
    