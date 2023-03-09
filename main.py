from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
import os
import schedule
import pandas as pd
import datetime as dt
import time


#Creation of email class with attributes and methods
class Email:
    def __init__(self, subject):
        '''Initialization of the class and assigning properties'''
        self.subject = subject
        self.htmlbody = ""
        self.sender = "brainniest.week1@gmail.com"
        self.senderpass = "itdrojkvolmubstc"
        self.attachments = []

    def htmladd(self, html):
        '''add text to the body'''
        self.htmlbody = self.htmlbody + '<p></p>' + html

    def attach(self, msg):
        '''Attaching files to the message using MIMEMultipart class object'''
        for f in self.attachments:
            fp = open(f, "rb")
            attachment = MIMEBase('application', 'octate-stream')
            attachment.set_payload(fp.read())
            fp.close()
            encoders.encode_base64(attachment)
            attachment.add_header("Content-Disposition", "attachment", filename=f)
            attachment.add_header('Content-ID', '<{}>'.format(f))
            msg.attach(attachment)

    def addattach(self, files):
        '''Attaching as property to the Email class'''
        self.attachments = self.attachments + files


def logger(name, email, att):
    '''Create log of sent emails'''
    now = dt.datetime.now()
    files = os.listdir("../RosterLog/")
    if "log.csv" in files:
        with open("../RosterLog/log.csv", "a") as f:
            f.write(f"{now},{name},{email},{att}\n")
    else:
        with open("../RosterLog/log.csv", "w") as f:
            f.write(f"{now},{name},{email},{att}\n")


def get_emails():
    '''Reads emails.csv in RosterLog folder and gets Name and email'''
    email_dict = {}
    files = os.listdir(r"RosterLog/")
    if "emails.csv" in files:
        df = pd.read_csv("RosterLog/emails.csv")
        df["Fullname"] = df["Name"] + " " + df["Lastname"]
        for index, row in df.iterrows():
            email_dict[row["Fullname"]] = row["email"]
        return email_dict
    else:
        print("Unable to find the roster!")
        return None


def send(email, recipients):
    '''Sends emails to the read recipients'''
    for key in recipients:
        msg = MIMEMultipart()
        msg['From'] = email.sender
        msg['Subject'] = email.subject
        msg.attach(MIMEText(f"Dear {key},\n", 'html'))
        msg.attach(MIMEText("Please find attached daily reports", 'html'))
        if email.attachments:
            email.attach(msg)
        s = smtplib.SMTP_SSL("smtp.gmail.com")
        s.login(email.sender, email.senderpass)
        s.sendmail(email.sender, recipients[key], msg.as_string())
        logger(key, recipients[key], email.attachments)


def main_func():
    '''main function to be run by schedule library'''
    email_dict = get_emails()

    if email_dict is not None:
        list_of_files = [f for f in os.listdir(r"Attachments/") if os.path.isfile(os.path.join(r"Attachments/", f))]
        if len(list_of_files) == 0:
            print("This program cannot run without attachments! Please place attachments to the Attachments folder!")
            return
        else:
            os.chdir(r"Attachments/")
            mail = Email("Reports for " + datetime.now().strftime('%Y/%m/%d'))
            mail.addattach(list_of_files)
            send(mail, email_dict)
    else:
        print("This program will not work without roster. Please upload the roster to RosterLog folder")



schedule.every().day.at("10:21").do(main_func)
while True:
    schedule.run_pending()
    time.sleep(1)
