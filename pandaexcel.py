from email.mime.nonmultipart import MIMENonMultipart
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
from email.mime.application import MIMEApplication

names=email_list['HolidayPartner']
emails=email_list['Email']
outer=MIMENonMultipart()

for i in range(len(email)):
    name=names[i]
    email=emails[i]
    #the message to be emailed
    message="Hello "+name
