import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd


def mail(receivers,NAME,ORB):
    email = "aadityakharkiatest@gmail.com"
    password = "pxduwpgyfmtvtzmj"
    send_to_email = receivers
    subject = "IAYP Parent consent Form"
    message = f"Respected Parent,\n \nKindly find the parent consent form for {NAME} with ORB number {ORB}. \n\n You are request to fill the form and send it back hariomtripathi@welhamboys.org \n\nRegards,\nHariom Tripathi\n(Teacher)"
    file_location = "Parent_Consent_IAYP-.doc"

    msg = MIMEMultipart()
    msg["From"] = email
    msg["To"] = send_to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(message, "plain"))

    with open(file_location, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload((attachment).read())

    encoders.encode_base64(part)

    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {file_location}",
    )

    msg.attach(part)
    text = msg.as_string()

    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(email, password)
    server.sendmail(email, send_to_email, text)
    server.quit()

df = pd.read_csv('IAYP.csv',dtype={"S.NO":int,"NAME":object,"EMAIL":object,"ORB_NUMBER":object,"AWARD_TYPE":object})

for i in range(0,1):
    receivers = df["EMAIL"][i]
    NAME = df["NAME"][i]
    ORB = df["ORB_NUMBER"][i]
    mail(receivers,NAME,ORB)
    print(f"mail sent to {NAME} at {receivers}")