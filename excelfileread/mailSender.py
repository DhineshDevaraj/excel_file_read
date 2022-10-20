import glob
import smtplib
from email.message import EmailMessage
import dataset
import pathlib


def sendmail():

    msg = EmailMessage()
    msg['Subject'] = 'Regarding Sample Data'
    msg['From'] = 'Data Team'
    msg['To'] = dataset.TO_ADDRESS

    with open('EmailTemplate.txt') as file:
        data = file.read()
        msg.set_content(data)

    for data_file in glob.glob("*.xlsx"):
        with open(data_file, "rb") as file:
            file_data = file.read()
            filename = file.name
            msg.add_attachment(file_data, maintype='application', subtype='xlsx', filename=filename)

    for data_file in glob.glob("*.png"):
        with open(data_file, "rb") as file:
            file_data = file.read()
            filename = file.name
            msg.add_attachment(file_data, maintype='application', subtype='png', filename=filename)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(dataset.USER_EMAIL, dataset.USER_PASSWORD)
        server.send_message(msg)

    print("Email Sent Successfully !!")


if __name__ == "__main__":
    sendmail()
