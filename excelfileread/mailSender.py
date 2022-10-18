import glob
import smtplib
from email.message import EmailMessage


def sendmail():
    msg = EmailMessage()
    msg['Subject'] = 'Regarding Sample Data'
    msg['From'] = 'Data Team'
    msg['To'] = 'abc@gamil.com'

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
        server.login("email", "password")
        server.send_message(msg)

    print("Email Sent Successfully !!")


if __name__ == "__main__":
    sendmail()
