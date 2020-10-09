import smtplib
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import encoders
import os
import pandas as pd 


def prepare_mail(name,send_from,subject, message, files=[]):
    """Compose and send email with provided info and attachments.

    Args:
        send_from (str): from name
        send_to (list[str]): to name(s)
        subject (str): message title
        message (str): message body
        files (list[str]): list of file paths to be attached to email
        server (str): mail server host name
        port (int): port number
        username (str): server auth username
        password (str): server auth password
        use_tls (bool): use TLS mode
    """
    msg = MIMEMultipart()
    msg['From'] = send_from
    # msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(message.format(name),'plain'))

    for path in files:
        part = MIMEBase('application', "octet-stream")
        with open(path, 'rb') as file:
            part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        'attachment; filename="{}"'.format(Path(path).name))
        msg.attach(part)

    return msg




files_dir="./email_files"
files=[]
for filename in os.listdir(files_dir):
    print(filename)
    files.append(os.path.join(files_dir, filename))



mail_file=open("./message.txt","r")
mail_message=mail_file.read()


my_email = "sharmanshuma@gmail.com"
password = "Anshuma0301."

subject="Feature- ‘Decorating your home in a post-COVID world–the antimicrobial materials you need'"

server="smtp.gmail.com"
port=587

smtp = smtplib.SMTP(server, port)
smtp.starttls()
    
smtp.ehlo() 
smtp.login(my_email, password)

email_list = pd.read_excel('contacts/Stylam_.xlsx')
print(email_list.columns)
names=email_list[' Journalist ']
emails=email_list[' Email ']



for i in range(len(emails)): 
    name = names[i]
    email = emails[i]
    # for every record get the name and the email addresses
    msg=prepare_mail(name,my_email,subject,mail_message,files)

    # print(msg.as_string())
    print("sending to :"+ email)
    msg['TO']=email
  
    # the message to be emailed 
    message = msg.as_string()
  
    # sending the email 
    smtp.sendmail(my_email, email, message)
  
# close the smtp server 
# smtp.close()  
smtp.quit()


    
