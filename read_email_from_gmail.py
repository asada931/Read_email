import smtplib
import time
import imaplib
import email
from selenium import webdriver
import re
from time import sleep





ORG_EMAIL   = "@gmail.com"
FROM_EMAIL  = "email" + ORG_EMAIL
FROM_PWD    = "Your password"
SMTP_SERVER = "imap.gmail.com"
#SMTP_PORT   = 993
# -------------------------------------------------
#
# Utility to read email from Gmail Using Python
#
# ------------------------------------------------

def get_email_body(msg):
    if msg.is_multipart():
        return get_email_body(msg.get_payload(0))
    else:
        return msg.get_payload(None,True)


def read_email_from_gmail():
    try:

        mail = imaplib.IMAP4_SSL(SMTP_SERVER)
        mail.login(FROM_EMAIL,FROM_PWD)
        latest_email_id=''

        while(True):
            mail.select('inbox',readonly=True)
            #to search inbox for specific email addresses 
            type, data = mail.search(None, '(FROM "Specific_Sender_Name")')
            mail_ids = data[0]
            id_list = mail_ids.split()
            # latest_email_id = id_list[-1]
            #print(latest_email_id)
            first_email_id = id_list[0]
            # latest_email_id = id_list[-1]
            # latest_email_id == mail_ids.split()[-1]
            # for latest_email_id in id_list:
            if (latest_email_id == id_list[-1]):
                print(latest_email_id)
                time.sleep(10)
                print('no new mail')
            else:
                latest_email_id = id_list[-1]
                typ, data = mail.fetch(id_list[-1], '(RFC822)')

                # for response_part in data:
                #   if isinstance(response_part, tuple):
                msg = email.message_from_string(data[0][1].decode('utf-8'))

                email_from = msg['from']
                email_subject = msg['subject']
                email_body = msg.get_payload()
                print('From : ' + email_from + '\n')
                print('Subject : ' + email_subject + '\n')
            #   
                body = get_email_body(msg).decode('utf_8')
                print(body)
                           
            continue
                    
                     
    except Exception as e:
     print(str(e))
printer()
read_email_from_gmail()
