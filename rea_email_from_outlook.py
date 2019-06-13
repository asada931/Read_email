import smtplib
import time
import pyOutlook
import imaplib
import email
import subprocess
import re
from time import sleep        		
from exchangelib import Credentials, Account
import win32com.client

#Outlook Folder Numbers
# 3  Deleted Items
# 4  Outbox
# 5  Sent Items
# 6  Inbox
# 9  Calendar
# 10 Contacts
# 11 Journal
# 12 Notes
# 13 Tasks
# 14 Drafts
# om email.utils import parseaddr
# -------------------------------------------------
#
# Utility to read email from Outlook Client Using Python
#
# ------------------------------------------------

#Method to get all display all accounts in your Outlook Client

def get_accounts():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    accounts= win32com.client.Dispatch("Outlook.Application").Session.Accounts;
    for account in accounts:
        global inbox
        inbox = outlook.Folders(account.DeliveryStore.DisplayName)
        print("****Account Name**********************************")
        print(account.DisplayName)
        print("***************************************************")
        folders = inbox.Folders

#Method to read email from inbox of default account

def read_email_from_outlookClient():
    try:
        latest_email = ''    
        while(True):
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.GetDefaultFolder(6)
            messages = inbox.Items
            mss=messages.GetLast()
            body = mss.body
            arr = body.split()
            if latest_email == mss.body:
                print('no new mail')
                sleep(5)
            else:
                latest_email = mss.body
                #Here printing the data of latest email
                print(body)
                print(mss.SenderEmailAddress)
                print(mss.subject)
                
    except Exception as e:
     print(str(e))
read_email_from_outlookClient()


# -------------------------------------------------
#
# additional method to read email from second account in Outlook Client 

# ------------------------------------------------
#           account_two = outlook.Folders("xyz@abc.com")
            # Inbox = account_two.Folders("Inbox")
            # msg=Inbox.Items
            # msgs=msg.GetLast()
            # print(msgs)

                    
