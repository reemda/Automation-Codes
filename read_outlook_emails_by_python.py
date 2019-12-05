# import libraries
import pandas as pd 
import numpy as np

import win32com.client
import win32com
import time
import datetime as dt
import unicodecsv as csv
from os import path

# folders function to print the name of the folders of outlook account 
def folders(fold):
    print([folder.Name for folder in fold])

# Outlook account name
account_name = 'Outlook Email Name'  

# Run win32com.client
outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
account = outlook.GetNamespace("MAPI").Folders[account_name]
messages=account.Items

# Read the name of folders in outlook account
folders(account.Folders)

# choose inbox folder
inbox =account.Folders["Inbox"]

# Read the name of the folders in inbox
[folder.Name for folder in inbox.Folders]

# Print how many emails found in inbox
messages = [message for message in sorted(inbox.Items, reverse=True, key=lambda msg: msg.ReceivedTime)]
print('{} messages found for {}'.format(len(messages), account_name))

# Create list to append the emails to 
emails = []
for message in messages:
    if ~hasattr(message, 'To'):
        to = ''
    else:
        to = message.To
    emails.append({'Received Time': message.ReceivedTime.strftime("%d/%m/%y %I:%M %p"), 'Sender': message.SenderName, 'To': to, 'Subject': message.Subject, 'Body': message.Body})
    

# Create pandas dataframe
emails_df=pd.DataFrame(emails)
emails_df.head()
