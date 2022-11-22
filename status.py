# -*- coding: utf-8 -*-
"""
Created on Fri Nov 18 15:47:01 2022

@author: bpank
"""

import win32com.client as client
import xlsxwriter
import datetime
import pandas as pd
import numpy as np

outlook = client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")

# read an email
def getSenderAddress(msg):
    if msg.Class == 43:
        if msg.SenderEmailType == "EX":
            return msg.Sender.GetExchangeUser().PrimarySmtpAddress
        else:
            return msg.SenderEmailAddress 

# folder codes - https://learn.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders

df = pd.DataFrame(columns = ['partner', 'subject', 'content', 'date', 'time'])

#Folder inbox
inbox = mapi.GetDefaultFolder(6)
inbox_messages = inbox.Items
inbox_messages.Sort('ReceivedTime')
inbox_length = len(inbox_messages)

for i in range(inbox_length):
    try:
        sender = getSenderAddress(inbox_messages[i])
    except:
        sender = None
    subject = inbox_messages[i].Subject
    content = inbox_messages[i].Body
    try:
        date = inbox_messages[i].SentOn.strftime("%Y.%m.%d")
        time = inbox_messages[i].SentOn.strftime("%H:%M:%S")
    except:
        date = ""
        time = ""
    email = {'partner':sender, 'subject':subject, 'content':content, 'date':date, 'time':time}
    df = df.append(email, ignore_index = True)
    
#Sent email inbox
sent = mapi.GetDefaultFolder(5)
sent_messages = sent.Items
sent_messages.Sort('ReceivedTime')
sent_length = len(sent_messages)

for i in range(sent_length):
    recip = ""
    for recipents in sent_messages[i].Recipients:
        recip = recip + ' ' + recipents.Address
    subject = sent_messages[i].Subject
    content = sent_messages[i].Body
    email = {'partner':recip, 'subject':subject, 'content':content}
    try:
        date = sent_messages[i].SentOn.strftime("%Y.%m.%d")
        time = sent_messages[i].SentOn.strftime("%H:%M:%S")
    except:
        date = ""
        time = ""
    email = {'partner':sender, 'subject':subject, 'content':content, 'date':date, 'time':time}
    df = df.append(email, ignore_index = True)

#Filter Teams invitations
df.dropna(subset=['partner'], inplace = True)



#Filter other Teams and Akrivis messages
df = df.loc[~df['partner'].str.contains("microsoft") & ~df['partner'].str.contains("akrivis")]

df.info()

df_pivot = pd.pivot_table(df,index=['partner'], values = ['date','time'], aggfunc = np.max)
df_pivot['subject_last_mail'] = df_pivot.merge(df,how='left',on='partner')['subject']


# print('number of emails in inbox: '+str(length))
# print('last email from: ' + getSenderAddress(messages[length-1]))
# print('Subject: ' + messages[length-1].Subject)
# print('Content: ' + messages[length-1].Body)
# print('Date: '+ messages.GetLast().SentOn.strftime("%Y.%m.%d"))
# print('Time: '+ messages[length-1].SentOn.strftime("%H:%M:%S"))



