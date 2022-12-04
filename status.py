# -*- coding: utf-8 -*-
"""
Created on Fri Nov 18 15:47:01 2022

@author: bpank
"""

import win32com.client as client
import xlsxwriter
from datetime import datetime
from datetime import date
import pandas as pd
import numpy as np

outlook = client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")
df = pd.DataFrame(columns = ['partner','recip','subject', 'date','státusz'])
current_date = datetime.now()

# read an email
def getSenderAddress(msg):
    if msg.Class == 43:
        if msg.SenderEmailType == "EX":
            return msg.Sender.GetExchangeUser().PrimarySmtpAddress
        else:
            return msg.SenderEmailAddress

def WriteExcel():
    try:
        now =str(datetime.now().strftime("%Y-%m-%d_%H-%M-%S"))
        df_pivot.to_excel("Status_"+now+".xlsx", index = False)
        print("Writing done")
    except BaseException as ex:
        print(f'Error in writing Excel output: {ex}, {type(ex)}')
        return
    
# folder codes - https://learn.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders

print(pd. __version__)

#Get folder inbox
salesMessagesIn = mapi.Folders("bit-bce-salesteam@bce.bitclub.hu").Folders(2)

inbox_messages = salesMessagesIn.Items
inbox_messages.Sort('ReceivedTime', Descending = True)
inbox_length = len(inbox_messages)

statusz = {}

type(inbox_messages)


for i in range(inbox_length):
    try:
        sender = getSenderAddress(inbox_messages[i])
        print(sender)
    except:
        print("Semmi")
        sender = None
    bitvalaszolt = False
    try:
        if "bitclub.hu" in sender:
            searched_email = inbox_messages[i].Recipients[i]
            print("BIT a küldő")
            if searched_email.AddressEntry.Type == "EX":
                searched_email = str(searched_email.AddressEntry.GetExchangeUser().PrimarySmtpAddress)
            else:
                searched_email = str(searched_email.AddressEntry.Address)
            if searched_email not in statusz:
                print("Még nem merült fel")
                statusz[searched_email] = "Bit válaszolt"
                bitvalaszolt = True
            else:
                print("Már volt ilyen email")
                pass
        else:
            print("Partner a küldő")
            searched_email = inbox_messages[i].Recipients[i]
            if sender not in statusz:
                print("Még nem volt ilyen partner")
                statusz[sender] = "Partner válaszolt" 
                bitvalaszolt = False
            else:
                pass
    except:
        pass
    subject = inbox_messages[i].Subject
    try:
        date = inbox_messages[i].SentOn.strftime("%Y.%m.%d")
    except:
        date = ""
    email = {'partner':sender, 'recip':searched_email, 'subject':subject, 'date':date, 'státusz': bitvalaszolt}
    df_email_temp = pd.DataFrame([email])
    df = pd.concat([df,df_email_temp], ignore_index = True)
    
statusz

#Filter Teams invitations
df.dropna(subset=['partner'], inplace = True)
#☺ & ~df['partner'].str.contains("bitclub")
#Filter other Teams and Akrivis messages
df = df.loc[~df['partner'].str.contains("teams.microsoft")]

df_pivot = pd.pivot_table(df,index=['partner'], values = ['date','subject','státusz'], aggfunc = np.max)
df_pivot = df_pivot.reset_index()
df_pivot['days_since_last_mail'] = (current_date - pd.to_datetime(df_pivot['date'])).dt.days
df_pivot = df_pivot.reindex(columns = ["partner","date","subject","days_since_last_mail","státusz"])


# df_pivot.info()
# df.info()

#WriteExcel()




# print('number of emails in inbox: '+str(length))
# print('last email from: ' + getSenderAddress(messages[length-1]))
# print('Subject: ' + messages[length-1].Subject)
# print('Content: ' + messages[length-1].Body)
# print('Date: '+ messages.GetLast().SentOn.strftime("%Y.%m.%d"))
# print('Time: '+ messages[length-1].SentOn.strftime("%H:%M:%S"))



