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
df = pd.DataFrame(columns = ['partner','subject', 'date','time'])
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


#Get folder inbox
salesMessagesIn = mapi.Folders("bit-bce-salesteam@bce.bitclub.hu").Folders(2)

inbox_messages = salesMessagesIn.Items
inbox_messages.Sort('ReceivedTime')
inbox_length = len(inbox_messages)

for i in range(inbox_length):
    try:
        sender = getSenderAddress(inbox_messages[i])
    except:
        sender = None
    subject = inbox_messages[i].Subject
    try:
        date = inbox_messages[i].SentOn.strftime("%Y.%m.%d")
        time = inbox_messages[i].SentOn.strftime("%H:%M:%S")
    except:
        date = ""
        time = ""
    email = {'partner':sender, 'subject':subject, 'date':date, 'time':time}
    df_email_temp = pd.DataFrame([email])
    df = pd.concat([df,df_email_temp], ignore_index = True)
    


#Filter Teams invitations
df.dropna(subset=['partner'], inplace = True)

#Filter other Teams and Akrivis messages
df = df.loc[~df['partner'].str.contains("teams.microsoft") & ~df['partner'].str.contains("akrivis")]

df_pivot = pd.pivot_table(df,index=['partner'], values = ['date','time','subject'], aggfunc = np.max)
df_pivot = df_pivot.reset_index()
df_pivot['days_since_last_mail'] = (current_date - pd.to_datetime(df_pivot['date'])).dt.days
df_pivot = df_pivot.reindex(columns = ["partner","date","time","subject","days_since_last_mail"])

# df_pivot.info()
# df.info()

WriteExcel()

#To do: a bit-sales email inboxára szűrve scraping
#Kiírni Excelbe: distincten partnemailek, utolsó email subject, utolsó email dátum - idő, utolsó email óta eltelt napok száma, STÁTUSZ (lényeg)
#STÁTUSZ lehetséges értékek: partner válaszolt, bit válaszolt


# print('number of emails in inbox: '+str(length))
# print('last email from: ' + getSenderAddress(messages[length-1]))
# print('Subject: ' + messages[length-1].Subject)
# print('Content: ' + messages[length-1].Body)
# print('Date: '+ messages.GetLast().SentOn.strftime("%Y.%m.%d"))
# print('Time: '+ messages[length-1].SentOn.strftime("%H:%M:%S"))



