'''
Ez a program az elindulopont minden csapatnak.
Uncommenteljetek a ket reszbol az egyiket az egyeni teszteleshez
'''

# pip install pywin32 
# csak windows-on lehet dolgozni
import win32com.client as client
import xlsxwriter
import datetime

outlook = client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")

# send an email
# message =outlook.CreateItem(0)
# message.Display()
# message.To = "attila.edmond.izsak@bce.bitclub.hu"
# message.Subject = "Python testing"
# message.HTMLBody = "Ezt a python programbol kuldtem. <b>wowo</b>"
# message.Send()


# read an email
def getSenderAddress(msg):
    if msg.Class == 43:
        if msg.SenderEmailType == "EX":
            return msg.Sender.GetExchangeUser().PrimarySmtpAddress
        else:
            return msg.SenderEmailAddress 


# access last email in inbox
# folder codes - https://learn.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders



#Szükséges adatok kinyerése
inbox = mapi.GetDefaultFolder(6)
outbox = mapi.GetDefaultFolder(5)
outmessages = outbox.Items
messages = inbox.Items
messages.Sort('ReceivedTime')
length = len(messages)
# print('number of emails in inbox: '+str(length))
# print('last email from: ' + getSenderAddress(messages[length-1]))
# # print('Subject: ' + messages[length-1].Subject)
# print('Content: '+ messages[length-1].Body) # messages[length-1].HTMLBody csak akkor mukodik ha HTML tartalmu emailt kaptunk
# print('Date: '+ messages.GetLast().SentOn.strftime("%Y.%m.%d"))
# print('Time: '+ messages[length-1].SentOn.strftime("%H:%M:%S"))


cc1 = messages[length-1].Body
# cc2 = cc1.split("Cc: ")[1].split("Subject: ")[0]

import re
# kukacok = [m.start() for m in re.finditer('@', cc2)]

print(cc1)
# print(kukacok)

# for i in kukacok:


# Kiírás excelbe
def write_excel():
    Sender = getSenderAddress(messages[length-1])
    Cc = ""
    Subject = messages[length-1].Subject
    Date = messages.GetLast().SentOn.strftime("%Y.%m.%d")
    Time = messages[length-1].SentOn.strftime("%H:%M:%S")

    Recip = messages[length-1].Recipients
    for r in Recip:
        if r.AddressEntry.Type == "EX":
            Cc = str(Cc) + str(r.AddressEntry.GetExchangeUser().PrimarySmtpAddress) + "; "
        else:
            Cc = str(Cc) + str(r.AddressEntry.Address)+ "; "

    print(Recip)

    now =str(datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S"))

    workbook = xlsxwriter.Workbook('last_mail_'+now+'.xlsx')
    worksheet = workbook.add_worksheet()


    worksheet.write(1, 1, 'Küldő fél')
    worksheet.write(1, 2, 'Cc')
    worksheet.write(1, 3, 'Tárgy')
    worksheet.write(1, 4, 'Dátum')
    worksheet.write(1, 5, 'Idő')

    worksheet.write(2, 1, Sender)
    worksheet.write(2, 2, Cc)
    worksheet.write(2, 3, Subject)
    worksheet.write(2, 4, Date)
    worksheet.write(2, 5, Time)

    workbook.close()

# write_excel()


def write_txt_chat():
    # inbox elemek megszerzése
    salesMessagesIn = mapi.Folders("bit-bce-salesteam@bce.bitclub.hu").Folders(2).Items
    #elküldött elemek megszerzése (valamiért csak 3 elem van benne)
    salesMessagesOut = mapi.Folders("bit-bce-salesteam@bce.bitclub.hu").Folders(4).Items
    print(salesMessagesOut.Count)


    # bejövő CC-k kigyűjtése listába, hogy később ezek alapján tudjuk szűrni a maileket
    CcIn = []
    x = 0
    for i in salesMessagesIn:

        Recip = i.Recipients
        for r in Recip:
            if r.AddressEntry.Type == "EX":
                CcIn.append(str(r.AddressEntry.GetExchangeUser().PrimarySmtpAddress))
            else:
                CcIn.append(str(r.AddressEntry.Address))

        #gyorsabb futás érdekében csak az első 5 mail elemzése egyenlőre
        if x == 5:
            break
        x +=1


    CcIn = list(dict.fromkeys(CcIn))
    print(CcIn)

    # Elküldött CC-k kigyűjtése listába, hogy később ezek alapján tudjuk szűrni a maileket
    CcOut = []
    x = 0
    for i in salesMessagesOut:

        Recip = i.Recipients
        for r in Recip:
            if r.AddressEntry.Type == "EX":
                CcOut.append(str(r.AddressEntry.GetExchangeUser().PrimarySmtpAddress))
            else:
                CcOut.append(str(r.AddressEntry.Address))

        #gyorsabb futás érdekében csak az első 5 mail elemzése egyenlőre
        if x == 5:
            break
        x +=1

    # Duplikációk törlése    
    CcOut = list(dict.fromkeys(CcOut))
    print(CcOut)

    mergeCc = CcIn+CcOut
    mergeCc = list(dict.fromkeys(mergeCc))
    Partners = []
    
    #Partnerek kigyűjtése az alapján, hogy a mail címük nem bites
    for i in mergeCc:
        if  "bce.bitclub" not in i:
            Partners.append(i)

    print(Partners)


    #A mappa nevek és sorszámok kiírása (csak, hogy lássuk milyen mappák vannak)
    # for idx, folder in enumerate(mapi.Folders("bit-bce-salesteam@bce.bitclub.hu").Folders):
    #     print(idx+1, folder)



write_txt_chat()




