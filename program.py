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
messages = inbox.Items
messages.Sort('ReceivedTime')
length = len(messages)
print('number of emails in inbox: '+str(length))
print('last email from: ' + getSenderAddress(messages[length-1]))
print('Subject: ' + messages[length-1].Subject)
print('Content: '+ messages[length-1].Body) # messages[length-1].HTMLBody csak akkor mukodik ha HTML tartalmu emailt kaptunk
print('Date: '+ messages.GetLast().SentOn.strftime("%Y.%m.%d"))
print('Time: '+ messages[length-1].SentOn.strftime("%H:%M:%S"))


# Kiírás excelbe
def write_excel():
    Sender = getSenderAddress(messages[length-1])
    Cc = ""
    Subject = messages[length-1].Subject
    Date = messages.GetLast().SentOn.strftime("%Y.%m.%d")
    Time = messages[length-1].SentOn.strftime("%H:%M:%S")

    Recip = messages[length-1].Recipients
    for r in Recip:
        Cc = str(Cc) + str(r.AddressEntry)

    now =str(datetime.datetime.now().strftime("%d-%m-%Y_%H-%M"))

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


