'''
Ez a program az elindulopont minden csapatnak.
Uncommenteljetek a ket reszbol az egyiket az egyeni teszteleshez
'''

# pip install pywin32 
# csak windows-on lehet dolgozni
import win32com.client as client

outlook = client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")

''' # send an email
message =outlook.CreateItem(0)
message.Display()
message.To = "attila.edmond.izsak@bce.bitclub.hu"
message.Subject = "Python testing"
message.HTMLBody = "Ezt a python programbol kuldtem. <b>wowo</b>"
message.Send()
'''

''' # read an email
def getSenderAddress(msg):
    if msg.Class == 43:
        if msg.SenderEmailType == "EX":
            return msg.Sender.GetExchangeUser().PrimarySmtpAddress
        else:
            return msg.SenderEmailAddress 


# access last email in inbox
# folder codes - https://learn.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
inbox = mapi.GetDefaultFolder(6)
messages = inbox.Items
length = len(messages)
print('number of emails in inbox: '+str(length))
print('last email from: ' + getSenderAddress(messages[length-1]))
print('Subject: ' + messages[length-1].Subject)
print('Content: '+ messages[length-1].Body) # messages[length-1].HTMLBody csak akkor mukodik ha HTML tartalmu emailt kaptunk
'''

