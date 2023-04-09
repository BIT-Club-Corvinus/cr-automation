from asyncio import sleep
import win32com.client as client
import xlsxwriter
import datetime
import os
import datetime
import time

outlook = client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")

# Bejövő mailek kinyerése
salesMessagesIn = mapi.Folders("bit-bce-salesteam@bce.bitclub.hu").Folders(2).Items
salesMessagesIn.Sort('ReceivedTime', True)


#Mai dátum "létrehozása" az excel nevéhez
nowdate = datetime.datetime.now()
nowdate = nowdate.strftime("%d_%m_%Y_%Hh%Mm%Ss")
nowdate = str(nowdate)

#Új munkafüzet létrehozása
workbook = xlsxwriter.Workbook('Partner_Communication_'+nowdate+'.xlsx')
worksheet = workbook.add_worksheet()

#Címsor feltöltése
worksheet.write(1, 1, 'Partner')
worksheet.write(1, 2, 'CCs')
worksheet.write(1, 3, 'Date')
worksheet.write(1, 4, 'Time')
worksheet.write(1, 5, 'Subject')
worksheet.write(1, 6, 'Days since last mail')
worksheet.write(1, 7, 'State')
worksheet.write(1,8, 'Contact')

#lc = line Cunter az excel fájl sorain léptetéshez
lc = 2

#Végig megyünk a bejövő maileken
for m in salesMessagesIn:
    
    #Limit 200 sorban, mivel a régebbi üzenetek már úgy sem relevánsak
    if lc == 200:
        break
    
    Recip = m.Recipients
    mBody = m.Body
    
    #PMail = Partner email címe
    #pName = Partnercég neve
    #sMail = Sender Mail a küldő emailcíme
    #allpMail = Minden NEM bites, YBG, vagy ubc cím ide kerül, ezt írjuk a CCs oszlopba
    #bitMail = bites kapcsolattartó neve
    #bitMails = ide gyűjti a bites címzetteket
    allpMail=""
    pMail=""
    pName=""
    sMail=""
    bitMail=""
    bitMails=[]
    
    #Sender vizsgálata
    try:
                 
        if m.SenderEmailType == "EX":
            sMail = m.Sender.GetExchangeUser().PrimarySmtpAddress
                    
        else:
            sMail = m.SenderEmailAddress
            
        #Ha a Sender nem bites mailcím, akkor hozzá adjuk ahhoz a változóhoz amit beírunk majd a "CCs" oszlopba illetve a "Partner" oszlopba
        if ("bce.bitclub.hu" not in sMail):
            allpMail+=(sMail+", ")
            pName = sMail.split('@')[1]
            if (pName  == "gmail.com"):
                pName = sMail.split('@')[0]
        #Ha a Sender bites mailcím, akkor hozzá adjuk ahhoz a változóhoz amit beírunk majd a "Contact" oszlopba
        else:
            bitMail=sMail.split('@')[0]
        
    except:
        print("Nem rendes mail.")

    #Végigmegyünk az aktuális üzenet CC-in
    for r in Recip:
        
        if (r.AddressEntry.Type == "EX"):
            pMail = str(r.AddressEntry.GetExchangeUser().PrimarySmtpAddress) 
        
        else:
            pMail = str(r.AddressEntry.Address)
            
        #Ha a CC nem bites mailcím, akkor hozzáadjuk ahhoz a változóhoz amit beírunk majd a "CCs" oszlopba illetve a "Partner" oszlopba
        if ("bce.bitclub.hu" not in pMail):
            allpMail+=(pMail+", ")
            pName = pMail.split('@')[1]
            if (pName  == "gmail.com"):
                pName = pMail.split('@')[0]
        
        #Ha a CC bites és nem bites küldte, akkor a címzetteket összegyűjti a bitMails változóba
        elif("bce.bitclub.hu" not in sMail):
            bitMails.append(pMail.split('@')[0].lower())

        #kiválasztja a bites címzettek közül(bitMails elemei) a kapcsolattartót(bitMail)
        if "bce-bitclub-salesteam" in bitMails:
                bitMails.delete(bitMails.index("bce-bitclub-salesteam"))
        for i in bitMails:
            if not(i=="lili.torok") and not(i=="gergo.komezei"):
                bitMail=i
                break
        if bitMail=="":
            for i in bitMails:
                bitMail+=(i+', ')
        
  
    worksheet.write(lc, 1, pName)        
    worksheet.write(lc, 2, allpMail)
    worksheet.write(lc, 8, bitMail)
    
    #Dátum és idő kinyerése
    mDateTime = str(m.ReceivedTime)
    mDateTime = mDateTime.split(".")[0]
    
    mDate = mDateTime.split(" ")[0]
    mTime = mDateTime.split(" ")[1]
    
    worksheet.write(lc, 3, mDate)
    worksheet.write(lc, 4, mTime)
    
    #Subject kinyerése
    mSubject = m.Subject
    worksheet.write(lc, 5, mSubject)
    
    #Days since last mail
    mDaysPasd = datetime.datetime.now() - m.ReceivedTime.replace(tzinfo=None)
    worksheet.write(lc, 6, mDaysPasd)
    
    #State. Minden üzenet mellé 1-es ha a bit volt a küldő, 0 ha bárki más.
    # Így ha szűrjük a táblázatot subjectre vagy partnerre akkor látszik, hogy az utolsó üzenet milyen típusú volt.
    
    if ("bce.bitclub.hu" in sMail):
        state = 1
    else:
        state= 0

        
    worksheet.write(lc, 7, state)   
    
    lc +=1
    
    
workbook.close()
