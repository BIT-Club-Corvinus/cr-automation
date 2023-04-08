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
    allpMail=""
    pMail=""
    pName=""
    sMail=""
    
    #Végig megyünk az aktuális üzenet CC-in
    for r in Recip:
        if (r.AddressEntry.Type == "EX"):
            pMail = str(r.AddressEntry.GetExchangeUser().PrimarySmtpAddress)
            
            #Ha a CC nem bites mailcím, akkor hozzá adjuk ahhoz a változóhoz amit beírunk majd a "CCs" oszlopba illetve a "Partner" oszlopba
            if ("bce.bitclub.hu" not in pMail):
                allpMail+=(pMail+", ")
                pName = pMail.split('@')[1]
                
        else:
            pMail = str(r.AddressEntry.Address)
            
            #Ha a CC nem bites mailcím, akkor hozzá adjuk ahhoz a változóhoz amit beírunk majd a "CCs" oszlopba illetve a "Partner" oszlopba
            if ("bce.bitclub.hu" not in pMail):
                allpMail+=(pMail+", ")
                pName = pMail.split('@')[1]
                
    try:
                 
        if m.SenderEmailType == "EX":
            sMail = m.Sender.GetExchangeUser().PrimarySmtpAddress
                    
        else:
            sMail = m.SenderEmailAddress
            
        #Ha a Sender nem bites mailcím, akkor hozzá adjuk ahhoz a változóhoz amit beírunk majd a "CCs" oszlopba illetve a "Partner" oszlopba
        if ("bce.bitclub.hu" not in sMail):
            allpMail+=(sMail+", ")
            pName = sMail.split('@')[1]
        
    except:
        print("Nem rendes mail.")
    
    worksheet.write(lc, 1, pName)        
    worksheet.write(lc, 2, allpMail)
    
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
    

alma = "alma"