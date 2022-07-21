import win32com.client as win32
olApp = win32.Dispatch("Outlook.Application")
olNS = olApp.GetNameSpace("MAPI")

emails =["ndjpro@gmail.com","nschristie@me.com","scmd@me.com"]


for email in emails:
    mailItem = olApp.CreateItem(0) #this is the value of a mail object in documentation
    mailItem.Subject = "Placeholder title!"
    mailItem.BodyFormat = 1
    mailItem.Body ="""
    Yo!

    Placeholder text!

    Best,
    Name
    Title
    

    """
    mailItem.To = email 
    mailItem._oleobj_.Invoke(*(64209,0,8,0,olNS.Accounts.Item("placeholder@outlook.com")))

    # mailItem.BodyFormat = 2
    # mailItem.HTMLBody = "<HTML Markup>"

    mailItem.Display()
    #wow this works

    mailItem.Send()
