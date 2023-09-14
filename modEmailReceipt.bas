Attribute VB_Name = "modEmailReceipt"
Option Compare Database

Sub EmailReceipt(myInitiator, myInitiation)
 
   Dim mOutlookApp As Outlook.Application
   Dim objMsg As MailItem

    Set mOutlookApp = GetObject("", "Outlook.application")
    Set objMsg = mOutlookApp.CreateItem(olMailItem)
   
    With objMsg
        .To = myInitiator
        .Subject = myInitiation
        .Display
    End With
   ' SendKeys "%(S)"

    Set objMsg = Nothing
End Sub
