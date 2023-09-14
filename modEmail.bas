Attribute VB_Name = "modEmail"
Option Compare Database

Sub EmailIF(myFile)
Dim mySubject, myBody, myPassword, myPath
Dim myrecp As String

mySubject = myFile
myrecp = CurrentDb.TableDefs("tblPASSWORD").OpenRecordset.Fields("EMAIL").Value
myBody = vbCr & vbCr & "Regards," & vbCr & "BAXTER" & vbCr & vbCr
myPath = CurrentProject.Path & "\TAX_INITIATIONS\" & myFile

'Email using Outlook

   Dim mOutlookApp As Outlook.Application
   Dim objMsg As MailItem

    Set mOutlookApp = GetObject("", "Outlook.application")
    Set objMsg = mOutlookApp.CreateItem(olMailItem)
   With objMsg
        .To = "Chanda.binwal@gds.ey.com" 'myrecp
        .Subject = mySubject
        .Attachments.Add myPath
        .Display
    End With
    'SendKeys "%(S)"
    Set objMsg = Nothing
End Sub
Sub EmailReport(myFile)
Dim mySubject, myBody, myPassword, myPath
Dim myrecp As String

mySubject = myFile
myrecp = CurrentDb.TableDefs("tblPASSWORD").OpenRecordset.Fields("REPORT_EMAIL").Value
myBody = vbCr & vbCr & "Regards," & vbCr & "BAXTER" & vbCr & vbCr
myPath = CurrentProject.Path & "\REPORTS\" & myFile
    'Email using Outlook
   Dim mOutlookApp As Outlook.Application
   Dim objMsg As MailItem

    Set mOutlookApp = GetObject("", "Outlook.application")
    Set objMsg = mOutlookApp.CreateItem(olMailItem)
   With objMsg
        .To = "Chanda.binwal@gds.ey.com" 'myrecp
        .Subject = mySubject
        .Attachments.Add myPath
        .Display
    End With
    'SendKeys "%(S)"
    Set objMsg = Nothing
End Sub
