Attribute VB_Name = "modExport"
Option Compare Database

Sub ExportIF()

Dim myPath, myFileName, myTemplate, myProgram, myLast, myFirst, myEMPID, myDivision, myHome, myHost, myWbk
Dim xlApp As Excel.Application
Dim initEmail As String, initName As String
Dim FS As Object
Dim Backup_Str As String

Set FS = CreateObject("Scripting.FileSystemObject")

myProgram = CurrentDb.TableDefs("rptTAX").OpenRecordset.Fields("FirstOfField8").Value
myLast = CurrentDb.TableDefs("rptTAX").OpenRecordset.Fields("FirstOfField2").Value
myFirst = CurrentDb.TableDefs("rptTAX").OpenRecordset.Fields("FirstOfField3").Value
myEMPID = CurrentDb.TableDefs("rptTAX").OpenRecordset.Fields("FirstOfField4").Value
myDivision = CurrentDb.TableDefs("rptTAX").OpenRecordset.Fields("FirstOfField10").Value
myHome = CurrentDb.TableDefs("rptTAX").OpenRecordset.Fields("FirstOfField17").Value
myHost = CurrentDb.TableDefs("rptTAX").OpenRecordset.Fields("FirstOfField18").Value
initEmail = CurrentDb.TableDefs("rptTAX").OpenRecordset.Fields("FirstOfField16").Value

myPath = CurrentProject.Path
myTemplate = myPath & "\EYTAX_IF_TEMPLATE.xlsx"
'myFileName = "Tax Initiation_" & myProgram & "_" & myLast & "_" & myFirst & "_" & myEMPID & "_" & myHome & "_" & myHost & "_" & Format(Now(), "yyyymmddhhmmss") & ".xlsx"
myFileName = myLast & "_" & myFirst & "_" & myProgram & "_" & myEMPID & "_" & myHome & "_" & myHost & "_" & Format(Now(), "yyyy-mm-dd-hhmmss") & ".xlsx"
initName = "TAX_IF_RECEIVED_" & myLast & "_" & myFirst & "_" & myEMPID & "_" & myProgram & "_" & myHost

Set xlApp = New Excel.Application
xlApp.Workbooks.Open myTemplate
Set myWbk = xlApp.ActiveWorkbook
myWbk.Sheets("rptTAX").Cells.ClearContents
myWbk.Sheets("rptEQT").Cells.ClearContents
myWbk.Sheets("rptLCH").Cells.ClearContents
myWbk.Save
myWbk.Close

DoCmd.TransferSpreadsheet acExport, , "rptTAX", myTemplate
DoCmd.TransferSpreadsheet acExport, , "rptEQT", myTemplate
DoCmd.TransferSpreadsheet acExport, , "rptLCH", myTemplate

xlApp.Workbooks.Open myTemplate
xlApp.ActiveWorkbook.SaveAs myPath & "\TAX_INITIATIONS\" & myFileName
xlApp.ActiveWorkbook.Close False
xlApp.Quit

myFileName = Left(myFileName, Len(myFileName) - 4)

Backup_Str = myPath & "\TAX_INITIATIONS\Backup of " & myFileName & "xlk"
If FS.FileExists(Backup_Str) Then FS.deletefile (Backup_Str)
'EmailIF (myFileName)
EmailReceipt myInitiator:=initEmail, myInitiation:=initName

End Sub

Sub ExportReport()
Dim myFileName, myReport, myFile, myTable
Dim xlApp As Excel.Application

myTable = "SS_REPORT"
DeleteQuery myTable
CurrentDb.QueryDefs("qrySS_REPORT").Execute

myFileName = "Tax Initiation Report " & Format(Now(), "yyyy-mm-dd-hhmmss") & ".xlsx"
myFile = CurrentProject.Path & "\REPORTS\" & myFileName
DoCmd.TransferSpreadsheet acExport, , "SS_REPORT", myFile 'save data from table SS report to excel file
CurrentDb.QueryDefs("SS_REPORT_APPEND").Execute

Set xlApp = New Excel.Application

Set myReport = xlApp.Workbooks.Open(myFile)
myReport.Sheets(1).Select
myReport.Sheets(1).Activate

xlApp.ActiveWindow.Zoom = 80

With myReport.Sheets(1).Cells.Font
    .Name = "Arial"
    .Size = 10
    .Underline = xlUnderlineStyleNone
    .ColorIndex = xlAutomatic
End With
    
myReport.Sheets(1).Rows("1:1").Font.Bold = True
myReport.Sheets(1).Columns("A:AZ").EntireColumn.AutoFit
myReport.Sheets(1).Range("A1").Select

myReport.Close True

xlApp.Quit

'EmailReport (myFileName)

End Sub

Sub ExportReportMonthly()
Dim myFileName, myReport, myFile, myTable
Dim xlApp As Excel.Application

myFileName = "Monthly Tax Initiation Report " & Format(Now() - 5, "yyyy-mmm") & ".xlsx"
myFile = CurrentProject.Path & "\REPORTS\" & myFileName
DoCmd.TransferSpreadsheet acExport, , "SS_REPORT_MONTHLY", myFile

Set xlApp = New Excel.Application

Set myReport = xlApp.Workbooks.Open(myFile)
myReport.Sheets(1).Select
myReport.Sheets(1).Activate

xlApp.ActiveWindow.Zoom = 80

With myReport.Sheets(1).Cells.Font
    .Name = "Arial"
    .Size = 10
    .Underline = xlUnderlineStyleNone
    .ColorIndex = xlAutomatic
End With
    
myReport.Sheets(1).Rows("1:1").Font.Bold = True
myReport.Sheets(1).Columns("A:AZ").EntireColumn.AutoFit
myReport.Sheets(1).Range("A1").Select

myReport.Close True

xlApp.Quit

'EmailReport (myFileName)

End Sub
