Attribute VB_Name = "modImport"
Option Compare Database

Sub ImportIF()

Dim myPath, myFile, myTable, mySpec, myFileExist, myCount
Dim xlApp As Excel.Application
Dim xlMain As Excel.Workbook
Dim FS
Dim assignID, tempFile
Dim db As Database
Dim rs1 As DAO.Recordset
Dim strsqlIn As String

DoCmd.SetWarnings False
Set db = CurrentDb

Set FS = CreateObject("Scripting.FileSystemObject")
myPath = CurrentProject.Path & "\INITIATIONS"
myTable = "IF_DATA"
mySpec = "EYTAX_IF"
DeleteQuery (myTable)
myCount = 1

Do While myCount > 0

    myFileExist = myPath & "\" & "EYTAX_IF.txt"
    myFile = Dir(myPath & "\EYTAX*.dat")
    myFile1 = Dir(myPath & "\EYTAX*.dat")
    myFile = myPath & "\" & myFile
    
    'CB 18 July - extract AssignID from dat file name
    tempFile = Mid(myFile1, 7, Len(myFile1) - 7)
    assignID = Mid(tempFile, 1, InStr(1, myFile1, "_", 0) - 1)
    
    Set FS = CreateObject("Scripting.FileSystemObject")
    If FS.FileExists(myFile) Then
    Else
        MsgBox "The file: " & vbCr & vbCr & myFile & vbCr & vbCr & " does not exist for import.  Please put it in the INITIATIONS folder and try again.", vbExclamation
        Exit Sub
    End If
    
    If FS.FileExists(myFileExist) Then FS.deletefile (myFileExist)
    Name myFile As myFileExist
    DoCmd.TransferText acImportDelim, mySpec, myTable, myFileExist
    
    If FS.FileExists(myFileExist) Then FS.deletefile (myFileExist)
    strsqlIn = "SELECT * from IF_DATA where Field1='TAX';"
    Set rs1 = CurrentDb.OpenRecordset(strsqlIn)
    
    'MsgBox rs1.RecordCount
     For icnt = 1 To rs1.RecordCount
        'ID = rs1.Fields("empid").Value
        'Name = rs1.Fields("empname").Value
        'DoCmd.RunSQL ("insert into HISTORIC_IF_DATA values ('" & ID & "','" & Name & "','" & "" & "','" & "" & "')")
            Val0 = rs1.Fields("Field1").Value
            Val1 = rs1.Fields("Field2").Value
            Val2 = rs1.Fields("Field3").Value
            Val3 = rs1.Fields("Field4").Value
            Val4 = rs1.Fields("Field5").Value
            Val5 = rs1.Fields("Field6").Value
            Val6 = rs1.Fields("Field7").Value
            Val7 = rs1.Fields("Field8").Value
            Val8 = rs1.Fields("Field9").Value
            Val9 = rs1.Fields("Field10").Value
            Val10 = rs1.Fields("Field11").Value
            Val11 = rs1.Fields("Field12").Value
            Val12 = rs1.Fields("Field13").Value
            Val13 = rs1.Fields("Field14").Value
            Val14 = rs1.Fields("Field15").Value
            Val15 = rs1.Fields("Field16").Value
            Val16 = rs1.Fields("Field17").Value
            Val17 = rs1.Fields("Field18").Value
            Val18 = rs1.Fields("Field19").Value
            Val19 = rs1.Fields("Field20").Value
            Val20 = rs1.Fields("Field21").Value
            Val21 = rs1.Fields("Field22").Value
            Val22 = rs1.Fields("Field23").Value
            Val23 = rs1.Fields("Field24").Value
            Val24 = rs1.Fields("Field25").Value
            Val25 = rs1.Fields("Field26").Value
            Val26 = rs1.Fields("Field27").Value
            Val27 = rs1.Fields("Field28").Value
            
            strSL = "insert into HISTORIC_IF_DATA (Field1, Field2, Field3, Field4, Field5, Field6, Field7, Field8, Field9, Field10, Field11, Field12, Field13, Field14, Field15, Field16, Field17, Field18, Field19, Field20, Field21, Field22, Field23, Field24, Field25, Field26, Field27, Field28, Field29, Field30,[TIME_STAMP],[AssignID])" & _
            "values ('" & Val0 & "','" & Val1 & "','" & Val2 & "','" & Val3 & "','" & Val4 & "','" & Val5 & "','" & Val6 & "','" & Val7 & "','" & Val8 & "','" & Val9 & "','" & Val10 & "','" & Val11 & "','" & Val12 & "','" & Val13 & "','" & Val14 & "','" & Val15 & "','" & Val16 & "','" & Val17 & "','" & Val18 & "','" & Val19 & "','" & Val20 & "','" & Val21 & "','" & Val22 & "','" & Val23 & "','" & Val24 & "','" & Val25 & "','" & Val26 & "','" & Val27 & "','" & "" & "','" & "" & "','" & Now() & "','" & assignID & "');"
            DoCmd.RunSQL (strSL)
     rs1.MoveNext
  Next icnt
    myFile = Dir(myPath & "\EYTAX*.dat")
    If myFile = "" Then
        myCount = 0
    Else
        myCount = 1
    End If
    
Loop

'CurrentDb.QueryDefs("qryHISTORIC_IF_DATA_APPEND").Execute

End Sub

