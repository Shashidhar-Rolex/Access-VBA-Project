Attribute VB_Name = "modProcess"
Option Compare Database

Sub ProcessIF()

Dim myTest, myTable

DeleteQuery ("rptTAX")
DeleteQuery ("tblTEMP_TAX")

CurrentDb.QueryDefs("qryTEMP_TAX").Execute

On Error Resume Next
myTest = CurrentDb.TableDefs("tblTEMP_TAX").OpenRecordset.Fields("Field1").Value
If Err.Number > 0 Then Exit Sub
On Error GoTo 0

Do While myTest <> ""

    CurrentDb.QueryDefs("zrptTAX").Execute
    
    ExportIF
    'MsgBox Err.Number
    CurrentDb.QueryDefs("qryTEMP_TAX_DELETE").Execute
    DeleteQuery ("rptTAX")

    On Error Resume Next
    myTest = CurrentDb.TableDefs("tblTEMP_TAX").OpenRecordset.Fields("Field1").Value
    If Err > 0 Then Exit Sub
    Err = 0

Loop

End Sub
