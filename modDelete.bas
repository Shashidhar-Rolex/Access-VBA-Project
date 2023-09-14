Attribute VB_Name = "modDelete"
Option Compare Database

Sub DeleteQuery(myTable)
Dim StrSql

StrSql = "DELETE [" & myTable & "].*" _
       & "FROM [" & myTable & "];"

CurrentDb.QueryDefs("zqryDELETE").SQL = StrSql
CurrentDb.QueryDefs("zqryDELETE").Execute

End Sub

