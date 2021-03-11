Attribute VB_Name = "SQL"
Option Explicit
Option Private Module

Public Function ExecuteQuery(pSql As String) As Object
  On Error GoTo handleError

  Dim cxn As Object
  Dim db As String
  Dim dbName As String

'  dbName = ThisWorkbook.Path & "\DB.accdb" 'Access
  dbName = ThisWorkbook.Path & "\DB.db" 'SQLite

  If Not Dir(dbName, vbArchive) = Empty Then
    Set cxn = CreateObject("ADODB.Connection")

'    db = "Provider=Microsoft.ACE.OLEDB.12.0; Data source=" & dbName 'Access
    db = "Driver=SQLite3 ODBC Driver; Database=" & dbName 'SQLite

    cxn.Open db

    Set ExecuteQuery = cxn.Execute(pSql)
    Set cxn = Nothing
  Else
    MsgBox "No se ha encontrado la base de datos"
  End If

handleError:
  If Err.Number <> 0 Then
    MsgBox Err.Description
    End
  End If
End Function
