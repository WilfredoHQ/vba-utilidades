Attribute VB_Name = "SQL"
Option Explicit
Option Private Module

Public Function ExecuteQuery(pSql As String) As Object
  On Error GoTo handleError

  Dim cn As Object
  Dim ddbb As String
  Dim dbName As String

'  dbName = ThisWorkbook.Path & "\DDBB.accdb" 'Access
  dbName = ThisWorkbook.Path & "\DDBB.db" 'SQLite

  If Not Dir(dbName, vbArchive) = "" Then
    Set cn = CreateObject("ADODB.Connection")

'    ddbb = "Provider=Microsoft.ACE.OLEDB.12.0; Data source=" & dbName 'Access
    ddbb = "Driver=SQLite3 ODBC Driver; Database=" & dbName 'SQLite

    cn.Open ddbb

    Set ExecuteQuery = cn.Execute(pSql)
    Set cn = Nothing
  Else
    MsgBox "No se ha encontrado la base de datos"
  End If

handleError:
  If Err.Number <> 0 Then
    MsgBox Err.Description
    End
  End If
End Function
