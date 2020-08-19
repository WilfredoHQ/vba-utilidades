Attribute VB_Name = "CRUD"
Option Explicit
Option Private Module

Public Function GetRecords(pInTheSheet As Worksheet, pFromTheRow As Integer, pInTheColumn As Integer, pWantedValue As String, pIsExact As Boolean) As String
  Dim lastRowContained As Long
  Dim i As Long

  lastRowContained = GetEmptyCell(pInTheSheet, CLng(pFromTheRow), CLng(pInTheColumn)) - 1

  For i = pFromTheRow To lastRowContained
    If pIsExact = True Then
      If UCase(pInTheSheet.Cells(i, pInTheColumn)) = UCase(pWantedValue) Then
        GetRecords = i
      End If
    Else
      If UCase(pInTheSheet.Cells(i, pInTheColumn)) Like UCase("*" & pWantedValue & "*") Then
        GetRecords = IIf(GetRecords = Empty, i, GetRecords & " " & i)
      End If
    End If
  Next i

  GetRecords = IIf(GetRecords = Empty, 0, GetRecords)
End Function

Public Sub ManageRecord(pInTheSheet As Worksheet, pInTheRow As Long, pFromTheColumn As Integer, pData() As String)
  Dim i As Integer

  For i = pFromTheColumn To pFromTheColumn + UBound(pData)
    pInTheSheet.Cells(pInTheRow, i) = pData(i - pFromTheColumn)
  Next i
End Sub

Public Sub DeleteRecord(pInTheSheet As Worksheet, pRow As Long)
  On Error GoTo handleError

  pInTheSheet.Rows(pRow).EntireRow.Delete

handleError:
  If Err.Number = 1004 Then
    MsgBox "No se ha encontrado el registro"
  End If
End Sub
