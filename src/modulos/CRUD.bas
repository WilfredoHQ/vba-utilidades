Attribute VB_Name = "CRUD"
Option Explicit
Option Private Module

Public Function GetRecords(pFromTheSheet As Worksheet, pFromTheRow As Byte, pInTheColumn As Byte, pWantedValue As String, Optional pIsApproximateMatch As Boolean) As String
  Dim lastRowContained As Long
  Dim currentRow As Long

  lastRowContained = GetEmptyCell(pFromTheSheet, pFromTheRow, pInTheColumn) - 1

  For currentRow = pFromTheRow To lastRowContained
    If pIsApproximateMatch Then
      If LCase(pFromTheSheet.Cells(currentRow, pInTheColumn)) Like LCase("*" & pWantedValue & "*") Then
        GetRecords = IIf(GetRecords = Empty, currentRow, GetRecords & " " & currentRow)
      End If
    Else
      If LCase(pFromTheSheet.Cells(currentRow, pInTheColumn)) = LCase(pWantedValue) Then
        GetRecords = currentRow
      End If
    End If
  Next currentRow

  GetRecords = IIf(GetRecords = Empty, 0, GetRecords)
End Function

Public Sub ManageRecord(pInTheSheet As Worksheet, pInTheRow As Long, pFromTheColumn As Byte, pData() As String)
  Dim currentColumn As Integer

  For currentColumn = pFromTheColumn To pFromTheColumn + UBound(pData)
    pInTheSheet.Cells(pInTheRow, currentColumn) = pData(currentColumn - pFromTheColumn)
  Next currentColumn
End Sub

Public Sub DeleteRecord(pInTheSheet As Worksheet, pTheRow As Long)
  pInTheSheet.Rows(pTheRow).EntireRow.Delete
End Sub
