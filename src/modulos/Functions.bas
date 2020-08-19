Attribute VB_Name = "Functions"
Option Explicit
Option Private Module

Public Function SHA256(pString As String) As String
  Dim oD As Object, oT As Object, oSHA256 As Object
  Dim textToHash() As Byte, bytes() As Byte

  Set oT = CreateObject("System.Text.UTF8Encoding")
  Set oSHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
  Set oD = CreateObject("MSXML2.DOMDocument")

  textToHash = oT.GetBytes_4(pString)
  bytes = oSHA256.ComputeHash_2((textToHash))

  With oD
    .LoadXML "<root />"
    .DocumentElement.DataType = "bin.Hex"
    .DocumentElement.nodeTypedValue = bytes
  End With

  SHA256 = Replace(oD.DocumentElement.Text, vbLf, "")

  Set oT = Nothing
  Set oSHA256 = Nothing
  Set oD = Nothing
End Function

Public Function GetEmptyCell(pInTheSheet As Worksheet, pFromTheRow As Long, pInTheColumn As Long) As Long
  Do Until IsEmpty(pInTheSheet.Cells(pFromTheRow, pInTheColumn))
    pFromTheRow = pFromTheRow + 1
  Loop

  GetEmptyCell = pFromTheRow
End Function

Public Function ValidateFields(pInTheForm As MSForms.UserForm) As Boolean
  Dim ctrl As MSForms.Control
  Dim oReg As Object

  For Each ctrl In pInTheForm.Controls
    If CBool(InStr(ctrl.Tag, "required")) Then
      If ctrl.Text = Empty Then
        ValidateFields = False
        MsgBox "Este campo es obligatorio"
        ctrl.SetFocus
        Exit Function
      End If
    End If

    If CBool(InStr(ctrl.Tag, "number")) Then
      If Not IsNumeric(ctrl.Text) And ctrl.Text <> Empty Then
        ValidateFields = False
        MsgBox "Este campo solo acepta valores numericos"
        ctrl.Text = Empty
        ctrl.SetFocus
        Exit Function
      End If
    End If

    If CBool(InStr(ctrl.Tag, "date")) Then
      If Not IsDate(ctrl.Text) And ctrl.Text <> Empty Then
        ValidateFields = False
        MsgBox "Este campo debe de ser un fecha"
        ctrl.Text = Empty
        ctrl.SetFocus
        Exit Function
      End If
    End If

    If CBool(InStr(ctrl.Tag, "email")) Then
      Set oReg = CreateObject("VBScript.RegExp")
      oReg.Pattern = "^([a-zA-Z0-9_\-\.]+)@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$"

      If Not oReg.Test(ctrl.Text) And ctrl.Text <> Empty Then
        ValidateFields = False
        Set oReg = Nothing
        MsgBox "Este campo debe ser un email"
        ctrl.Text = Empty
        ctrl.SetFocus
        Exit Function
      End If
    End If
  Next ctrl

  ValidateFields = True
End Function

Public Function GetSerialNumber() As String
  Dim fso As Object, drv As Object
  Dim serialNumberDec As Long
  Dim serialNumberHex As String

  Set fso = CreateObject("Scripting.FileSystemObject")
  Set drv = fso.GetDrive("C")

  With drv
    If .IsReady Then
      serialNumberDec = Abs(.SerialNumber)
    Else
      serialNumberDec = -1
    End If
  End With

  serialNumberHex = Application.WorksheetFunction.Dec2Hex(serialNumberDec)
  GetSerialNumber = serialNumberHex

  Set drv = Nothing
  Set fso = Nothing
End Function

Public Function GenerateUuid() As String
  Dim k As Integer
  Dim h As String

  GenerateUuid = Space(36)

  For k = 1 To Len(GenerateUuid)
    Randomize
    Select Case k
      Case 9, 14, 19, 24: h = "-"
      Case 15:            h = "4"
      Case 20:            h = Hex(Rnd * 3 + 8)
      Case Else:          h = Hex(Rnd * 15)
    End Select
    Mid(GenerateUuid, k, 1) = h
  Next k
End Function

Public Function StrFormat(pString As String, pValues() As String) As String
  Dim i As Long

  For i = LBound(pValues) To UBound(pValues)
    pString = Replace(pString, "{" & i & "}", pValues(i))
  Next

  StrFormat = pString
End Function
