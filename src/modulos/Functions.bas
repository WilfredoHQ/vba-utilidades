Attribute VB_Name = "Functions"
Option Explicit
Option Private Module

Public Function Sha256(pTheString As String) As String
  Dim oDoc As Object, oUTF As Object, oSHA256 As Object

  Set oDoc = CreateObject("MSXML2.DOMDocument")
  Set oUTF = CreateObject("System.Text.UTF8Encoding")
  Set oSHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")

  With oDoc
    .LoadXML "<root />"
    .DocumentElement.DataType = "bin.Hex"
    .DocumentElement.nodeTypedValue = oSHA256.ComputeHash_2(oUTF.GetBytes_4(pTheString))
  End With

  Sha256 = Replace(oDoc.DocumentElement.Text, vbLf, "")

  Set oDoc = Nothing
  Set oUTF = Nothing
  Set oSHA256 = Nothing
End Function

Public Function GetEmptyCell(pFromTheSheet As Worksheet, pFromTheRow As Byte, pInTheColumn As Byte) As Long
  GetEmptyCell = pFromTheRow

  Do Until IsEmpty(pFromTheSheet.Cells(GetEmptyCell, pInTheColumn))
    GetEmptyCell = GetEmptyCell + 1
  Loop
End Function

Public Function ValidateFields(pInTheForm As MSForms.UserForm) As Boolean
  Dim ctrl As MSForms.Control
  Dim oRegex As Object

  ValidateFields = True
  Set oRegex = CreateObject("VBScript.RegExp")
  oRegex.Pattern = "^([a-zA-Z0-9_\-\.]+)@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$"

  For Each ctrl In pInTheForm.Controls
    If TypeOf ctrl Is MSForms.TextBox Or TypeOf ctrl Is MSForms.ComboBox Then
      With ctrl
        If .Text <> Empty And CBool(InStr(.Tag, "number")) And Not IsNumeric(.Text) Then
          ValidateFields = False: MsgBox "El siguiente campo debe ser un NUMERO"
        ElseIf .Text <> Empty And CBool(InStr(.Tag, "date")) And Not IsDate(.Text) Then
          ValidateFields = False: MsgBox "El siguiente campo debe ser una FECHA"
        ElseIf .Text <> Empty And CBool(InStr(.Tag, "email")) And Not oRegex.Test(.Text) Then
          ValidateFields = False: MsgBox "El siguiente campo debe ser un EMAIL"
        ElseIf CBool(InStr(.Tag, "required")) And .Text = Empty Then
          ValidateFields = False: MsgBox "El siguiente campo es OBLIGATORIO"
        End If

        If Not ValidateFields Then ctrl.Text = Empty: ctrl.SetFocus: Exit Function
      End With
    End If
  Next ctrl

  Set oRegex = Nothing
End Function

Public Function GetSerialNumber() As String
  Dim oFileSys As Object, oDrv As Object

  Set oFileSys = CreateObject("Scripting.FileSystemObject")
  Set oDrv = oFileSys.GetDrive("C")

  GetSerialNumber = Application.WorksheetFunction.Dec2Hex(Abs(oDrv.SerialNumber))

  Set oDrv = Nothing
  Set oFileSys = Nothing
End Function

Public Function GenerateUuid() As String
  Dim currentSpace As Integer
  Dim randomChr As String

  GenerateUuid = Space(36)

  For currentSpace = 1 To Len(GenerateUuid)
    Randomize

    Select Case currentSpace
      Case 9, 14, 19, 24: randomChr = "-"
      Case 15: randomChr = "4"
      Case 20: randomChr = Hex(Rnd * 3 + 8)
      Case Else: randomChr = Hex(Rnd * 15)
    End Select

    Mid(GenerateUuid, currentSpace, 1) = randomChr
  Next currentSpace
End Function

Public Function StrFormat(pTheString As String, pValues() As String) As String
  Dim currentIndex As Byte

  StrFormat = pTheString

  For currentIndex = LBound(pValues) To UBound(pValues)
    StrFormat = Replace(StrFormat, "{" & currentIndex & "}", pValues(currentIndex))
  Next
End Function
