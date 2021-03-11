Attribute VB_Name = "Procedures"
Option Explicit
Option Private Module

Public Sub CleanControls(pInTheForm As MSForms.UserForm, Optional pWithTheTag As String)
  On Error Resume Next

  Dim ctrl As MSForms.Control

  For Each ctrl In pInTheForm.Controls
    If pWithTheTag <> Empty And Not CBool(InStr(ctrl.Tag, pWithTheTag)) Then GoTo continueForLoop
    If TypeOf ctrl Is MSForms.TextBox Then ctrl.Text = Empty
    If TypeOf ctrl Is MSForms.ComboBox Then ctrl.Style = fmStyleDropDownCombo: ctrl.Text = Empty: ctrl.Style = fmStyleDropDownList
    If TypeOf ctrl Is MSForms.CheckBox Or TypeOf ctrl Is MSForms.OptionButton Then ctrl.Value = False
    If TypeOf ctrl Is MSForms.ListBox Then ctrl.Clear

continueForLoop:
  Next ctrl
End Sub

Public Sub FillComboBox(pFromTheSheet As Worksheet, pFromTheRow As Integer, pInTheColumn As Byte, pComboBox As MSForms.ComboBox)
  Dim lastRowContained As Long
  Dim currentRow As Long

  lastRowContained = GetEmptyCell(pFromTheSheet, CLng(pFromTheRow), pInTheColumn) - 1

  For currentRow = pFromTheRow To lastRowContained
      pComboBox.AddItem pFromTheSheet.Cells(currentRow, pInTheColumn)
  Next currentRow
End Sub

Public Sub GeneratePDF(pTheSheet As Worksheet, pFileName As String)
  pTheSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=ThisWorkbook.Path & "\assets\pdf\" & pFileName & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
End Sub

Public Sub FormDesign(pToForm As MSForms.UserForm)
  On Error Resume Next

  Dim ctrl As MSForms.Control

  pToForm.BackColor = vbWhite

  For Each ctrl In pToForm.Controls
    ctrl.BackColor = vbWhite
    ctrl.BackStyle = fmBackStyleTransparent
    ctrl.TabStop = False

    If TypeOf ctrl Is MSForms.CommandButton Then
      ctrl.MousePointer = fmMousePointerUpArrow
      ctrl.TakeFocusOnClick = False
    End If

    If TypeOf ctrl Is MSForms.TextBox Or TypeOf ctrl Is MSForms.ComboBox Then
      ctrl.SelectionMargin = False
      ctrl.SpecialEffect = fmSpecialEffectEtched
      ctrl.TabStop = True
      ctrl.TextAlign = fmTextAlignCenter
    End If

    If TypeOf ctrl Is MSForms.ComboBox Then
      ctrl.Style = fmStyleDropDownList
    End If

    If TypeOf ctrl Is MSForms.CheckBox Or TypeOf ctrl Is MSForms.OptionButton Then
      ctrl.SpecialEffect = fmSpecialEffectFlat
    End If
  Next ctrl
End Sub
