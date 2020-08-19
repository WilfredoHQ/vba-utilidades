Attribute VB_Name = "Images"
Option Explicit
Option Private Module

Public Sub UploadImage(pInTheControl As MSForms.Image)
  Dim fileName As String
  fileName = Application.GetOpenFilename("JPG Files (*.jpg), *.jpg")

  If fileName <> "Falso" Then
    pInTheControl.Picture = LoadPicture(CStr(fileName))
    pInTheControl.Tag = fileName
  End If
End Sub

Public Sub SaveImage(pFromTheControl As MSForms.Image, pNewFileName As String)
  Dim location As String
  location = ThisWorkbook.Path

  If pFromTheControl.Tag <> Empty Then
    FileCopy pFromTheControl.Tag, location & "\assets\images\" & pNewFileName & ".jpg"
    pFromTheControl.Tag = Empty
  End If
End Sub

Public Sub LoadImage(pInTheControl As MSForms.Image, pNewFileName As String)
  On Error GoTo handleError

  Dim location As String
  location = ThisWorkbook.Path

  pInTheControl.Picture = LoadPicture(location & "\assets\images\" & pNewFileName & ".jpg")

handleError:
  If Err = 53 Then
    pInTheControl.Picture = LoadPicture(location & "\assets\images\sin_foto.jpg")
  End If
End Sub

Public Sub DeleteImage(pNewFileName As String)
  On Error Resume Next

  Dim location As String
  location = ThisWorkbook.Path

  Kill location & "\assets\images\" & pNewFileName & ".jpg"
End Sub

Public Sub CleanImage(pTheControl As MSForms.Image)

  pTheControl.Picture = LoadPicture("")

End Sub
