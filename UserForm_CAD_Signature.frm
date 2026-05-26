VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_CAD_Signature 
   Caption         =   "CAD TEMPLATE Signature Upload"
   ClientHeight    =   3945
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12330
   OleObjectBlob   =   "UserForm_CAD_Signature.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_CAD_Signature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SheetTypeName As String

    Dim sourceWs As Worksheet
    Dim targetWs As Worksheet
    Dim img As Shape
    Dim newImg As Shape

Private Sub cmdBrowse_Signature_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Select an Image"
        .Filters.Add "Image Files", "*.jpg; *.jpeg; *.png; *.bmp; *.gif", 1
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            txtImagePath_Signature.Text = .SelectedItems(1)
        End If
    End With
End Sub

Private Sub Upload_ImageSignature_Click()
    Application.ScreenUpdating = False

    Image2.Visible = True
    Image2.Unprotect
    Dim ws As Worksheet
    Dim imgPath As String
    Dim img As Shape
    Dim cell As Range
    
    ' Set the worksheet where you want to upload the image
    Set ws = ThisWorkbook.Sheets("Signature")
    
    ' Get the image path from the TextBox
    imgPath = txtImagePath_Signature.Text
    
    
    ' Check if the path is not empty
    If imgPath <> "" Then
        ' Add the image to the worksheet
        Set img = ws.Shapes.AddPicture(imgPath, _
                                       msoFalse, msoCTrue, _
                                       100, 50, 125, 50)
        
        ' Example: Change the name of the image
    Call DeleteImageIfExists_Signature

    ' Example: Change the name of the image
        img.Name = "ImageSignature"
  
    ' Reference the image by its name
    Set img = ws.Shapes("ImageSignature") ' Change this to the name of your image
    
    ' Set the cell where you want to move the image
    Set cell = ws.Range("A1")
    
    ' Move the image to the specified cell
    img.Left = cell.Left
    img.Top = cell.Top
    img.Width = cell.Width
    img.Height = cell.Height
        
        MsgBox "Image uploaded successfully!", vbInformation
    Else
        MsgBox "Please select an image first.", vbExclamation
    End If

    Image2.Visible = False
    Image2.Protect
    
    
    Call CopytoCAD
    
    Application.ScreenUpdating = True
    Unload Me
End Sub



    


'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'*********************************** Remove and replace all the template header images  **************************************************>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
    
Sub CopytoCAD()
               
                CAD.Visible = True
                CAD.Select
                CAD.Unprotect Password:="roadway123"
                    Call DeleteImageIfExists_CADSignature
                    Call ChangeImage_Signature
                CAD.Select
                CAD.Protect Password:="roadway123"
                CAD.Visible = False

   START.Select

End Sub

Sub ChangeImage_Signature()
        ' Set the source and target worksheets
            Dim TargetCell As Range
            Set sourceWs = ThisWorkbook.Sheets("Signature")
            Set targetWs = ThisWorkbook.Sheets("CAD_template")
            
        ' Reference the image by its name in the source worksheet
            Set img = sourceWs.Shapes("ImageSignature")

        ' Copy the image to the target worksheet
            img.Copy
            'targetWs.Range("A1").Select
            targetWs.Paste
            
        ' Reference the newly pasted image
            Set newImg = targetWs.Shapes(targetWs.Shapes.Count)
            
        ' Example: Change the name of the new image
            newImg.Name = "ImageSignature"
           
        ' Optional: Move the new image to a specific cell (e.g., A1)
            With targetWs.Range("B33")
                newImg.Left = .Left
                newImg.Top = .Top - .Height
                newImg.Width = .Width
            End With
End Sub


Sub DeleteImageIfExists_Signature()
    Dim ws As Worksheet
    Dim img As Shape
    Dim imgName As String
    
    ' Set the worksheet where the image might be located
    Set ws = ThisWorkbook.Sheets("Signature")
    
    ' Specify the name of the image to check
    imgName = "ImageSignature"
    
    ' Check if an image with the specified name exists
    On Error Resume Next
    Set img = ws.Shapes(imgName)
    On Error GoTo 0
    
    ' If the image exists, delete it
    If Not img Is Nothing Then
        img.Delete
    End If
End Sub


Sub DeleteImageIfExists_CADSignature()
    Dim ws As Worksheet
    Dim img As Shape
    Dim imgName As String
    
    ' Set the worksheet where the image might be located
    Set ws = ThisWorkbook.Sheets("CAD_template")
    
    ' Specify the name of the image to check
    imgName = "ImageSignature"
    
    ' Check if an image with the specified name exists
    On Error Resume Next
    Set img = ws.Shapes(imgName)
    On Error GoTo 0
    
    ' If the image exists, delete it
    If Not img Is Nothing Then
        img.Delete
    End If
End Sub


Private Sub UserForm_Initialize()
    Frame_CAD_Signature.BackColor = RGB(173, 216, 230)
End Sub
