VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_DocumentHeading 
   Caption         =   "Update Document Image and Header"
   ClientHeight    =   5325
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21270
   OleObjectBlob   =   "UserForm_DocumentHeading.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_DocumentHeading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBrowse_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Select an Image"
        .Filters.Add "Image Files", "*.jpg; *.jpeg; *.png; *.bmp; *.gif", 1
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            txtImagePath.Text = .SelectedItems(1)
        End If
    End With
End Sub


Sub DeleteImageIfExists()
    Dim ws As Worksheet
    Dim img As Shape
    Dim imgName As String
    
    ' Set the worksheet where the image might be located
    Set ws = ThisWorkbook.Sheets("Image")
    
    ' Specify the name of the image to check
    imgName = "UploadedImage"
    
    ' Check if an image with the specified name exists
    On Error Resume Next
    Set img = ws.Shapes(imgName)
    On Error GoTo 0
    
    ' If the image exists, delete it
    If Not img Is Nothing Then
        img.Delete
    End If
End Sub



Private Sub Upload_Image_Click()
    Application.ScreenUpdating = False
    PROJECT_DATA.Visible = True
    PROJECT_DATA.Unprotect
    PROJECT_DATA.Range("AP2") = txt_Heading.Value
    PROJECT_DATA.Visible = False
    PROJECT_DATA.Protect
    
    Image.Visible = True
    Image.Unprotect
    Dim ws As Worksheet
    Dim imgPath As String
    Dim img As Shape
    Dim cell As Range
    
    ' Set the worksheet where you want to upload the image
    Set ws = ThisWorkbook.Sheets("Image")
    
    ' Get the image path from the TextBox
    imgPath = txtImagePath.Text
    
    
    ' Check if the path is not empty
    If imgPath <> "" Then
        ' Add the image to the worksheet
        Set img = ws.Shapes.AddPicture(imgPath, _
                                       msoFalse, msoCTrue, _
                                       100, 50, 125, 50)
        
        ' Example: Change the name of the image
    Call DeleteImageIfExists

    ' Example: Change the name of the image
        img.Name = "UploadedImage"
  
    ' Reference the image by its name
    Set img = ws.Shapes("UploadedImage") ' Change this to the name of your image
    
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

    Image.Visible = False
    Image.Protect
    
    
    Call CopyImageToAnotherSheet
    
    Application.ScreenUpdating = True
    Unload Me
End Sub


Private Sub UserForm_Initialize()
    txt_Heading.Value = PROJECT_DATA.Range("AP2")
End Sub
