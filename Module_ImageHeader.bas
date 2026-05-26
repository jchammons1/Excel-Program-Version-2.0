Attribute VB_Name = "Module_ImageHeader"
Option Explicit
Dim SheetTypeName As String

    Dim sourceWs As Worksheet
    Dim targetWs As Worksheet
    Dim img As Shape
    Dim newImg As Shape
    


'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'*********************************** Remove and replace all the template header images  **************************************************>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
    
Sub CopyImageToAnotherSheet()


    DATA_VALIDATION.Visible = True
    DATA_VALIDATION.Select
    
        Dim VisibleSheets As Range
            For Each VisibleSheets In Range("G2:G9")
                SheetTypeName = VisibleSheets
                Sheets(SheetTypeName).Select
                Sheets(SheetTypeName).Unprotect
                    Call DeleteImageIfExists
                    Call ChangeImage
                Sheets(SheetTypeName).Select
                Sheets(SheetTypeName).Protect
            Next VisibleSheets

        ' The reason these sheets are separately selected instead of a loop is that they a not visible or need a password unlike the other sheets
                Form902.Select
                Form902.Unprotect Password:="roadway123"
                    Call DeleteImageIfExists2
                    Call ChangeImage2
                Form902.Select
                Form902.Protect Password:="roadway123"

                EngPayInv.Visible = True
                EngPayInv.Select
                EngPayInv.Unprotect
                    Call DeleteImageIfExists3
                    Call ChangeImage3
                EngPayInv.Select
                EngPayInv.Protect
                EngPayInv.Visible = False
                
                CAD.Visible = True
                CAD.Select
                CAD.Unprotect Password:="roadway123"
                    Call DeleteImageIfExists4
                    Call ChangeImage4
                CAD.Select
                CAD.Protect Password:="roadway123"
                CAD.Visible = False

   DATA_VALIDATION.Protect
   DATA_VALIDATION.Visible = False
   START.Select

End Sub

Sub ChangeImage()
        ' Set the source and target worksheets
            Dim TargetCell As Range
            Set sourceWs = ThisWorkbook.Sheets("Image")
            Set targetWs = ThisWorkbook.Sheets(SheetTypeName)
            
        ' Reference the image by its name in the source worksheet
            Set img = sourceWs.Shapes("UploadedImage")

        ' Copy the image to the target worksheet
            img.Copy
            'targetWs.Range("A1").Select
            targetWs.Paste
            
        ' Reference the newly pasted image
            Set newImg = targetWs.Shapes(targetWs.Shapes.Count)
            
        ' Example: Change the name of the new image
            newImg.Name = "UploadedImage"
           
        ' Optional: Move the new image to a specific cell (e.g., A1)
            With targetWs.Range("A1")
                newImg.Left = .Left
                newImg.Top = .Top
            End With
End Sub


Sub DeleteImageIfExists()
    Dim ws As Worksheet
    Dim img As Shape
    Dim imgName As String
    
    ' Set the worksheet where the image might be located
    Set ws = ThisWorkbook.Sheets("Signature")
    
    ' Specify the name of the image to check
    imgName = "UploadedImage_Signature"
    
    ' Check if an image with the specified name exists
    On Error Resume Next
    Set img = ws.Shapes(imgName)
    On Error GoTo 0
    
    ' If the image exists, delete it
    If Not img Is Nothing Then
        img.Delete
    End If
End Sub



Sub ChangeImage2()
        ' Set the source and target worksheets
            Dim TargetCell As Range
            Set sourceWs = ThisWorkbook.Sheets("Image")
            Set targetWs = ThisWorkbook.Sheets("902")
            
        ' Reference the image by its name in the source worksheet
            Set img = sourceWs.Shapes("UploadedImage")

        ' Copy the image to the target worksheet
            img.Copy
            targetWs.Range("A1").Select
            targetWs.Paste
            
        ' Reference the newly pasted image
            Set newImg = targetWs.Shapes(targetWs.Shapes.Count)
            
        ' Example: Change the name of the new image
            newImg.Name = "UploadedImage"
           
        ' Optional: Move the new image to a specific cell (e.g., A1)
            With targetWs.Range("A1")
                newImg.Left = .Left
                newImg.Top = .Top
            End With
End Sub


Sub DeleteImageIfExists2()
    Dim ws As Worksheet
    Dim img As Shape
    Dim imgName As String
    
    ' Set the worksheet where the image might be located
    Set ws = ThisWorkbook.Sheets("902")
    
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


Sub ChangeImage3()
        ' Set the source and target worksheets
            Dim TargetCell As Range
            Set sourceWs = ThisWorkbook.Sheets("Image")
            Set targetWs = ThisWorkbook.Sheets("EngPayInv")
            
        ' Reference the image by its name in the source worksheet
            Set img = sourceWs.Shapes("UploadedImage")

        ' Copy the image to the target worksheet
            img.Copy
            targetWs.Range("A1").Select
            targetWs.Paste
            
        ' Reference the newly pasted image
            Set newImg = targetWs.Shapes(targetWs.Shapes.Count)
            
        ' Example: Change the name of the new image
            newImg.Name = "UploadedImage"
           
        ' Optional: Move the new image to a specific cell (e.g., A1)
            With targetWs.Range("A1")
                newImg.Left = .Left
                newImg.Top = .Top
            End With
End Sub


Sub DeleteImageIfExists3()
    Dim ws As Worksheet
    Dim img As Shape
    Dim imgName As String
    
    ' Set the worksheet where the image might be located
    Set ws = ThisWorkbook.Sheets("EngPayInv")
    
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


Sub ChangeImage4()
        ' Set the source and target worksheets
            Dim TargetCell As Range
            Set sourceWs = ThisWorkbook.Sheets("Image")
            Set targetWs = ThisWorkbook.Sheets("CAD_Template")
            
        ' Reference the image by its name in the source worksheet
            Set img = sourceWs.Shapes("UploadedImage")

        ' Copy the image to the target worksheet
            img.Copy
            targetWs.Range("A1").Select
            targetWs.Paste
            
        ' Reference the newly pasted image
            Set newImg = targetWs.Shapes(targetWs.Shapes.Count)
            
        ' Example: Change the name of the new image
            newImg.Name = "UploadedImage"
           
        ' Optional: Move the new image to a specific cell (e.g., A1)
            With targetWs.Range("A1")
                newImg.Left = .Left
                newImg.Top = .Top
            End With
End Sub


Sub DeleteImageIfExists4()
    Dim ws As Worksheet
    Dim img As Shape
    Dim imgName As String
    
    ' Set the worksheet where the image might be located
    Set ws = ThisWorkbook.Sheets("CAD_Template")
    
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

