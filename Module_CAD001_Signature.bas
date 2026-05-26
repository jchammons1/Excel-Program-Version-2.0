Attribute VB_Name = "Module_CAD001_Signature"
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


