Attribute VB_Name = "Module_CAD"
 Option Explicit
    Dim CAD_Selected As String
 
Function SheetExists(SheetName As String) As Boolean
' Determine if the sheet name exist
    Dim ws As Worksheet
    SheetExists = False
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = SheetName Then
            SheetExists = True
            Exit Function
        End If
    Next ws
End Function
'
' Create next CAD Invoice
'
Sub Create_NewCAD()
    
    Application.ScreenUpdating = False
    CAD.Select
    CAD.Unprotect Password:="roadway123"
    CAD.Copy After:=CAD
    Sheets("CAD_Template (2)").Select

    
    Dim LastCAD As String
    Dim NewCAD As String
    Dim i As Integer
    Dim row As Integer
    Dim column As Integer
    Dim endrow As Integer
    
        i = 1
            Do While SheetExists("CAD_No_" & i)
                        i = i + 1
        Loop
        
 ' Name the next sheet after determining the next
        
   NewCAD = "CAD_No_" & i
   Sheets("CAD_Template (2)").Name = NewCAD
   LastCAD = "CAD_No_" & i - 1

        If LastCAD <> "CAD_No_0" Then
        ' CAD-002 past totals
            Sheets(NewCAD).Select
            Sheets(NewCAD).Range("B7") = i
            Sheets(NewCAD).Range("BK53:BL999").Value = Sheets(LastCAD).Range("BO53:BP999").Value          ' copy past period total quantity
            Sheets(NewCAD).Range("AU53:AU999").Value = Sheets(LastCAD).Range("AT53:AT999").Value        ' copy past fuel adjustment totals
            'Example GY2 = Last Invoice Participating Fuel Adjustment / HB2 = Acculative Participating Fuel Adjustment
            Sheets(NewCAD).Range("GY2").Value = Sheets(LastCAD).Range("HB2").Value                      ' copy past fuel adjustment totals
            Sheets(NewCAD).Range("GZ2").Value = Sheets(LastCAD).Range("HC2").Value
            Sheets(NewCAD).Range("HA2").Value = Sheets(LastCAD).Range("HD2").Value
         ' CAD-001 past totals
            Sheets(NewCAD).Range("I10").Value = Sheets(LastCAD).Range("J10").Value
            Sheets(NewCAD).Range("I11").Value = Sheets(LastCAD).Range("J11").Value
            Sheets(NewCAD).Range("I12").Value = Sheets(LastCAD).Range("J12").Value
            Sheets(NewCAD).Range("I13").Value = Sheets(LastCAD).Range("J13").Value
            Sheets(NewCAD).Range("I16").Value = Sheets(LastCAD).Range("J16").Value
            Sheets(NewCAD).Range("I24").Value = Sheets(LastCAD).Range("J24").Value
            Sheets(NewCAD).Range("I25").Value = Sheets(LastCAD).Range("J25").Value
            Sheets(NewCAD).Range("I26").Value = Sheets(LastCAD).Range("J26").Value
            Sheets(NewCAD).Range("I27").Value = Sheets(LastCAD).Range("J27").Value
            Sheets(NewCAD).Range("I28").Value = Sheets(LastCAD).Range("J28").Value
            Sheets(NewCAD).Range("I31").Value = Sheets(LastCAD).Range("J31").Value
            Sheets(NewCAD).Range("I32").Value = Round(Sheets(LastCAD).Range("J32").Value, 4)
            Sheets(NewCAD).Range("I33").Value = Round(Sheets(LastCAD).Range("J33").Value, 4)
            Else
            Sheets(NewCAD).Select                                                                       ' Names the Sheet No. 1
            Sheets(NewCAD).Range("B7") = i
        End If
             
'Do Loop to unlock the cells and change the color to light yellow for visual data entry for the user.
' This code is also in the template.  Not sure that I need this loop below
    Sheets(NewCAD).Select
            row = 50
            column = 1
            endrow = CountPayItems + 53
    
    Do While row < endrow
   ' Dim ParticipatingStatus As String
    
    '    Select Case ParticipatingStatus = Cells(row, column + 49).Value
    If Cells(row, column + 49).Value = "Yes" Or Cells(row, column + 49).Value = "No" _
    Or Cells(row, column + 49).Value = "Correction Participating" Or Cells(row, column + 49).Value = "Correction Non-Participating" Then
          '  Case "Yes", "No", "Correction Participating", "Correction Non-Participating"
                Cells(row, column + 63).Value = Cells(row, column + 62) * Cells(row, column + 6)
                Cells(row, column + 64).Value = 0
                Cells(row, column + 65).Value = 0
                Cells(row, column + 66).Value = 0
                Cells(row, column + 67).Value = 0
     '       End Select
        End If
        row = row + 1
    Loop
''
'
'
'I might want to create the new CAD data directly from the database or when the fuel adjustment is entered, then the CAD template automatically updates.
'  I think when creating the new CAD might be better if there are any modifications to the database, it would automatically capture that.
'
'
'
    Sheets(NewCAD).Select
    Sheets(NewCAD).Protect Password:="roadway123"
End Sub



 Sub Organize_CAD()
    
    ActiveSheet.Unprotect Password:="roadway123"
    CAD_Selected = ActiveSheet.Name
    Range("A53:DE999").Clear
    Range("G30").Value = PROJECT_DATA.Range("J2").Value     ' This value is not a formula directly referencing Project Data because the value might change.
    Range("A53").Select                            'Create a starting position to work from on the CAD
       
    PAYITEMTYPE.Visible = True
    PAYITEMTYPE.Select
    Dim RowType As Range                            'Declare a variable for the Pay Item Type and the range of data from the Data Validation sheet
    For Each RowType In Range("A2:A55")            ' Maximum of 100 types the user can add to the program. 4 default types. 50 User Defined and 50 supplemental types.
' Update message
            UserForm_DASHBOARD.lblMessage.Caption = "Processing Participating Type " & RowType
                    DoEvents
                                                    
                                                   ' The original estimates only go to A2 and A55
        If IsEmpty(RowType) = False Then
        
    ' Select pay items on Database
    Database.Visible = True
    Database.Select
        Call SortPayItems
    Database.Unprotect
    Database.Range("$A$1:$CS$999").AutoFilter
    Range("A2").Select

' Filter pay items to Participating Items'
    ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=47, Criteria1:="=Yes"
' Filter pay items to type'
    ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=46, Criteria1:=RowType.Value
    Range("A2").Select
    Selection.End(xlDown).Select
            If Selection <> "" Then
                Range("A2:C2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
            ' Pasting copied Pay Item, Item Description, and Contract quantity to the CAD
                Sheets(CAD_Selected).Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                ActiveCell.Offset(0, 5).Range("A1").Select
            ' Pasting copied Unit and Unit Price to the Sheets(CAD_Selected)
                Database.Select
                Range("D2:E2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
                Sheets(CAD_Selected).Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                ActiveCell.Offset(0, 3).Range("A1").Select
            ' Pasting copied Subtotal Contract Amount to the Sheets(CAD_Selected)
                Database.Select
                Range("F2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
                Sheets(CAD_Selected).Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                ActiveCell.Offset(0, 1).Range("A1").Select
             ' Pasting copied all other data to the Sheets(CAD_Selected)
                Database.Select
                Range("G2:CS2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
                Sheets(CAD_Selected).Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                    If IsEmpty(Selection.End(xlDown).Value) = True Then
                        ActiveCell.Offset(1, -1).Range("A1").Select
                        Else
                        Selection.End(xlDown).Select
                        ActiveCell.Offset(1, -1).Range("A1").Select
                    End If
                ' Dim RowTypeValue As String
                        ActiveCell.FormulaR1C1 = "=SUMIFS(C,C[40]," & Chr(34) & RowType.Value & Chr(34) & ",C[41],""Yes"")"
                    Call Format_SubtotalAmount_Participating
                ActiveCell.Offset(0, -7).Range("A1").Select
                ActiveCell.FormulaR1C1 = "Subtotal Participating " & RowType.Value & " Items"
                    Call Format_SubtotalText_Participating
                ActiveCell.Offset(2, -1).Range("A1").Select
                
                End If
        End If
    Next RowType
 
            
            Sheets(CAD_Selected).Select
            
            ActiveCell.Offset(0, 1).Range("A1").Select
            ActiveCell.FormulaR1C1 = "FUEL & MATERIAL ADJUSTMENTS PARTICIPATING"
            Selection.Font.Bold = True
            Selection.Font.Size = 16
            ActiveCell.Offset(0, 8).Range("A1").Select
            ActiveCell.FormulaR1C1 = 0          '**********************************************fix THIS RIGHT HERE***************************
            Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            Selection.Font.Size = 16
            Selection.ShrinkToFit = True

        ' Sum and format the project participating totals
            ActiveCell.Offset(2, -8).Range("A1").Select
            ActiveCell.FormulaR1C1 = "PROJECT PARTICIPATING TOTAL"
            Selection.Font.Bold = True
            Selection.Font.Size = 16
            ActiveCell.Offset(0, 7).Range("A1").Select
            ActiveCell.FormulaR1C1 = "=SUMIFS(C,C[41],""Yes"")"
            Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            Selection.Font.Bold = True
            Selection.ShrinkToFit = True
            Selection.Font.Size = 16
            ActiveCell.Offset(2, -8).Range("A1").Select
            
    'Determine if there are any Non-Participating Pay Items.  If not, then skip the Call
            Database.Select
            Database.Range("$A$1:$CS$999").AutoFilter
                ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=47, Criteria1:="No"
                Range("A2").Select
                Selection.End(xlDown).Select
            If Selection <> "" Then
                Sheets(CAD_Selected).Select
                ActiveCell.Offset(0, 0).Range("A1").Select     ' move 1 space to the left because of the stop location from participating total
                Call Organize_CAD_NP
                Database.Range("$A$1:$CS$999").AutoFilter
                Else
                Sheets(CAD_Selected).Select
                Database.Range("$A$1:$CS$999").AutoFilter
            End If
            
            
    Database.Range("$A$1:$CS$999").AutoFilter
    Sheets(CAD_Selected).Select
' Update message
    UserForm_DASHBOARD.lblMessage.Caption = "Finalize and Format document..."
                    DoEvents
                    
        Call Finish_CAD
        Call CAD_ChangeOrder
        Call Fuel_Summary
        Call SetPrintArea_Estimate
        Call Format_CAD
        Call Format_ChangeOrder
        Call Format_FuelSummary
        
    PAYITEMTYPE.Visible = False
    PAYITEMTYPE_SUPPLEMENTAL.Visible = False
    DATA_VALIDATION.Visible = False
    Database.Visible = False
    Sheets(CAD_Selected).Protect Password:="roadway123"
    
 End Sub
 
 

  Sub Organize_CAD_NP()
    
    ActiveSheet.Select
    CAD_Selected = ActiveSheet.Name
        
    PAYITEMTYPE.Select
    Dim RowType As Range                            'Declare a variable for the Pay Item Type and the range of data from the Data Validation sheet
    For Each RowType In Range("A2:A55")            ' Maximum of 100 types the user can add to the program. 4 default types. 50 User Defined and 50 supplemental types.
                                                    ' The original estimates only go to A2 and A55
' Update message
            UserForm_DASHBOARD.lblMessage.Caption = "Processing Non-Participating Type " & RowType
                    DoEvents
                    
        If IsEmpty(RowType) = False Then
        
    ' Select pay items on Database
    Database.Select
    Database.Range("$A$1:$CS$999").AutoFilter
    Range("A2").Select

' Filter pay items to Participating Items'
    ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=47, Criteria1:="No"
' Filter pay items to type'
    ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=46, Criteria1:=RowType.Value
    Range("A2").Select
    Selection.End(xlDown).Select
            If Selection <> "" Then
                Range("A2:C2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
            ' Pasting copied Pay Item, Item Description, and Contract quantity to the CAD
                Sheets(CAD_Selected).Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                ActiveCell.Offset(0, 5).Range("A1").Select
            ' Pasting copied Unit and Unit Price to the Sheets(CAD_Selected)
                Database.Select
                Range("D2:E2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
                Sheets(CAD_Selected).Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                ActiveCell.Offset(0, 3).Range("A1").Select
            ' Pasting copied Subtotal Contract Amount to the Sheets(CAD_Selected)
                Database.Select
                Range("F2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
                Sheets(CAD_Selected).Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                ActiveCell.Offset(0, 1).Range("A1").Select
             ' Pasting copied all other data to the Sheets(CAD_Selected)
                Database.Select
                Range("G2:CS2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
                Sheets(CAD_Selected).Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                    If IsEmpty(Selection.End(xlDown).Value) = True Then
                        ActiveCell.Offset(1, -1).Range("A1").Select
                        Else
                        Selection.End(xlDown).Select
                        ActiveCell.Offset(1, -1).Range("A1").Select
                    End If
                ' Dim RowTypeValue As String
                        ActiveCell.FormulaR1C1 = "=SUMIFS(C,C[40]," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C[41],""No"")"
                    Call Format_SubtotalAmount_NonParticipating
                ActiveCell.Offset(0, -7).Range("A1").Select
                ActiveCell.FormulaR1C1 = "Subtotal Non-Participating " & RowType.Value & " Items"
                    Call Format_SubtotalText_NonParticipating
                ActiveCell.Offset(2, -1).Range("A1").Select
                
                End If
        End If
    Next RowType
    
    Sheets(CAD_Selected).Select
        
        ActiveCell.Offset(0, 1).Range("A1").Select
        ActiveCell.FormulaR1C1 = "FUEL & MATERIAL ADJUSTMENTS NON-PARTICIPATING"
        Selection.Font.Bold = True
        Selection.Font.Size = 16
        Selection.Font.Color = -4165632
        ActiveCell.Offset(0, 8).Range("A1").Select
        ActiveCell.FormulaR1C1 = 0          '**********************************************fix THIS RIGHT HERE***************************
        Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        Selection.Font.Size = 16
        Selection.ShrinkToFit = True
        Selection.Font.Color = -4165632
        
    ' Sum and format the project Non-participating totals
        ActiveCell.Offset(2, -8).Range("A1").Select
        ActiveCell.FormulaR1C1 = "PROJECT NON-PARTICIPATING TOTAL"
        Selection.Font.Color = -4165632
        Selection.Font.Bold = True
        Selection.Font.Size = 16
        ActiveCell.Offset(0, 7).Range("A1").Select
        ActiveCell.FormulaR1C1 = "=SUMIFS(C,C[41],""No"")"
        Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        Selection.Font.Color = -4165632
        Selection.Font.Bold = True
        Selection.ShrinkToFit = True
        Selection.Font.Size = 16
        ActiveCell.Offset(2, -8).Range("A1").Select
        

        
 End Sub
 
 Sub Organize_CAD_SuppCorrect()
    ActiveSheet.Select
    CAD_Selected = ActiveSheet.Name
    PAYITEMTYPE_SUPPLEMENTAL.Visible = True
    PAYITEMTYPE_SUPPLEMENTAL.Select
    Dim RowType As Range                            'Declare a variable for the Pay Item Type and the range of data from the Data Validation sheet
    For Each RowType In Range("A2:A51")            ' Maximum of 100 types the user can add to the program. 4 default types. 50 User Defined and 50 supplemental types.
                                                    ' The original estimates only go to A2 and A55
        If IsEmpty(RowType) = False Then
        
    ' Select pay items on Database
    Database.Select
    Database.Range("$A$1:$CS$999").AutoFilter
    Range("A2").Select

' Filter pay items to Participating Items'
    ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=47, Criteria1:=Array( _
        "Correction Non-Participating", "Correction Participating", "Yes", "No"), Operator:= _
        xlFilterValues
        
' Filter pay items to type'
    ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=46, Criteria1:=RowType.Value
    Range("A2").Select
    Selection.End(xlDown).Select
            If Selection <> "" Then
                Range("A2:C2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
            ' Pasting copied Pay Item, Item Description, and Contract quantity to the CAD
                Sheets(CAD_Selected).Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                ActiveCell.Offset(0, 5).Range("A1").Select
            ' Pasting copied Unit and Unit Price to the Sheets(CAD_Selected)
                Database.Select
                Range("D2:E2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
                Sheets(CAD_Selected).Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                ActiveCell.Offset(0, 3).Range("A1").Select
            ' Pasting copied Subtotal Contract Amount to the Sheets(CAD_Selected)
                Database.Select
                Range("F2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
                Sheets(CAD_Selected).Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                ActiveCell.Offset(0, 1).Range("A1").Select
             ' Pasting copied all other data to the Sheets(CAD_Selected)
                Database.Select
                Range("G2:CS2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
                Sheets(CAD_Selected).Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                    If IsEmpty(Selection.End(xlDown).Value) = True Then
                        ActiveCell.Offset(1, -1).Range("A1").Select
                        Else
                        Selection.End(xlDown).Select
                        ActiveCell.Offset(1, -1).Range("A1").Select
                    End If
                ' Dim RowTypeValue As String
                       ActiveCell.FormulaR1C1 = "=SUMIFS(C,C[40]," & Chr(34) & RowType.Value & Chr(34) & ")"
                    Call Format_SubtotalAmount_Participating
                ActiveCell.Offset(0, -7).Range("A1").Select
                ActiveCell.FormulaR1C1 = "Subtotal " & RowType.Value & " Items"
                    Call Format_SubtotalText_Participating
                ActiveCell.Offset(2, -1).Range("A1").Select
                
                End If
        End If
    Next RowType
    
    Sheets(CAD_Selected).Select
 End Sub
 
 
 Sub Finish_CAD()
        
        Sheets(CAD_Selected).Select
        
        ActiveCell.Offset(0, 1).Range("A1").Select
        ActiveCell.FormulaR1C1 = "FUEL & MATERIAL ADJUSTMENTS PARTICIPATING AND NON-PARTICIPATING TOTAL"
        Selection.Font.Bold = True
        Selection.Font.Size = 16
        ActiveCell.Offset(0, 8).Range("A1").Select
        ActiveCell.FormulaR1C1 = 0
        Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        Selection.Font.Bold = True
        Selection.ShrinkToFit = True
        Selection.Font.Size = 16
        ActiveCell.Offset(2, -8).Range("A1").Select         ' move into position for possible change order
        
        ActiveCell.FormulaR1C1 = "PROJECT PARTICIPATING AND NON-PARTICIPATING TOTAL"
        Selection.Font.Bold = True
        Selection.Font.Size = 16
        ActiveCell.Offset(0, 7).Range("A1").Select
        ActiveCell.FormulaR1C1 = "=SUMIFS(C,C[41],""Yes"")+SUMIFS(C,C[41],""No"")"
        Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        Selection.Font.Bold = True
        Selection.ShrinkToFit = True
        Selection.Font.Size = 16
        
        ActiveCell.Offset(4, -8).Range("A1").Select
                Call Organize_CAD_SuppCorrect               ' This was put right here so the corrections and supplemental items would show up after the totals.  Delineate them visually.
        ActiveCell.Offset(2, 0).Range("A1").Select
        ActiveCell.Value = "End of CAD document"            ' Used to find the end of the document. If a change order exist or not, use this is a reference point.
        ActiveCell.Offset(2, 0).Range("A1").Select
        
        
        
        Sheets(CAD_Selected).Select
        Rows("53:5000").EntireRow.AutoFit                       ' Autofit all the rows
    
End Sub



