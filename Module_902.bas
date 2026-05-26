Attribute VB_Name = "Module_902"
Option Explicit
Dim RowType As Range


'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'*********************************** This section contains Form 902 to create the form for bidders to enter their bid amounts ***********>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
 Sub Form_902()
    

    Form902.Select
    Range("AA1").Select
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(""For the construction of:  "",Project_Data!R2C5,"" on "",TEXT(Project_Data!R2C17,""0.000""),"" mile(s) of county road(s) known as "",Project_Data!R2C1,"" "",Project_Data!R2C2,"" in "",Project_Data!R2C4, "" County, Mississippi under "",Project_Data!R2C12,"" Project No. "",Project_Data!R2C13)"
    Range("AA2").Select
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(Project_Data!R2C4,"" County, Mississippi"")"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "=Project_Data!R2C13"

    Range("AB1").Select
    Range("AB1") = _
        "=CONCATENATE(""I (We) agree to complete the entire project within "",Project_Data!R2C10,"" Working Days."")"

    Range("B4") = Range("AA1").Value
    Range("B4:F5").Select
           With Selection
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlTop
                .WrapText = True
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = True
            End With
            With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 13434879
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
            End With
            Selection.Locked = False
    
    Range("G3") = Range("AA2").Value
    Range("E3:G3").Select
           With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlTop
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = True
                .ReadingOrder = xlContext
                .MergeCells = True
            End With
    Range("B6") = Range("AB1").Value
    Range("B6:F7").Select
           With Selection
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlTop
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = True
                .ReadingOrder = xlContext
                .MergeCells = True
            End With
            With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 13434879
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
            End With
            Selection.Locked = False
            
' Organize Pay Items
    Form902.Range("A13").Select
    Call OrganizePayitems_902
    
'Determine if there are any Non-Participating Pay Items.  If not, then skip the Call
            Database.Select
            Database.Range("$A$1:$CS$999").AutoFilter
                ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=47, Criteria1:="No"
                Range("A2").Select
                Selection.End(xlDown).Select
            If Selection <> "" Then
                Form902.Select
                ActiveCell.Offset(1, 0).Range("A1").Select
                Call OrganizePayItems_902NP
                Database.Range("$A$1:$CS$999").AutoFilter
                Else
                Form902.Select
                Database.Range("$A$1:$CS$999").AutoFilter
            End If
            
            
    Form902.Select
    ActiveCell.Offset(4, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "TOTAL BID AMOUNT"
    ActiveCell.Select
    Selection.Font.Bold = True
    ActiveCell.Offset(0, 0).Range("A1:E1").Select
    Selection.Merge
    Selection.HorizontalAlignment = xlRight
    
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.Offset(0, 0).Range("A1:B1").Select
    Selection.Merge
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
'Do Loop to format the columns
 Call Format_902
   
'Setting the print area to find the row with "TOTAL BID AMOUNT" so that the page prints to the last row
      Dim ws As Worksheet
      Dim LastRow As Long
    
      Set ws = Form902
      LastRow = ws.Cells.Find("TOTAL BID AMOUNT", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row      ' find the last row with formatting, to be included in print range
      ws.PageSetup.PrintArea = ws.Range("A1:G" & LastRow).Address
    
' Remove filters from the database.  The pay item form will not add pay items if this sheet filters are not removed'
    Database.Select
    Cells.Select
    Selection.AutoFilter
    PAYITEMTYPE.Visible = False
    
    Form902.Select
    Rows("13:999").Select
    Selection.RowHeight = 30
    Call Format_PageSetup
    Call Flag_618B_803B_Form902
    
End Sub

'
'
'
 Sub OrganizePayitems_902()
    PAYITEMTYPE.Visible = True
    PAYITEMTYPE.Select
    Dim RowType As Range                         ' Declare a variable for the Pay Item Type and the range of data from the Data Validation sheet
    For Each RowType In Range("A2:A55")          ' Maximum of 50 types the user can add to the program.  4 default types
       
        If IsEmpty(RowType) = False Then
        
' Select pay items on Database
    Database.Visible = True
    Database.Select
    Database.Unprotect
    Database.Range("$A$1:$CS$999").AutoFilter
    Range("A2").Select
' Filter pay items to Participating Items'
    ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=47, Criteria1:="Yes"
' Filter pay items to type'
    ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=46, Criteria1:=RowType.Value
    Range("A2").Select
    Selection.End(xlDown).Select
            If Selection <> "" Then
                Range("A2:D2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
            ' Pasting copied participating and non-participating pay items
                Form902.Select
                ActiveCell.FormulaR1C1 = RowType.Value & " Participating Items:"
                ActiveCell.Offset(1, 1).Range("A1").Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            ' Return to the database and copy the quantity and units
                Database.Select
                Range("AU2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
            ' Pasting copied Participating Pay Items status
                Form902.Select
                ActiveCell.Offset(0, 6).Range("A1").Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                    If Selection.End(xlDown).Value <> "" Then
                        Selection.End(xlDown).Select
                            ActiveCell.Offset(1, -1).Range("A1").Select
                            ' Format box
                                Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
                                Selection.Borders(xlEdgeLeft).Weight = xlThin
                                Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
                                Selection.Borders(xlEdgeTop).Weight = xlThin
                                Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
                                Selection.Borders(xlEdgeBottom).Weight = xlThin
                                Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
                                Selection.Borders(xlEdgeRight).Weight = xlThin
                                Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
                                Selection.Borders(xlInsideVertical).Weight = xlThin
                                Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                                Selection.Borders(xlInsideHorizontal).Weight = xlThin
                        ActiveCell.Offset(0, -5).Range("A1").Select
                        ActiveCell.FormulaR1C1 = "Subtotal " & RowType.Value & " Participating Items"
                            ' Format
                                Selection.Font.Bold = True
                                ActiveCell.Offset(0, 0).Range("A1:E1").Select
                                Selection.Merge
                                Selection.HorizontalAlignment = xlRight
                                ActiveCell.Offset(2, -1).Range("A1").Select
                    ElseIf ActiveCell.Value <> "" Then
                            ActiveCell.Offset(1, -1).Range("A1").Select
                            ' Format box
                                Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
                                Selection.Borders(xlEdgeLeft).Weight = xlThin
                                Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
                                Selection.Borders(xlEdgeTop).Weight = xlThin
                                Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
                                Selection.Borders(xlEdgeBottom).Weight = xlThin
                                Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
                                Selection.Borders(xlEdgeRight).Weight = xlThin
                                Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
                                Selection.Borders(xlInsideVertical).Weight = xlThin
                                Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                                Selection.Borders(xlInsideHorizontal).Weight = xlThin
                        ActiveCell.Offset(0, -5).Range("A1").Select
                        ActiveCell.FormulaR1C1 = "Subtotal " & RowType.Value & " Participating Items"
                            ' Format
                                Selection.Font.Bold = True
                                ActiveCell.Offset(0, 0).Range("A1:E1").Select
                                Selection.Merge
                                Selection.HorizontalAlignment = xlRight
                                ActiveCell.Offset(2, -1).Range("A1").Select
                    Else
                       ActiveCell.Offset(2, -7).Range("A1").Select
                    End If
                End If
        End If
    Next RowType
 
    Database.Range("$A$1:$CZ$999").AutoFilter
End Sub


 Sub OrganizePayItems_902NP()
    PAYITEMTYPE.Select
    Dim RowType As Range                         ' Declare a variable for the Pay Item Type and the range of data from the Data Validation sheet
    For Each RowType In Range("A2:A55")          ' Maximum of 50 types the user can add to the program.  4 default types
       
        If IsEmpty(RowType) = False Then
        
' Select pay items on Database
    Database.Select
    Database.Range("$A$1:$CZ$999").AutoFilter
    Range("A2").Select
' Filter pay items to Participating Items'
    ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=47, Criteria1:="No"
' Filter pay items to type'
    ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=46, Criteria1:=RowType.Value
    Range("A2").Select
    Selection.End(xlDown).Select
            If Selection <> "" Then
                Range("A2:D2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
            ' Pasting copied non-participating pay items
                Form902.Select
                ActiveCell.FormulaR1C1 = RowType.Value & " Non-Participating Items:"
                ActiveCell.Offset(1, 1).Range("A1").Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            ' Return to the database and copy the quantity and units
                Database.Select
                Range("AU2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
            ' Pasting copied Participating Pay Items status
                Form902.Select
                ActiveCell.Offset(0, 6).Range("A1").Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                    If Selection.End(xlDown).Value <> "" Then
                        Selection.End(xlDown).Select
                        ActiveCell.Offset(1, -1).Range("A1").Select
                            ' Format box
                                Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
                                Selection.Borders(xlEdgeLeft).Weight = xlThin
                                Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
                                Selection.Borders(xlEdgeTop).Weight = xlThin
                                Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
                                Selection.Borders(xlEdgeBottom).Weight = xlThin
                                Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
                                Selection.Borders(xlEdgeRight).Weight = xlThin
                                Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
                                Selection.Borders(xlInsideVertical).Weight = xlThin
                                Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                                Selection.Borders(xlInsideHorizontal).Weight = xlThin
                        ActiveCell.Offset(0, -5).Range("A1").Select
                        ActiveCell.FormulaR1C1 = "Subtotal " & RowType.Value & " Non-Participating Items"
                            ' Format
                                Selection.Font.Bold = True
                                ActiveCell.Offset(0, 0).Range("A1:E1").Select
                                Selection.Merge
                                Selection.HorizontalAlignment = xlRight
                                ActiveCell.Offset(2, -1).Range("A1").Select
                    ElseIf ActiveCell.Value <> "" Then
                        ActiveCell.Offset(1, -1).Range("A1").Select
                            ' Format box
                                Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
                                Selection.Borders(xlEdgeLeft).Weight = xlThin
                                Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
                                Selection.Borders(xlEdgeTop).Weight = xlThin
                                Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
                                Selection.Borders(xlEdgeBottom).Weight = xlThin
                                Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
                                Selection.Borders(xlEdgeRight).Weight = xlThin
                                Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
                                Selection.Borders(xlInsideVertical).Weight = xlThin
                                Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                                Selection.Borders(xlInsideHorizontal).Weight = xlThin
                        ActiveCell.Offset(0, -5).Range("A1").Select
                        ActiveCell.FormulaR1C1 = "Subtotal " & RowType.Value & " Non-Participating Items"
                            ' Format
                                Selection.Font.Bold = True
                                ActiveCell.Offset(0, 0).Range("A1:E1").Select
                                Selection.Merge
                                Selection.HorizontalAlignment = xlRight
                                ActiveCell.Offset(2, -1).Range("A1").Select

                    Else
                       ActiveCell.Offset(2, -7).Range("A1").Select
                    End If
                End If
        End If
    Next RowType
 
    Database.Range("$A$1:$CZ$999").AutoFilter
End Sub


Sub Box_902()
    ActiveCell.Offset(1, -1).Range("A1").Select
        ' Format box
            Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
            Selection.Borders(xlEdgeLeft).Weight = xlThin
            Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
            Selection.Borders(xlEdgeTop).Weight = xlThin
            Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
            Selection.Borders(xlEdgeBottom).Weight = xlThin
            Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
            Selection.Borders(xlEdgeRight).Weight = xlThin
            Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
            Selection.Borders(xlInsideVertical).Weight = xlThin
            Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            Selection.Borders(xlInsideHorizontal).Weight = xlThin
    ActiveCell.Offset(0, -5).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Subtotal " & RowType.Value & " Participating Items"
        ' Format
            Selection.Font.Bold = True
            ActiveCell.Offset(0, 0).Range("A1:E1").Select
            Selection.Merge
            Selection.HorizontalAlignment = xlRight
            ActiveCell.Offset(2, -1).Range("A1").Select
End Sub

Sub Box_902_NP()
    ActiveCell.Offset(1, -1).Range("A1").Select
        ' Format box
            Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
            Selection.Borders(xlEdgeLeft).Weight = xlThin
            Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
            Selection.Borders(xlEdgeTop).Weight = xlThin
            Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
            Selection.Borders(xlEdgeBottom).Weight = xlThin
            Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
            Selection.Borders(xlEdgeRight).Weight = xlThin
            Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
            Selection.Borders(xlInsideVertical).Weight = xlThin
            Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            Selection.Borders(xlInsideHorizontal).Weight = xlThin
    ActiveCell.Offset(0, -5).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Subtotal " & RowType.Value & " Non-Participating Items"
        ' Format
            Selection.Font.Bold = True
            ActiveCell.Offset(0, 0).Range("A1:E1").Select
            Selection.Merge
            Selection.HorizontalAlignment = xlRight
            ActiveCell.Offset(2, -1).Range("A1").Select
End Sub
