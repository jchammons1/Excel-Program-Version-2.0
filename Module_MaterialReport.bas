Attribute VB_Name = "Module_MaterialReport"
Sub MaterialsReport()
    Material_Report.Visible = True  ' Only adding this here because previous version the sheet will be hidden after it leaves here.  This avoids having
                                    ' to tell people to unhide the sheet
    Application.ScreenUpdating = False
   
'Create and Modify Engineer's Estimate'
'Sort the entire sheet DataEngEst Database so pay items can be organized into the Engineer's Estimate'
    Material_Report.Select
    Material_Report.Unprotect
    ActiveSheet.ResetAllPageBreaks
    Range("A7:CZ999").Clear

'Sort the the Project Pay Item Database'
    Database.Visible = True
    Database.Unprotect
    Call SortPayItems

' Organize Pay Items
    Material_Report.Select
    Material_Report.Range("A7").Select
    Call OrganizePayitems_Materials
    
'Determine if there are any Non-Participating Pay Items.  If not, then skip the Call
            Database.Select
            Database.Range("$A$1:$CS$999").AutoFilter
                ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=47, Criteria1:="No"
                Range("A2").Select
                Selection.End(xlDown).Select
            If Selection <> "" Then
                Material_Report.Select
                ActiveCell.Offset(2, 0).Range("A1").Select
                Call OrganizePayItems_MaterialsNP
                Database.Range("$A$1:$CS$999").AutoFilter
                Else
                Material_Report.Select
                Database.Range("$A$1:$CS$999").AutoFilter
            End If
    
    Material_Report.Select
    ActiveCell.Select
    
    Material_Report.Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "End of the Materials Summary Report"
    ActiveCell.Select
    Selection.Font.Bold = True
    
'Do Loop to format the columns
 Call Format_MaterialsReport
    
    Range("A7").Select
    Cells.Find(What:="End of the Materials Summary Report", After:=ActiveCell, _
            LookIn:=xlFormulas2, LookAt:=xlPart, SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Offset(1, 0).Range("A1").Select

 
'Creates the page divides for each pay item for testing
 Call Dividers_MaterialsReport

'Setting the print area to find the row with "TOTAL BID AMOUNT" so that the page prints to the last row
      Dim ws As Worksheet
      Dim lastRow As Long
    
      Set ws = Material_Report
      lastRow = ws.Cells.Find("THE END", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row      ' find the last row with formatting, to be included in print range
      ws.PageSetup.PrintArea = ws.Range("A1:G" & lastRow).Address
    
' Remove filters from the database.  The pay item form will not add pay items if this sheet filters are not removed'
    Database.Select
    Selection.AutoFilter
    Database.Protect
    Database.Visible = False
    PAYITEMTYPE.Visible = False
    Material_Report.Select
    Rows("7:999").Select
    Selection.RowHeight = 30
   
    With Material_Report
        Format_PageSetup
    End With

    Call SetPrintAreaToColumnG_LastRow
    
    Material_Report.Protect
    Application.ScreenUpdating = True
    Range("A1").Select
End Sub


'
'
'************************************************************  MATERIALS REPORT END **********************************************************************
'
'
'
'
 Sub OrganizePayitems_Materials()
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
                Range("A2:C2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
            ' Pasting copied participating and non-participating pay items
                Material_Report.Select
                ActiveCell.FormulaR1C1 = RowType.Value & " Participating Items:"
                ActiveCell.Offset(1, 0).Range("A1").Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            ' Return to the database and copy the quantity and units
                Database.Select
                Range("D2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
            ' Pasting copied Participating Pay Items status
                Material_Report.Select
                ActiveCell.Offset(0, 4).Range("A1").Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                Database.Select
                Range("AU2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
            ' Pasting copied Participating Pay Items status
                Material_Report.Select
                ActiveCell.Offset(0, 3).Range("A1").Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False

                    If Selection.End(xlDown).Value <> "" Then
                        Selection.End(xlDown).Select
                    Else
                    End If
                ActiveCell.Offset(1, -7).Range("A1").Select
                End If
        End If
    Next RowType
 
    Database.Range("$A$1:$CZ$999").AutoFilter
End Sub


 Sub OrganizePayItems_MaterialsNP()
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
                Range("A2:C2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
            ' Pasting copied non-participating pay items
                Material_Report.Select
                ActiveCell.FormulaR1C1 = RowType.Value & " Non-Participating Items:"
                ActiveCell.Offset(1, 0).Range("A1").Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            ' Return to the database and copy the quantity and units
                Database.Select
                Range("D2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
            ' Pasting copied Participating Pay Items status
                Material_Report.Select
                ActiveCell.Offset(0, 3).Range("A1").Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                Database.Select
                Range("AU2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
            ' Pasting copied Participating Pay Items status
                Material_Report.Select
                ActiveCell.Offset(0, 4).Range("A1").Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                    
                    If Selection.End(xlDown).Value <> "" Then
                        Selection.End(xlDown).Select
                    Else
                    End If
                ActiveCell.Offset(1, -7).Range("A1").Select
                End If
        End If
    Next RowType
 
    Database.Range("$A$1:$CZ$999").AutoFilter
End Sub


'''******************************************************************************
'################################################################################

Sub Dividers_MaterialsReport()

   Material_Report.Select
   ActiveCell.Offset(1, 0).Range("A1").Select
   
    Database.Select
    Cells.Select
    
   Range("A1").Select
            row = 1
            endrow = CountPayItems - 25 - 21
            
    Do While row < endrow
        Database.Select
        Range("A" & row).Select
        ' Copy the first part of the pay item description
        Range(ActiveCell.Offset(1, 0).Range("A1"), ActiveCell.Offset(1, 0).Range("F1")).Select
        Selection.Copy
        Material_Report.Select
            Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=True
        Selection.End(xlDown).Select
        ActiveCell.Offset(1, 0).Range("A1").Select
        
        ' Copy the second part for Type and Participating
        Database.Select
        Range("A1").Select
        Range(ActiveCell.Offset(1, 0).Range("AT1"), ActiveCell.Offset(1, 0).Range("AU1")).Select
        Selection.Copy
        Material_Report.Select
            Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=True
        Range(Selection, Selection.End(xlUp)).Select
            Call Format_MaterialDivider
            Call Format2_MaterialDivider

            row = row + 1
    Loop
    
    Material_Report.Select
    ActiveCell.Offset(-2, 7).Range("A1").Select
    ActiveCell.FormulaR1C1 = "THE END"
    

End Sub

Sub Format_MaterialDivider()
    Selection.Font.Name = "Calibri"
    Selection.Font.Size = 36
    Selection.VerticalAlignment = xlCenter
    Selection.HorizontalAlignment = xlCenter
    Selection.ShrinkToFit = True
End Sub

Sub Format2_MaterialDivider()
    ActiveCell.Offset(0, 0).Range("A1").Select
    ActiveCell.Range("A1:G1").Select
    Selection.Merge
    Call Format_Lines
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.Range("A1:G1").Select
    Selection.Merge
    Call Format_Lines
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.Range("A1:G1").Select
    Selection.Merge
    Call Format_Lines
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.Range("A1:G1").Select
    Selection.Merge
    Call Format_Lines
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.Range("A1:G1").Select
    Selection.Merge
    Call Format_Lines
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.Range("A1:G1").Select
    Selection.Merge
    Call Format_Lines
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.Range("A1:G1").Select
    Selection.Merge
    Call Format_Lines
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.Range("A1:G1").Select
    Selection.Merge
    Call Format_Lines
    ''' End of data on to the next record
    
    ActiveCell.Offset(2, 0).Range("A1").Select
    Call InsertPageBreakAtActiveRow
    
End Sub

Sub Format_Lines()
    Selection.RowHeight = 60
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColor = 0
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
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
End Sub

Sub InsertPageBreakAtActiveRow()
    On Error Resume Next
    ActiveCell.EntireRow.PageBreak = xlPageBreakManual
    On Error GoTo 0
End Sub


Sub SetPrintAreaToColumnG_LastRow()



    With ActiveSheet
        .PageSetup.PrintArea = "$A:$G"
        .PageSetup.FitToPagesWide = 1
    End With


End Sub
    




