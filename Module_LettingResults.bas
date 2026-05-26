Attribute VB_Name = "Module_LettingResults"

 Sub Form_LettingResults()

' Clear Letting results sheet
  Application.ScreenUpdating = False
    BIDTAB.Unprotect
    LettingResults.Select
    LettingResults.Unprotect
    Range("C3:C6").Clear
    Range("A7:Z999").Clear
    Range("J1:M30").Clear
    
    Range("C4").Select
    Selection.NumberFormat = "[$-en-US]mmmm d, yyyy;@"
    Selection.HorizontalAlignment = xlCenter
    Selection.VerticalAlignment = xlCenter
    Selection.HorizontalAlignment = xlCenter
    Selection.Font.Size = 22
    Selection.Locked = False
    Selection.Interior.Color = 13434879

    LettingResults.Range("C5").Value = BIDTAB.Range("B8").Value             'County copied from Bid Tab
    Range("C5").Select
    Selection.Font.Size = 22
    Selection.HorizontalAlignment = xlCenter
    Selection.VerticalAlignment = xlCenter
    Selection.ShrinkToFit = True
    
    LettingResults.Range("C6").Value = BIDTAB.Range("D6").Value             'Project No copied from Bid Tab
    Range("C6").Select
    Selection.Font.Size = 22
    Selection.HorizontalAlignment = xlCenter
    Selection.VerticalAlignment = xlCenter

'Select Bid Results summary from Columns GA to GC at the top rows 1-15
    BIDTAB.Select
    Range("DA2:DD15").Select
    Selection.Copy
    
    LettingResults.Select
    
    Range("J1").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("L:L").Select
    Selection.Delete Shift:=xlToLeft
    Columns("J:L").Select
    ActiveSheet.Range("$J$1:$L$999").AutoFilter Field:=1, Criteria1:="<>"
    
    Columns("J:L").Select
    ActiveWorkbook.Worksheets("LettingResults").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("LettingResults").Sort.SortFields.Add2 Key:=Range( _
        "J2:J999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("LettingResults").Sort
        .SetRange Range("J1:L999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("J1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
        Range("B9").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.AutoFilter

        Range("B9:D9").Select           ' Select and format heading
        Selection.Font.Bold = True
        Call Format_LettingResults
        
        Range("B10").Select             ' Select and format Ranking Order
            If IsEmpty(Selection.End(xlDown).Value) = True Then
                ActiveCell.Select
                Else
                Range(Selection, Selection.End(xlDown)).Select
            End If
            Call Format_LettingResults
        
        Range("C10").Select             ' Select and format Contractor
            If IsEmpty(Selection.End(xlDown).Value) = True Then
                ActiveCell.Select
                Else
                Range(Selection, Selection.End(xlDown)).Select
            End If
            Selection.ShrinkToFit = False
            Selection.WrapText = True
            Call Format_LettingResults
        
        Range("D10").Select             ' Select and format Bid Amount
            If IsEmpty(Selection.End(xlDown).Value) = True Then
                ActiveCell.Select
                Else
                Range(Selection, Selection.End(xlDown)).Select
            End If
            Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            Call Format_LettingResults

    Range("B9").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(2, 0).Range("A1").Select

    Call STEP_2_LETTINGRESULTS
    Call Format_PageSetup
End Sub
'
'
'*************************************************************** BID TAB - STEP 2 Letting Results ***********************************************************************************************
'
'



'
'
''''''''''''Step 2 Letting Results''''''''''''''''''''''''''''
'
'

 Sub STEP_2_LETTINGRESULTS()
    ActiveCell.Offset(2, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Preliminary Estimate:"
    Selection.Font.Size = 22
    Selection.ShrinkToFit = True
    Selection.HorizontalAlignment = xlCenter
    ActiveCell.Offset(0, 1).Range("A1").Select


    BIDTAB.Select
    BIDTAB.Unprotect
    Range("A1").Select
   '  Find Total for 2 scenarios
    Dim finditemtype1 As Range
    Set finditemtype1 = ThisWorkbook.Worksheets("BIDTAB").Cells.Find(What:="PROJECT PARTICIPATING AND NON-PARTICIPATING TOTAL", After:=ActiveCell, LookIn:=xlValues, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False)
        If (Not finditemtype1 Is Nothing) Then
            finditemtype1.Select
            ActiveCell.Offset(0, 4).Range("A1").Select
            Selection.Copy
            LettingResults.Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
                Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                Selection.HorizontalAlignment = xlCenter
                Selection.Font.Size = 22
                Selection.ShrinkToFit = True
        Else
           ' Cells.Find(What:="PROJECT PARTICIPATING TOTAL", After:=ActiveCell, LookIn:=xlValues _
            , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
        End If
    
    ActiveCell.Offset(1, -1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Percent Plus or Minus:"
    Selection.HorizontalAlignment = xlCenter
    ActiveCell.Offset(0, 1).Range("A1").Select
    
    
    
    BIDTAB.Select
'  Find Percent Over / Under for 2 scenarios
    Dim finditemtype2 As Range
    Set finditemtype2 = ThisWorkbook.Worksheets("BIDTAB").Cells.Find(What:="PERCENT OVER / UNDER PARTICIPATING AND NON-PARTICIPATING OSARC ESTIMATE", After:=ActiveCell, LookIn:=xlValues, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False)
        If (Not finditemtype2 Is Nothing) Then
            finditemtype2.Select
            ActiveCell.Offset(0, 104).Range("A1").Select
            Selection.Copy
            LettingResults.Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
                Selection.NumberFormat = "0.0%"
                Selection.HorizontalAlignment = xlCenter
                Selection.Font.Size = 22
                Selection.ShrinkToFit = True
                Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                  Formula1:="=0"
                Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
                With Selection.FormatConditions(1).Font
                    .Color = -16383844
                    .TintAndShade = 0
                End With
            Selection.FormatConditions(1).StopIfTrue = False
        Else
           ' Cells.Find(What:="PERCENT OVER / UNDER PARTICIPATING", After:=ActiveCell, LookIn:=xlValues _
            , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
        End If
        
    ActiveCell.Offset(1, -1).Range("A1").Select
    ActiveCell.FormulaR1C1 = BIDTAB.Range("DF12").Value                                                  ' Status of the Award
    Selection.HorizontalAlignment = xlCenter
    Selection.Font.Size = 22
    Selection.ShrinkToFit = True
    Selection.Font.Bold = True
        Selection.FormatConditions.Add Type:=xlTextString, String:="AWARDED", _
        TextOperator:=xlContains
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        Selection.FormatConditions(1).Font.Color = -11489280
        Selection.FormatConditions(1).StopIfTrue = False
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Awarded on "
    Selection.HorizontalAlignment = xlCenter
    Selection.Font.Size = 22
    Selection.ShrinkToFit = True
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = BIDTAB.Range("DF13").Value                                                  ' Date of the Award
        Selection.NumberFormat = "[$-en-US]mmmm d, yyyy;@"
        Selection.Font.Size = 22
        Selection.ShrinkToFit = True
        Selection.HorizontalAlignment = xlCenter

    
'Setting the print area to find the row with "Awarded on ***Month XX, YEAR***" so that the page prints to the last row
    Dim ws As Worksheet
    Dim lastRow As Long
    Set ws = LettingResults
        lastRow = ws.Cells.Find("Awarded on ", SearchOrder:=xlByRows, SearchDirection:=xlNext).row      ' find the last row with formatting, to be included in print range
        ws.PageSetup.PrintArea = ws.Range("A1:D" & lastRow).Address
'Protecting and finishing up
    BIDTAB.Protect
    LettingResults.Select
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Rows("4:999").Select
    Rows("4:999").EntireRow.AutoFit
    LettingResults.Protect
    Application.ScreenUpdating = True
    LettingResults.Select
    Range("C4").Select
End Sub



