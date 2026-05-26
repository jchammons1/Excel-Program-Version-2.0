Attribute VB_Name = "Module_DatabaseUpdate"
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'*********************************** This section automatically takes the winning bidder unit cost and updates the Database *************>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'
'
'

Sub DatabaseUnitCostUpdate()
Attribute DatabaseUnitCostUpdate.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Copy data from the BidTab to the DatabaseTemp
'
Application.ScreenUpdating = False
    AvgUnitCost.Visible = True
    AvgUnitCost.Select
    AvgUnitCost.Unprotect
    AvgUnitCost.Range("A2:DB9999").Clear
    AvgUnitCost.Range("A2").Select
    
    BIDTAB.Select
    BIDTAB.Range("A21:DB999").Select
    
    Selection.Copy
    AvgUnitCost.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Columns("CT:CZ").Select
    Selection.Delete Shift:=xlToLeft
    Range("CT2:CU999").Select


    Selection.Cut           ' Cut Column CT winning bid unit cost
    Range("E2").Select
    ActiveSheet.Paste       ' Paste data to Column E
    
    Cells.Select
    Selection.AutoFilter    ' Filter out the non payitems rows
    ActiveSheet.Range("$A$1:$DA$9999").AutoFilter Field:=1, Criteria1:=Array( _
        "APPROVED:", "Signature and Date", "="), Operator:=xlFilterValues
'
'
' Delete filtered rows
'
'
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("AvgUnitCost") ' Change "Sheet1" to your sheet name

    ' Set the range to the used range excluding the header row
    Set rng = ws.UsedRange.Offset(1, 0).SpecialCells(xlCellTypeVisible)

    ' Loop through the range and delete visible rows
    For Each cell In rng
        If cell.EntireRow.Hidden = False Then
            cell.EntireRow.Delete
        End If
    Next cell
    
    Cells.Select
    Selection.AutoFilter

    Database.Visible = True
    Database.Unprotect
    Database.Range("A2:CS9999").Clear
    Database.Select
    Database.Range("A2").Select
    
    AvgUnitCost.Select
    AvgUnitCost.Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Database.Select
    Database.Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Call AverageUnitCost
    
    AvgUnitCost.Protect
    AvgUnitCost.Visible = False

    Database.Protect
    Database.Visible = False
    Application.ScreenUpdating = True
    BIDTAB.Select
    BIDTAB.Range("B4").Select

End Sub

