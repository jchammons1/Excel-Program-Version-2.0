Attribute VB_Name = "Module_Estimates"
Option Explicit
    Dim ESTIMATE_SELECTED As String
    

'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'*********************************** This section contains the Estimates.  This estimate section is used for the ************************>>>>>>>>>>>>>>>>>>
'*********************************** Engineers, OSARC Preliminary, Contractor and Bid Tab Estiamtes *************************************>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'
'
'

'Check to see if the cell
'Create and Modify Estimate'
 Sub Estimates()
    Database.Visible = True
    Database.Unprotect
    PAYITEMTYPE.Visible = True
    
    ESTIMATE_SELECTED = ActiveSheet.Name
    Sheets(ESTIMATE_SELECTED).Unprotect
    Sheets(ESTIMATE_SELECTED).Select
    Range("A21:CZ999").Clear    ' Most estimates only need to clear to Column "CS" but the BidTab needs to clear to Column "CZ"
    Range("A21").Select
        
        Call SortPayItems
        Sheets(ESTIMATE_SELECTED).Select
        Call OrganizePayItems
    'Determine if there are any Non-Participating Pay Items.  If not, then skip the Call
            Database.Select
            Database.Range("$A$1:$CS$999").AutoFilter
                ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=47, Criteria1:="No"
                Range("A2").Select
                Selection.End(xlDown).Select
            If Selection <> "" Then
                Sheets(ESTIMATE_SELECTED).Select
                ActiveCell.Offset(0, -1).Range("A1").Select     ' move 1 space to the left because of the stop location from participating total
                Call OrganizePayItems_NP
                Database.Range("$A$1:$CS$999").AutoFilter
                Else
                Sheets(ESTIMATE_SELECTED).Select
                Database.Range("$A$1:$CS$999").AutoFilter
            End If
            
    ' Sum and format accumulative totals
            Sheets(ESTIMATE_SELECTED).Select
            ActiveCell.FormulaR1C1 = "PROJECT PARTICIPATING AND NON-PARTICIPATING TOTAL"
            Selection.Font.Bold = True
            Selection.Font.Size = 16
            ActiveCell.Offset(0, 4).Range("A1").Select
            ActiveCell.FormulaR1C1 = "=SUMIFS(C,C[41],""Yes"")+SUMIFS(C,C[41],""No"")"
            Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            Selection.Font.Bold = True
            Selection.ShrinkToFit = True
            Selection.Font.Size = 16
            
            
     ' If the BidTab Estimate, then insert the following extra calculations
        If ESTIMATE_SELECTED = "BIDTAB" Then
            ActiveCell.Offset(2, -4).Range("A1").Select
            ActiveCell.FormulaR1C1 = "PERCENT OVER / UNDER PARTICIPATING AND NON-PARTICIPATING OSARC ESTIMATE"
                    ActiveCell.Offset(2, 0).Range("A1").Select
            ActiveCell.FormulaR1C1 = "DIFFERENCE PARTICIPATING AND NON-PARTICIPATING BIDDER AMOUNT"
                    ActiveCell.Offset(2, 0).Range("A1").Select
            ActiveCell.FormulaR1C1 = "PERCENT OVER / UNDER PARTICIPATING AND NON-PARTICIPATING BIDDER AMOUNT"
            ActiveCell.Offset(3, 0).Range("A1").Select
                Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
                Selection.Borders(xlEdgeBottom).Weight = xlThin
            ActiveCell.Offset(0, -1).Range("A1").Select
            ActiveCell.FormulaR1C1 = "Signature and Date"
            ActiveCell.Offset(1, 0).Range("A1").Select
            ActiveCell.FormulaR1C1 = "APPROVED:"
            ActiveCell.Offset(0, 1).Range("A1").Select
            ActiveCell.FormulaR1C1 = "=CONCAT(R18C2, "" P.E."","", "",R18C4)"
            ActiveCell.Offset(1, -1).Range("A1").Select
            Else
          End If

     ' Format pay items in the estimate
        Call Format_Estimates
        Sheets(ESTIMATE_SELECTED).Select
        
    Database.Range("$A$1:$CS$999").AutoFilter
    Database.Protect
    PAYITEMTYPE.Visible = False
    Database.Visible = False
End Sub



'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'*********************************** This section organizes the pay items by type********************************************************>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'
' This loops through each row on the Pay Item Type sheet form Row 2 to 55 for all the types. For estimates, I do not want to inlcude the supplemental pay items
' The supplemental pay items are only added after construction has started based on an approved change order
'
 Sub OrganizePayItems()
  '  Dim EstimateSelected As String
  '  EstimateSelected = ActiveSheet.Name
    
    PAYITEMTYPE.Select
    Dim RowType As Range                         ' Declare a variable for the Pay Item Type and the range of data from the Data Validation sheet
    For Each RowType In Range("A2:A55")          ' Maximum of 50 types the user can add to the program.  4 default types
       
        If IsEmpty(RowType) = False Then
        
' Select pay items on Database
    Database.Select
    Database.Range("$A$1:$CS$999").AutoFilter
    Range("A2").Select
' Filter pay items to Participating Items'
    ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=47, Criteria1:="Yes"
' Filter pay items to type'
    ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=46, Criteria1:=RowType.Value
    Range("A2").Select
    Selection.End(xlDown).Select
            If Selection <> "" Then
                Range("A2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Range(Selection, Selection.End(xlToRight)).Select
                Selection.Copy
            ' Pasting copied Participating Pay Items to the Engineer's Estimate'
                Sheets(ESTIMATE_SELECTED).Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                    If Selection.End(xlDown).Value <> "" Then
                        Selection.End(xlDown).Select
                    Else
                    End If
                ActiveCell.Offset(1, 1).Range("A1").Select
                ActiveCell.FormulaR1C1 = "Subtotal Participating " & RowType.Value & " Items"
                ActiveCell.Select
                Selection.Font.Bold = True
                Selection.WrapText = True
                Selection.Font.Size = 14
                ActiveCell.Offset(0, 4).Range("A1").Select
            ' Dim RowTypeValue As String
                ActiveCell.FormulaR1C1 = "=SUMIFS(C,C[40]," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C[41],""Yes"")"
                ActiveCell.Select
                Selection.Font.Bold = True
                Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                Selection.Font.Size = 14
                ActiveCell.Offset(2, -5).Range("A1").Select
                End If
        End If
    Next RowType
    
    Sheets(ESTIMATE_SELECTED).Select
    ' Sum and format the project participating totals
        ActiveCell.Offset(0, 1).Range("A1").Select
        ActiveCell.FormulaR1C1 = "PROJECT PARTICIPATING TOTAL"
        Selection.Font.Bold = True
        Selection.Font.Size = 16
        ActiveCell.Offset(0, 4).Range("A1").Select
        ActiveCell.FormulaR1C1 = "=SUMIFS(C,C[41],""Yes"")"
        Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        Selection.Font.Bold = True
        Selection.ShrinkToFit = True
        Selection.Font.Size = 16
        ActiveCell.Offset(2, -4).Range("A1").Select
        
    ' If the OSARC or Contractors Estimate, then insert the following extra calculations
        If ESTIMATE_SELECTED = "OSARC" Or ESTIMATE_SELECTED = "CONTRACTOR" Then
                ActiveCell.FormulaR1C1 = "PARTICIPATING CONTINGENCIES"
                ActiveCell.Offset(0, 4).Range("A1").Select
                ActiveCell.FormulaR1C1 = "=R[-2]C*0.05"
                Selection.Font.Size = 16
                Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                Selection.ShrinkToFit = True
                ActiveCell.Offset(2, -4).Range("A1").Select
                ActiveCell.FormulaR1C1 = "PROJECT PARTICIPATING TOTAL plus (+) CONTINGENCIES"
                ActiveCell.Offset(0, 4).Range("A1").Select
                ActiveCell.FormulaR1C1 = "=R[-4]C+R[-2]C"
                Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                Selection.Font.Size = 16
                Selection.ShrinkToFit = True
                ActiveCell.Offset(2, -4).Range("A1").Select
                
                ActiveCell.FormulaR1C1 = "PROJECT PARTICIPATING TOTAL plus (+) CONTINGENCIES rounded to the nearest $100"  ' Added this line on 1/15/26
                ActiveCell.Offset(0, 4).Range("A1").Select
                ActiveCell.FormulaR1C1 = "=ROUND(R[-2]C/100,0)*100"
                Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                Selection.Font.Size = 16
                Selection.ShrinkToFit = True
                ActiveCell.Offset(2, -5).Range("A1").Select
                
                
                ActiveCell.FormulaR1C1 = "=CONCATENATE(R5C3,"" "",""FUNDS REQUESTED"")"
                ActiveCell.Offset(0, 2).Range("A1").Select
                ActiveCell.FormulaR1C1 = "=R[-2]C[3]"
                Selection.ShrinkToFit = True
                Selection.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
                Selection.Font.Size = 16
                ActiveCell.Offset(1, -2).Range("A1").Select
                ActiveCell.FormulaR1C1 = _
                    "=CONCATENATE(R5C3,"" "", ""TYPE FUNDS TO THE ENGINEERING FUND"")"
                ActiveCell.Offset(0, 2).Range("A1").Select
                ActiveCell.FormulaR1C1 = "=ROUND(PROJECT_DATA!R2C11*R[-9]C[3]/100/100,0)*100"
                Selection.ShrinkToFit = True
                ActiveCell.Select
                Selection.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
                Selection.Font.Size = 16
                Selection.ShrinkToFit = True
                ActiveCell.Offset(1, -2).Range("A1").Select
                ActiveCell.FormulaR1C1 = _
                    "=CONCATENATE(""TRANSFER FUNDS TO THE PROJECT'S ENGINEERING FUND"")"
                ActiveCell.Offset(3, 1).Range("A1").Select
                Else
        End If
        
    ' If the BidTab Estimate, then insert the following extra calculations
        If ESTIMATE_SELECTED = "BIDTAB" Then
                ActiveCell.FormulaR1C1 = "PERCENT OVER / UNDER PARTICIPATING OSARC ESTIMATE"
                ActiveCell.Offset(2, 0).Range("A1").Select
            Else
        End If
        
End Sub


'
'
' Non-Participating Pay Items are after all the participating items'
'
 Sub OrganizePayItems_NP()
    
    PAYITEMTYPE.Select
    Dim RowType As Range                         ' Declare a variable for the Pay Item Type and the range of data from the Data Validation sheet
    For Each RowType In Range("A2:A55")          ' Maximum of 50 types the user can add to the program.  4 default types
       
        If IsEmpty(RowType) = False Then
        
    ' Select pay items on Database
    Database.Select
    Database.Range("$A$1:$CS$999").AutoFilter
    Range("A2").Select

' Filter pay items to Non-Participating Items'
    ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=47, Criteria1:="No"
' Filter pay items to type'
    ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=46, Criteria1:=RowType.Value
    Range("A2").Select
    Selection.End(xlDown).Select
            If Selection <> "" Then
                Range("A2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Range(Selection, Selection.End(xlToRight)).Select
                Selection.Copy
            ' Pasting copied Participating Pay Items to the Engineer's Estimate'
                Sheets(ESTIMATE_SELECTED).Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                    If Selection.End(xlDown).Value <> "" Then
                        Selection.End(xlDown).Select
                    Else
                    End If
                ActiveCell.Offset(1, 1).Range("A1").Select
                ActiveCell.FormulaR1C1 = "Subtotal NON-Participating " & RowType.Value & " Items"
                ActiveCell.Select
                Selection.Font.Bold = True
                Selection.WrapText = True
                Selection.Font.Size = 14
                Selection.Font.Color = -4165632
                ActiveCell.Offset(0, 4).Range("A1").Select
            ' Dim RowTypeValue As String
                ActiveCell.FormulaR1C1 = "=SUMIFS(C,C[40]," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C[41],""No"")"
                ActiveCell.Select
                Selection.Font.Bold = True
                Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                Selection.Font.Size = 14
                Selection.Font.Color = -4165632
                ActiveCell.Offset(2, -5).Range("A1").Select
                End If
        End If
    Next RowType
    
    Sheets(ESTIMATE_SELECTED).Select
    ' Sum and format the project Non-participating totals
        ActiveCell.Offset(0, 1).Range("A1").Select
        ActiveCell.FormulaR1C1 = "PROJECT NON-PARTICIPATING TOTAL"
        Selection.Font.Color = -4165632
        Selection.Font.Bold = True
        Selection.Font.Size = 16
        ActiveCell.Offset(0, 4).Range("A1").Select
        ActiveCell.FormulaR1C1 = "=SUMIFS(C,C[41],""No"")"
        Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        Selection.Font.Color = -4165632
        Selection.Font.Bold = True
        Selection.ShrinkToFit = True
        Selection.Font.Size = 16
        ActiveCell.Offset(2, -4).Range("A1").Select
        
    ' If the OSARC or Contractors Estimate, then insert the following extra calculations
        If ESTIMATE_SELECTED = "OSARC" Or ESTIMATE_SELECTED = "CONTRACTOR" Then
                ActiveCell.FormulaR1C1 = "NON-PARTICIPATING CONTINGENCIES"
                ActiveCell.Offset(0, 4).Range("A1").Select
                ActiveCell.FormulaR1C1 = "=R[-2]C*0.05"
                Selection.Font.Size = 16
                Selection.ShrinkToFit = True
                ActiveCell.Offset(2, -4).Range("A1").Select
                ActiveCell.FormulaR1C1 = "PROJECT NON-PARTICIPATING TOTAL plus (+) CONTINGENCIES"
                ActiveCell.Offset(0, 4).Range("A1").Select
                ActiveCell.FormulaR1C1 = "=R[-4]C+R[-2]C"
                Selection.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
                Selection.Font.Size = 16
                Selection.ShrinkToFit = True
                ActiveCell.Offset(2, -5).Range("A1").Select
                ActiveCell.FormulaR1C1 = "=CONCATENATE(R5C3,"" "",""FUNDS REQUESTED"")"
                ActiveCell.Offset(0, 2).Range("A1").Select
                ActiveCell.FormulaR1C1 = "=R[-2]C[3]"
                Selection.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
                Selection.Font.Size = 16
                ActiveCell.Offset(1, -2).Range("A1").Select
                ActiveCell.FormulaR1C1 = _
                    "=CONCATENATE(R5C3,"" "", ""TYPE FUNDS TO THE ENGINEERING FUND"")"
                ActiveCell.Offset(0, 2).Range("A1").Select
                ActiveCell.FormulaR1C1 = "=PROJECT_DATA!R2C11*R[-7]C[3]/100"
                ActiveCell.Select
                Selection.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
                Selection.Font.Size = 16
                Selection.ShrinkToFit = True
                ActiveCell.Offset(1, -2).Range("A1").Select
                ActiveCell.FormulaR1C1 = _
                    "=CONCATENATE(""TRANSFER FUNDS TO THE PROJECT'S ENGINEERING FUND"")"
                ActiveCell.Offset(3, 1).Range("A1").Select
                Else
        End If
        
     ' If the OSARC or Contractors Estimate, then insert the following extra calculations
        If ESTIMATE_SELECTED = "BIDTAB" Then
            ActiveCell.FormulaR1C1 = "PERCENT OVER / UNDER NON-PARTICIPATING OSARC ESTIMATE"
            Selection.Font.Color = -4165632
            ActiveCell.Offset(4, 0).Range("A1").Select
            Else
        End If
            
End Sub




