Attribute VB_Name = "Module_FuelCalculations"
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'*********************************** This section contains calculations for the fuel adjustment on the CAD ******************************>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################


'
'
'
Sub Fuel_Summary()

' copy fuel summary
' First find which document is present and start adding the fuel summary
' Otherwise the fuel summary will overwrite the change order data
        Range("A52").Select
        
        Dim ws As Worksheet
        Set ws = ActiveSheet
        
        If Not ws.Cells.Find("End of Change Order Summary", SearchOrder:=xlByRows, SearchDirection:=xlNext) Is Nothing Then
            Cells.Find(What:="End of Change Order Summary", After:=ActiveCell, LookIn _
            :=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
            xlNext, MatchCase:=False, SearchFormat:=False).Activate
            ActiveCell.Offset(2, 0).Range("A1").Select
            ActiveCell.Range("A1:L999").Clear
                Else
            Cells.Find(What:="End of CAD document", After:=ActiveCell, LookIn:= _
            xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
            xlNext, MatchCase:=False, SearchFormat:=False).Activate
            ActiveCell.Offset(2, 0).Range("A1").Select
            ActiveCell.Range("A1:L999").Clear
        End If

    Range("GA1:GK17").Copy
    ActiveCell.Select
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
' Select pay items on Database
    Database.Select
    Database.Range("$A$1:$CS$999").AutoFilter
    Range("A2").Select
' Filter pay items
    ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=8, Criteria1:=Array("A1" _
        , "A2", "A3", "A4", "A5", "A6", "B", "BA1", "C", "D", "E", "GT", "GY", "M", "S"), Operator _
        :=xlFilterValues
    ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=47, Criteria1:="=No", _
        Operator:=xlOr, Criteria2:="=Yes"
    Range("A2").Select
    
    Selection.End(xlDown).Select
        If Selection <> "" Then

                
        ' Select pay items on Database
            Database.Select
            Database.Range("$A$1:$CS$999").AutoFilter
            Range("A2").Select
        ' Filter pay items
    ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=8, Criteria1:=Array("A1" _
        , "A2", "A3", "A4", "A5", "A6", "B", "BA1", "C", "D", "E", "GT", "GY", "M", "S"), Operator _
        :=xlFilterValues
    ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=47, Criteria1:="=No", _
        Operator:=xlOr, Criteria2:="=Yes"
            Range("A2").Select
            Selection.End(xlDown).Select
                    If Selection <> "" Then
                        Range("A2:B2").Select
                        Range(Selection, Selection.End(xlDown)).Select
                        Selection.Copy
                    ' Pasting copied Pay Item, Item Description, and Contract quantity to the CAD
                       CAD.Select
                             Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False
                        ActiveCell.Offset(0, 2).Range("A1").Select
                    ' Pasting copied Unit and Unit Price to theCAD
                        Database.Select
                        Range("H2").Select
                        Range(Selection, Selection.End(xlDown)).Select
                        Selection.Copy
                       CAD.Select
                             Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False
                        ActiveCell.Offset(0, 1).Range("A1").Select
                    ' Pasting copied Subtotal Contract Amount to theCAD
                        Database.Select
                        Range("R2:S2").Select
                        Range(Selection, Selection.End(xlDown)).Select
                        Selection.Copy
                       CAD.Select
                             Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False
                        ActiveCell.Offset(0, 2).Range("A1").Select
                     ' Pasting copied Type
                        Database.Select
                        Range("AT2").Select
                        Range(Selection, Selection.End(xlDown)).Select
                        Selection.Copy
                        CAD.Select
                             Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False
                        ActiveCell.Offset(0, 4).Range("A1").Select
                      ' Pasting copied Participating
                        Database.Select
                        Range("AU2").Select
                        Range(Selection, Selection.End(xlDown)).Select
                        Selection.Copy
                        CAD.Select
                             Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False
                            
                            
                            
                            If IsEmpty(Selection.End(xlDown).Value) = True Then
                                ActiveCell.Offset(2, -9).Range("A1").Select
                                Else
                                Selection.End(xlDown).Select
                                ActiveCell.Offset(2, -9).Range("A1").Select
                            End If
                    End If
                        
            CAD.Select
            ActiveCell.Value = "End of Fuel Adjustment Summary"
            
        Else
' If statement needed because if no pay items have a fuel code, then there is nothing to select.  Hence, the ActiveCell.Value = "End of Fuel Adjustment Summary"
' text will not be at the bottom, then a debug error occurs when the print area is trying to be set based on that text in Module_PrintArea
        CAD.Select
        ActiveCell.Offset(1, 0).Range("A1").Select
        ActiveCell.Value = "No Pay Items have a fuel adjustment code"
        ActiveCell.Offset(1, 0).Range("A1").Select
        ActiveCell.Value = "End of Fuel Adjustment Summary"

    End If
    Database.Range("$A$1:$CS$999").AutoFilter
    CAD.Select


End Sub

