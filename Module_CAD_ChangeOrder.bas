Attribute VB_Name = "Module_CAD_ChangeOrder"
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'**************************** This section contains the Change Order Tracking data that goes at the bottom of the CAD-002 section *******>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################


 Sub CAD_ChangeOrder()
' Select pay items on Database
    Database.Select
    Database.Range("$A$1:$CS$999").AutoFilter
    Range("A2").Select
' Filter pay items to check if there is any change order
    ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=51, Criteria1:="=*No.*"
    Range("A2").Select
    Selection.End(xlDown).Select
        If Selection <> "" Then
            DATA_VALIDATION.Visible = True
            DATA_VALIDATION.Select
            Dim ChangeOrder As Range                         ' Declare a variable for the Pay Item Type and the range of data from the Data Validation sheet
            For Each ChangeOrder In Range("J3:J32")          ' Maximum of 50 types the user can add to the program.  4 default types
               
                
        ' Select pay items on Database
            Database.Select
            Database.Range("$A$1:$CS$999").AutoFilter
            Range("A2").Select
        ' Filter pay items to Change Order No'
            ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=51, Criteria1:=ChangeOrder.Value
            Range("A2").Select
            Selection.End(xlDown).Select
                    If Selection <> "" Then
                        Range("A2:C2").Select
                        Range(Selection, Selection.End(xlDown)).Select
                        Selection.Copy
                    ' Pasting copied Pay Item, Item Description, and Contract quantity to the CAD
                       CAD.Select
                             Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False
                        ActiveCell.Offset(0, 5).Range("A1").Select
                    ' Pasting copied Unit and Unit Price to theCAD
                        Database.Select
                        Range("D2:E2").Select
                        Range(Selection, Selection.End(xlDown)).Select
                        Selection.Copy
                       CAD.Select
                             Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False
                        ActiveCell.Offset(0, 3).Range("A1").Select
                    ' Pasting copied Subtotal Contract Amount to theCAD
                        Database.Select
                        Range("F2").Select
                        Range(Selection, Selection.End(xlDown)).Select
                        Selection.Copy
                       CAD.Select
                             Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False
                        ActiveCell.Offset(0, 1).Range("A1").Select
                     ' Pasting copied Type
                        Database.Select
                        Range("AT2:AU2").Select
                        Range(Selection, Selection.End(xlDown)).Select
                        Selection.Copy
                       CAD.Select
                             Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False
                        ActiveCell.Offset(0, 2).Range("A1").Select
                     ' Pasting copied Type
                        Database.Select
                        Range("AY2").Select
                        Range(Selection, Selection.End(xlDown)).Select
                        Selection.Copy
                       CAD.Select
                             Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=False
                            
                            If IsEmpty(Selection.End(xlDown).Value) = True Then
                                ActiveCell.Offset(1, -3).Range("A1").Select
                                Else
                                Selection.End(xlDown).Select
                                ActiveCell.Offset(1, -3).Range("A1").Select
                            End If


                        ' Dim Change Order Value As String
                                ActiveCell.FormulaR1C1 = "=SUMIFS(C,C[3]," & _
                                 Chr(34) & _
                                 ChangeOrder.Value & _
                                 Chr(34) & ")"
                                 
                            Call Format_SubtotalAmount_Participating
                        ActiveCell.Offset(0, -7).Range("A1").Select
                        ActiveCell.FormulaR1C1 = "Total of Change Order " & ChangeOrder.Value & " Items"
                            Call Format_SubtotalText_Participating
                        ActiveCell.Offset(2, -1).Range("A1").Select
                        
                        End If
            Next ChangeOrder
            CAD.Select
            ActiveCell.Value = "End of Change Order Summary"
            
        Else

    End If
    Database.Range("$A$1:$CS$999").AutoFilter
    CAD.Select
End Sub

