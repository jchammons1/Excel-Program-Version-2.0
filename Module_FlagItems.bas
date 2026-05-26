Attribute VB_Name = "Module_FlagItems"

'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'*********************************** This section contains items that need to be highlighted or flagged if not the correct amount *******>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################


 Sub Flag_618B_803B()
  
  'Determine if S-618-B and S-803-B exist in the pay items and if yes, flag them for being the incorrect value.
  ' Only looks at the 1st instance
    Dim ws As Worksheet
    Set ws = ActiveSheet
        If Not ws.Cells.Find("S-618-B", SearchOrder:=xlByRows, SearchDirection:=xlNext) Is Nothing Then
          ' Flag these pay items for the reviewer if the value is not correct
            Cells.Find(What:="S-618-B", After:=ActiveCell, LookIn:=xlValues, _
                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                MatchCase:=False, SearchFormat:=False).Activate
            ActiveCell.Offset(0, 4).Range("A1").Select
                If ActiveCell.Value <> 10 Then
                    ActiveCell.Select
                    Selection.Font.Bold = True
                    Selection.Interior.Color = 65535
                End If
        End If
        
    Range("A2").Select
    
    Dim ws1 As Worksheet
    Set ws1 = ActiveSheet
        If Not ws1.Cells.Find("S-803-B", SearchOrder:=xlByRows, SearchDirection:=xlNext) Is Nothing Then
          ' Flag these pay items for the reviewer if the value is not correct
            Cells.Find(What:="S-803-B", After:=ActiveCell, LookIn:=xlValues, _
                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                MatchCase:=False, SearchFormat:=False).Activate
                ActiveCell.Offset(0, 4).Range("A1").Select
                    If ActiveCell.Value <> 4000 Then
                        ActiveCell.Select
                        Selection.Font.Bold = True
                        Selection.Interior.Color = 65535
                    End If
         End If
End Sub

'
'
' Flag each bidders that exceed the amount for only the winning bidder
 Sub Flag_618B_803B_Bidders()
      ' Determine if S-618-B and S-803-B exist in the pay items and if yes, flag them for being the incorrect value.
  ' Only looks at the 1st instance
  ' Need to check the winning bidder to match sure the unit cost is correct
  ' Conditional formatting is needed because the user has not yet typed the unit cost until the bid letting
  
    Dim ws5 As Worksheet
    Set ws5 = ActiveSheet
        If Not ws5.Cells.Find("S-618-B", SearchOrder:=xlByRows, SearchDirection:=xlPrevious) Is Nothing Then
          ' Flag these pay items for the reviewer if the value is not correct
                Cells.Find(What:="S-618-B", After:=ActiveCell, LookIn:=xlValues, _
                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                MatchCase:=False, SearchFormat:=False).Activate
            ActiveCell.Offset(0, 105).Range("A1").Select
            Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
                Formula1:="=10"
            Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
            With Selection.FormatConditions(1).Font
                .Bold = True
                .Italic = False
                .TintAndShade = 0
            End With
            With Selection.FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
            End With
            Selection.FormatConditions(1).StopIfTrue = False
        End If
        
    Dim ws6 As Worksheet
    Set ws6 = ActiveSheet
        If Not ws6.Cells.Find("S-803-B", SearchOrder:=xlByRows, SearchDirection:=xlPrevious) Is Nothing Then
          ' Flag these pay items for the reviewer if the value is not correct
                Cells.Find(What:="S-803-B", After:=ActiveCell, LookIn:=xlValues, _
                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                MatchCase:=False, SearchFormat:=False).Activate
            ActiveCell.Offset(0, 105).Range("A1").Select
            Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
                Formula1:="=4000"
            Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
            With Selection.FormatConditions(1).Font
                .Bold = True
                .Italic = False
                .TintAndShade = 0
            End With
            With Selection.FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
            End With
            Selection.FormatConditions(1).StopIfTrue = False
        End If
End Sub


'
'
'

 Sub Flag_618B_803B_Form902()
    'Find cell with "S-618-B" and enter Additional Construction signs Unit Price = $10.00 and $0.00 total
    '############# ADDITIONAL CONSTRUTION SIGNS ###################
    Range("A12").Select
    
Dim ws902 As Worksheet
Set ws902 = Form902

        If Not ws902.Cells.Find("S-618-B", SearchOrder:=xlByRows, SearchDirection:=xlPrevious) Is Nothing Then
          ' Flag these pay items for the reviewer if the value is not correct
            Cells.Find(What:="S-618-B", After:=ActiveCell, LookIn:=xlValues _
                , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
            ActiveCell.Select
            ActiveCell.Offset(0, 4).Range("A1").Select
            ActiveCell.FormulaR1C1 = "10"
                    Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorDark2
                        .TintAndShade = -9.99786370433668E-02
                        .PatternTintAndShade = 0
                    End With
            ActiveCell.Offset(0, 1).Range("A1").Select
                ActiveCell.FormulaR1C1 = "0"
                    Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorDark2
                        .TintAndShade = -9.99786370433668E-02
                        .PatternTintAndShade = 0
                    End With
                Range("A12").Select
                Else
                Range("A12").Select
            End If

            
'Dim ws902 As Worksheet
'Set ws902 = Form902
    If Not ws902.Cells.Find("S-803-B", SearchOrder:=xlByRows, SearchDirection:=xlNext) Is Nothing Then
        Cells.Find(What:="S-803-B", After:=ActiveCell, LookIn:=xlValues _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    'Find cell with "S-803-B" and enter Test Pile = $4,000 and $0.00 total
    '############# ADDITIONAL CONSTRUTION SIGNS ###################
   Range("A12").Select
    Cells.Find(What:="S-803-B", After:=ActiveCell, LookIn:=xlValues _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Select
    ActiveCell.Offset(0, 4).Range("A1").Select
    ActiveCell.FormulaR1C1 = "4000"
                Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark2
                .TintAndShade = -9.99786370433668E-02
                .PatternTintAndShade = 0
            End With
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "0"
                Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark2
                .TintAndShade = -9.99786370433668E-02
                .PatternTintAndShade = 0
            End With
        Range("A12").Select
        Else
        Range("A12").Select
    End If
    
End Sub
