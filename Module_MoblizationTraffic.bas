Attribute VB_Name = "Module_MoblizationTraffic"
 Sub Mobilization_Traffic()
' The mobilization and traffic are calculated in the loop.  Need to delete them again so the mobilziation and traffic calculates on the participating items without those amounts.
' Clear mobilization and traffic maintenance to avoid calculation error AGAIN.  I am clearing it at the beginning as well. 2nd clearing.

Call Mobilization_MaintenanceTraffic_Clear
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
Range("EZ2").ClearContents  ' Clear S-618-A place holder cost
Range("A52").Select
 'Find cell with "S-618-A"                                           ############# TRAFFIC MAINTENANCE ###################
    If Not ws.Cells.Find("S-618-A", SearchOrder:=xlByRows, SearchDirection:=xlNext) Is Nothing Then
            Cells.Find(What:="S-618-A", After:=ActiveCell, LookIn:=xlValues, _
                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False).Activate
            ActiveCell.Select
            ActiveCell.Offset(0, 4).Range("A1").Select
            ActiveCell.Value = Round(((Range("FD3").Value + Range("HD2").Value) / Range("FE3").Value), 3)
            ' FD3 = PROJECT_PARTICIPATING_AND_NON-PARTICPATING ITEMS (Current)
            ' HD2 = Accumulative Total Fuel Adjustment
            ' FE3 = PROJECT_PARTICIPATING_AND_NON-PARTICPATING ITEMS (Contract Total)
                Selection.NumberFormat = "0.000"

        ' Select the allowed to date and calculate the current plus the past
            ActiveCell.Offset(0, -1).Select
            ActiveCell.Value = Round((ActiveCell.Offset(0, 1).Range("A1").Value - ActiveCell.Offset(0, 59).Range("A1").Value), 3)
                Selection.NumberFormat = "0.000"
                Selection.HorizontalAlignment = xlCenter
                Selection.VerticalAlignment = xlCenter
                Selection.Locked = True
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = -0.149998474074526
                    .PatternTintAndShade = 0
                End With
                
            '
            '**************************************************************************************************************************************************
            '**************************************************************************************************************************************************
            '**************************************************************************************************************************************************
            ' In rare instances, the maintenance of traffic calculation with volatile contract price swings from change orders and odd corrects will cause this.
            ' See Project SAP-78(11) Invoice #2 to #3.  The values on the file copy were manually edited to 0.900 because it was calculation 1.087.
            '
    
                ActiveCell.Offset(0, 1).Select
                If ActiveCell.Value >= 1 Then
                    ActiveCell.Value = 1
                    ActiveCell.Offset(0, -1).Value = Round((1 - ActiveCell.Offset(0, 59).Range("A1").Value), 3)
                ElseIf ActiveCell.Offset(0, -1).Value < 0 Then
                    ActiveCell.Offset(0, -1).Value = 0
                    ActiveCell.Value = ActiveCell.Offset(0, 58).Range("A1").Value
                    ActiveCell.Offset(0, 3).Value = ActiveCell.Offset(0, 59).Range("A1").Value
                    
                End If
            ''**************************************************************************************************************************************************
            '**************************************************************************************************************************************************
            ''**************************************************************************************************************************************************
            '
            ActiveCell.Offset(0, 3).Select
            ActiveCell.Value = ActiveCell.Offset(0, -3).Value * ActiveCell.Offset(0, -1).Value
                Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            ActiveCell.Offset(0, 56).Range("A1").Value = ActiveCell.Value                                       ' Past Period Cost
            ActiveCell.Offset(0, 57).Range("A1").Value = ActiveCell.Offset(0, -4).Range("A1").Value           ' Current Period Quantity BM / 57
            ActiveCell.Offset(0, 58).Range("A1").Value = ActiveCell.Offset(0, -4).Range("A1").Value * ActiveCell.Offset(0, -1).Range("A1").Value               ' Current Period Cost BN
            ActiveCell.Offset(0, 59).Range("A1").Value = ActiveCell.Offset(0, -3).Range("A1").Value            ' Total Quanity
            ActiveCell.Offset(0, 60).Range("A1").Value = ActiveCell.Value                                      ' Total Cost
            Range("EZ2").Value = ActiveCell.Value
    End If
         
         
         

    Range("A52").Select
' This has to be here because the search starts from the top down.  Moved this after the S-618-A calculation so this subtotal does not get calculated
' with the S-618-A
'Find cell with "S-200-A" and copy the amounts to GC1:GD1 for use in Step 4                  ############# MOBILIZATION ###################
    If Not ws.Cells.Find("S-200-A", SearchOrder:=xlByRows, SearchDirection:=xlNext) Is Nothing And Range("FA2").Value >= 0.1 Then
        Cells.Find(What:="S-200-A", After:=ActiveCell, LookIn:=xlValues _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Select
        ActiveCell.Offset(0, 3).Range("A1").Select
        ActiveCell = 1
        ActiveCell.Select
            Selection.Locked = True
            Selection.NumberFormat = "0.000"
            Selection.FormulaHidden = False
        Selection.HorizontalAlignment = xlCenter
        Selection.VerticalAlignment = xlCenter
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.149998474074526
                .PatternTintAndShade = 0
            End With
        ActiveCell.Offset(0, 1).Range("A1").Value = ActiveCell.Value
        ActiveCell.Value = ActiveCell.Offset(0, 1).Range("A1").Value - ActiveCell.Offset(0, 59).Range("A1").Value
    ElseIf Not ws.Cells.Find("S-200-A", SearchOrder:=xlByRows, SearchDirection:=xlNext) Is Nothing And Range("FA2").Value >= 0.05 Then
        Range("A52").Select
        Cells.Find(What:="S-200-A", After:=ActiveCell, LookIn:=xlValues _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Select
        ActiveCell.Offset(0, 3).Range("A1").Select
        ActiveCell = 0.5
    ActiveCell.Select
        Selection.Locked = True
        Selection.NumberFormat = "0.000"
        Selection.FormulaHidden = False
        Selection.HorizontalAlignment = xlCenter
        Selection.VerticalAlignment = xlCenter
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.149998474074526
            .PatternTintAndShade = 0
        End With
        ActiveCell.Offset(0, 1).Range("A1").Value = ActiveCell.Value
        ActiveCell.Value = ActiveCell.Offset(0, 1).Range("A1").Value - ActiveCell.Offset(0, 59).Range("A1").Value
    ElseIf Not ws.Cells.Find("S-200-A", SearchOrder:=xlByRows, SearchDirection:=xlNext) Is Nothing Then
    Range("A52").Select
    Cells.Find(What:="S-200-A", After:=ActiveCell, LookIn:=xlValues _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Select
        ActiveCell.Offset(0, 3).Range("A1").Select
        ActiveCell = 0
        Selection.Locked = True
        Selection.NumberFormat = "0.000"
        Selection.FormulaHidden = False
        Selection.HorizontalAlignment = xlCenter
        Selection.VerticalAlignment = xlCenter
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.149998474074526
            .PatternTintAndShade = 0
        End With
        ActiveCell.Offset(0, 1).Range("A1").Value = ActiveCell.Value
        ActiveCell.Value = ActiveCell.Offset(0, 1).Range("A1").Value - ActiveCell.Offset(0, 60).Range("A1").Value
        Else
        Range("A52").Select
    End If
    
    
    If Not ws.Cells.Find("S-200-A", SearchOrder:=xlByRows, SearchDirection:=xlNext) Is Nothing Then
    Range("A52").Select        ' This has to be here because the search starts from the top down
        'Find cell with "S-200-A" and clear contentes                  ############# MOBILIZATION ###################
            Cells.Find(What:="S-200-A", After:=ActiveCell, LookIn:=xlValues _
            , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
            ActiveCell.Select
            ActiveCell.Offset(0, 7).Range("A1").Select
            ActiveCell.Select
            ActiveCell.Value = ActiveCell.Offset(0, -3).Value * ActiveCell.Offset(0, -1).Value
            ActiveCell.Offset(0, 57).Range("A1").Value = ActiveCell.Offset(0, -4).Range("A1").Value           ' Current Period Quantity BM / 57
            ActiveCell.Offset(0, 58).Range("A1").Value = ActiveCell.Offset(0, -4).Range("A1").Value * ActiveCell.Offset(0, -1).Range("A1").Value               ' Current Period Cost BN
            ActiveCell.Offset(0, 59).Range("A1").Value = ActiveCell.Offset(0, -3).Range("A1").Value            ' Total Quanity
            ActiveCell.Offset(0, 60).Range("A1").Value = ActiveCell.Value                                      ' Total Cost
     End If
Range("A52").Select
    
End Sub

Sub Mobilization_MaintenanceTraffic_Clear()

    Range("A52").Select

    Dim ws As Worksheet
    Set ws = ActiveSheet
' Clear mobilization and traffic maintenance to avoid calculation error
' Find cell with "S-200-A" and clear contents                  ############# MOBILIZATION ###################
        If Not ws.Cells.Find("S-200-A", SearchOrder:=xlByRows, SearchDirection:=xlNext) Is Nothing Then
            Cells.Find(What:="S-200-A", After:=ActiveCell, LookIn:=xlValues, _
                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False).Activate
            ActiveCell.Select
            ActiveCell.Offset(0, 3).Range("A1").Select
            ActiveCell.ClearContents
            ActiveCell.Offset(0, 4).Range("A1").Select
            ActiveCell.ClearContents
            ActiveCell.Offset(0, 57).Range("A1").Select
            ActiveCell.ClearContents
            ActiveCell.Offset(0, 1).Range("A1").Select
            ActiveCell.ClearContents
            ActiveCell.Offset(0, 1).Range("A1").Select
            ActiveCell.ClearContents
            ActiveCell.Offset(0, 1).Range("A1").Select
            ActiveCell.ClearContents
         End If
         
Range("A52").Select
 'Find cell with "S-618-A"                                           ############# TRAFFIC MAINTENANCE ###################
        If Not ws.Cells.Find("S-618-A", SearchOrder:=xlByRows, SearchDirection:=xlNext) Is Nothing Then
            Cells.Find(What:="S-618-A", After:=ActiveCell, LookIn:=xlValues, _
                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False).Activate
            ActiveCell.Select
            ActiveCell.Offset(0, 3).Range("A1").Select
            ActiveCell.ClearContents
            ActiveCell.Offset(0, 1).Range("A1").Select
            ActiveCell.ClearContents
            ActiveCell.Offset(0, 3).Range("A1").Select
            ActiveCell.ClearContents
            ActiveCell.Offset(0, 57).Range("A1").Select
            ActiveCell.ClearContents
            ActiveCell.Offset(0, 1).Range("A1").Select
            ActiveCell.ClearContents
            ActiveCell.Offset(0, 1).Range("A1").Select
            ActiveCell.ClearContents
            ActiveCell.Offset(0, 1).Range("A1").Select
            ActiveCell.ClearContents
         End If

 Range("A52").Select

End Sub
