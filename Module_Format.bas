Attribute VB_Name = "Module_Format"
'' Test text 5-26-26
    Public Function PayItemSelectedRegion() As Range
        Database.Unprotect
        Set PayItemSelectedRegion = Database.Range("A2").CurrentRegion
    End Function
    Public Function CountPayItems() As Integer
        CountPayItems = PayItemSelectedRegion.Rows.Count + 25 + 21        ' Counts the number of pay items and adds 25 additional rows to format because of extra lines inserted that are not formatted
    End Function

'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'*********************************** This section contains formats for various documents*************************************************>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'
'
'
    Sub Format_Estimates()

            row = 21
            column = 1
            endrow = CountPayItems
            
    Do While row < endrow
        If Cells(row, column + 46).Value = "Yes" Or Cells(row, column + 46).Value = "No" Then

        Range(Cells(row, column), Cells(row, column + 5)).Select              'Format entire row from A to L columns
        Selection.Locked = True
        Selection.HorizontalAlignment = xlCenter
        Selection.VerticalAlignment = xlCenter
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
        Selection.Font.Size = 16
        Selection.ShrinkToFit = True
         
        Cells(row, column + 1).Select                   'Column B - Pay Item Description
        Selection.HorizontalAlignment = xlLeft
        Selection.VerticalAlignment = xlCenter
        Selection.WrapText = True
        Selection.Font.Size = 14
        
        Cells(row, column + 2).Select                   'Column C - Quantities / Contract Amount
        Selection.ShrinkToFit = True
        Selection.Font.Size = 16
        Selection.NumberFormat = "0.000"

        Cells(row, column + 3).Select                   'Column D - Quantities / Current Period
        Selection.ShrinkToFit = True
        Selection.NumberFormat = "0.000"

        Range(Cells(row, column + 4), Cells(row, column + 5)).Select                  'Column E - Quantities / Allowed To Date
        Selection.ShrinkToFit = True
        Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        End If
            row = row + 1
    Loop
        
    Columns("G:CY").Select
    Selection.EntireColumn.Hidden = True
    Rows("21:999").EntireRow.AutoFit

End Sub
    
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'*********************************** This section contains Bid Tab formatting for special calculations **********************************>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'
'
'
  Sub Format_BIDTAB()

            row = 11
            column = 1
            endrow = CountPayItems

    Do While row < endrow
        If Cells(row, column + 2).Value = "PROJECT PARTICIPATING TOTAL" Or Cells(row, column + 2).Value = "PERCENT OVER / UNDER PARTICIPATING OSARC ESTIMATE" _
        Or Cells(row, column + 2).Value = "PERCENT OVER / UNDER NON-PARTICIPATING OSARC ESTIMATE" _
        Or Cells(row, column + 2).Value = "PROJECT PARTICIPATING AND NON-PARTICIPATING TOTAL" _
        Or Cells(row, column + 2).Value = "PERCENT OVER / UNDER PARTICIPATING AND NON-PARTICIPATING OSARC ESTIMATE" Then
            
            Range(Cells(row, column), Cells(row, column + 5)).Select              'Format entire row from A to L columns
            Selection.Locked = True
            Selection.HorizontalAlignment = xlCenter
            Selection.VerticalAlignment = xlCenter
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
            Selection.Font.Size = 16
            Selection.ShrinkToFit = True
             
            Cells(row, column + 1).Select                   'Column B - Pay Item Description
            Selection.HorizontalAlignment = xlLeft
            Selection.VerticalAlignment = xlCenter
            Selection.WrapText = True
            Selection.Font.Size = 14
            
            Cells(row, column + 2).Select                   'Column C - Quantities / Contract Amount
            Selection.ShrinkToFit = True
            Selection.Font.Size = 16
            Selection.NumberFormat = "0.000"
    
            Cells(row, column + 3).Select                   'Column D - Quantities / Current Period
            Selection.ShrinkToFit = True
            Selection.NumberFormat = "0.000"
    
            Range(Cells(row, column + 4), Cells(row, column + 5)).Select                  'Column E - Quantities / Allowed To Date
            Selection.ShrinkToFit = True
            Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        End If

    Loop
        
    Columns("G:CY").Select
    Selection.EntireColumn.Hidden = True
    Rows("21:999").EntireRow.AutoFit

End Sub


'
'
'
'
 Sub Format_CAD()
            row = 50
            column = 1
            endrow = CountPayItems + 53
  '  Dim ParticipatingStatus As String
    
    Do While row < endrow

       ' Select Case ParticipatingStatus = Cells(row, column + 49).Value
        '    Case "Yes", "No", "Correction Participating", "Correction Non-Participating"
        If Cells(row, column + 49).Value = "Yes" Or Cells(row, column + 49).Value = "No" Or _
        Cells(row, column + 49).Value = "Correction Participating" Or Cells(row, column + 49).Value = "Correction Non-Participating" Then

            Range(Cells(row, column), Cells(row, column + 11)).Select              'Format entire row from A to L columns
            Selection.HorizontalAlignment = xlCenter
            Selection.VerticalAlignment = xlCenter
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
            Selection.Font.Size = 16
            Selection.ShrinkToFit = True
             
            Cells(row, column + 1).Select                   'Column B - Pay Item Description
            Selection.HorizontalAlignment = xlLeft
            Selection.VerticalAlignment = xlCenter
            Selection.WrapText = True
            Selection.Font.Size = 14
            
            Cells(row, column + 2).Select                   'Column D - Quantities / Current Period
            Selection.ShrinkToFit = True
            Selection.NumberFormat = "0.000"
            
            Cells(row, column + 3).Select                   'Column C - Quantities / Contract Amount
            Selection.Locked = False
                With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 13434879
                .TintAndShade = 0
                .PatternTintAndShade = 0
                End With
            Selection.ShrinkToFit = True
            Selection.Font.Size = 16
            Selection.NumberFormat = "0.000"
    
            Range(Cells(row, column + 4), Cells(row, column + 5)).Select                  'Column E - Quantities / Allowed To Date
            Selection.ShrinkToFit = True
            Selection.NumberFormat = "0.000"
            
            Range(Cells(row, column + 6), Cells(row, column + 9)).Select                  'Column E - Quantities / Allowed To Date
            Selection.ShrinkToFit = True
            Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            
            Cells(row, column + 11).Select                  'Column L Fuel Month Year
            Selection.Font.Size = 12
            Selection.WrapText = True
            
            Range(Cells(row, column + 42), Cells(row, column + 46)).Select               'Column AQ to AU - Fuel totals
            Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            Selection.Font.Bold = True
            Selection.HorizontalAlignment = xlCenter
            Selection.VerticalAlignment = xlCenter
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
            
            Cells(row, column + 62).Select                'Past Period Quantity
            Selection.NumberFormat = "0.000"
            
            Cells(row, column + 64).Select                'Current Period Quantity
            Selection.NumberFormat = "0.000"
            
            Cells(row, column + 66).Select                'Current Period Quantity
            Selection.NumberFormat = "0.000"
            
            Cells(row, column + 63).Select                'Past Period Cost
            Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            
            Cells(row, column + 65).Select                'Current Period Cost
            Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            
            Cells(row, column + 67).Select                'Current Period Cost
            Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            
            Range(Cells(row, column + 62), Cells(row, column + 67)).Select
            Selection.Font.Bold = True
            Selection.HorizontalAlignment = xlCenter
            Selection.VerticalAlignment = xlCenter
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
         'End Select
         
        End If
                      
        row = row + 1
    Loop

End Sub

 Sub Format_ChangeOrder()
            row = 50
            column = 1
            endrow = CountPayItems + 53 + 25
            
    Do While row < endrow
        If Cells(row, column + 10).Value = "Change Order" Then

        Range(Cells(row, column), Cells(row, column + 11)).Select              'Format entire row from A to L columns
        Selection.HorizontalAlignment = xlCenter
        Selection.VerticalAlignment = xlCenter
        Selection.Borders(xlEdgeLeft).LineStyle = xlDash
        Selection.Borders(xlEdgeLeft).Weight = xlThin
        Selection.Borders(xlEdgeTop).LineStyle = xlDash
        Selection.Borders(xlEdgeTop).Weight = xlThin
        Selection.Borders(xlEdgeBottom).LineStyle = xlDash
        Selection.Borders(xlEdgeBottom).Weight = xlThin
        Selection.Borders(xlEdgeRight).LineStyle = xlDash
        Selection.Borders(xlEdgeRight).Weight = xlThin
        Selection.Borders(xlInsideVertical).LineStyle = xlDash
        Selection.Borders(xlInsideVertical).Weight = xlThin
        Selection.Borders(xlInsideHorizontal).LineStyle = xlDash
        Selection.Borders(xlInsideHorizontal).Weight = xlThin
        Selection.Font.Size = 16
        Selection.ShrinkToFit = True
         
        Cells(row, column + 1).Select                   'Column B - Pay Item Description
        Selection.HorizontalAlignment = xlLeft
        Selection.VerticalAlignment = xlCenter
        Selection.WrapText = True
        Selection.Font.Size = 14
        
        Cells(row, column + 2).Select                   'Column C - Quantities / Contract Amount
            Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual _
                , Formula1:="=0"
            Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
            With Selection.FormatConditions(1).Font
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.499984740745262
            End With
            With Selection.FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent3
                .TintAndShade = 0.799981688894314
            End With
            Selection.FormatConditions(1).StopIfTrue = False
        Selection.ShrinkToFit = True
        Selection.Font.Size = 16
        Selection.NumberFormat = "0.000"
        
        Range(Cells(row, column + 6), Cells(row, column + 8)).Select                  'Column E - Quantities / Allowed To Date
        Selection.ShrinkToFit = True
        Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        
        Range(Cells(row, column + 9), Cells(row, column + 10)).Select                   'Participating Status
        Selection.WrapText = True
        
        Cells(row, column + 10).Select                   'Participating Status
        Selection.Font.Size = 12
        
        End If
                      
        row = row + 1
    Loop

End Sub


 Sub Format_BidTabColumnA()
        Selection.HorizontalAlignment = xlCenter
        Selection.Locked = False
        Selection.Interior.Color = 13434879
        Selection.Font.Size = 16
        Selection.ShrinkToFit = True
        Selection.NumberFormat = "#,##0.00"
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
End Sub

 Sub Format_BidTabColumnB()
        Selection.HorizontalAlignment = xlCenter
        Selection.Font.Size = 16
        Selection.ShrinkToFit = True
        Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
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
End Sub

 Sub Format_BidTabColumnPercent()
        Selection.HorizontalAlignment = xlCenter
        Selection.Font.Size = 16
        Selection.ShrinkToFit = True
        Selection.NumberFormat = "0.00%"
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
            With Selection.FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 13551615
                .TintAndShade = 0
            End With
        Selection.FormatConditions(1).StopIfTrue = False
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
End Sub


 Sub Format_900()
          row = 19
            column = 1
            endrow = CountPayItems
        Do While row < endrow
        If Cells(row, column).Value <> "" Then
            Cells(row, column).Select
            ActiveCell.Offset(0, 0).Range("A1:G1").Select
            With Selection
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = True
                .ReadingOrder = xlContext
                .MergeCells = True
            End With
        End If
        
        If Cells(row, column).Value <> "" Then
            Cells(row, column + 7).Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = True
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.NumberFormat = "#,##0.000"
        End If
        
        If Cells(row, column).Value <> "" Then
            Cells(row, column + 8).Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = True
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
        End If
        
        If InStr(Cells(row, column + 1).Value, "ITEMS:") > 0 Then
            Range(Cells(row, column + 1), Cells(row, column + 8)).Select
            Selection.Font.Bold = True
            Selection.Font.Underline = xlUnderlineStyleSingle
            Selection.Font.Size = 12
            Selection.HorizontalAlignment = xlCenter
        End If
    row = row + 1
    Loop
End Sub

 Sub Format_WebBidList()

' Formats the columns of pay items
          row = 11
            column = 1
            endrow = CountPayItems
    Do While row < endrow
        If Cells(row, column).Value <> "" Then
            Cells(row, column).Select
            ActiveCell.Offset(0, 0).Range("A1:G1").Select
            With Selection
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = True
                .ReadingOrder = xlContext
                .MergeCells = True
            End With
        End If
        
        If Cells(row, column).Value <> "" Then
            Cells(row, column + 7).Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = True
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.NumberFormat = "#,##0.000"
        End If
        
        If Cells(row, column).Value <> "" Then
            Cells(row, column + 8).Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = True
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
        End If
            
        If InStr(Cells(row, column + 1).Value, "ITEMS:") > 0 Then
            Cells(row, column + 1).Select
            Selection.Font.Bold = True
            Selection.Font.Underline = xlUnderlineStyleSingle
            Selection.Font.Size = 12
        End If
        
    row = row + 1
    Loop
End Sub


 Sub Format_902()
    Form902.Select
            row = 13
            column = 1
            endrow = CountPayItems
    Do While row < endrow
        If Cells(row, column + 7).Value = "Yes" Or Cells(row, column + 7).Value = "No" Then
            Range(Cells(row, column), Cells(row, column + 6)).Select               'Column A to G formatting gridlines
            Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
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
            Selection.HorizontalAlignment = xlCenter
            Selection.VerticalAlignment = xlCenter
            Selection.ShrinkToFit = True
            Selection.Font.Size = 12
        End If
        
        If Cells(row, column + 7).Value = "Yes" Or Cells(row, column + 7).Value = "No" Then
            Cells(row, column + 2).Select
            Selection.HorizontalAlignment = xlLeft
            Selection.VerticalAlignment = xlCenter
            Selection.WrapText = True
        End If

        If Cells(row, column + 7).Value = "Yes" Or Cells(row, column + 7).Value = "No" Then
            Cells(row, column + 3).Select
            Selection.NumberFormat = "#,##0.000"
        End If
        
        If Cells(row, column + 7).Value = "Yes" Or Cells(row, column + 7).Value = "No" Then
            Range(Cells(row, column + 5), Cells(row, column + 6)).Select
            Selection.ClearContents
        End If
        
        If Cells(row, column + 7).Value = "Yes" Or Cells(row, column + 7).Value = "No" Then
            Cells(row, column).Select
            Selection.HorizontalAlignment = xlCenter
            Selection.VerticalAlignment = xlCenter
                   
        Dim i As Integer                        ' Adding Ref.No. to Column A
            i = i + 1
            Cells(row, column) = i
            Cells(row, column).Select
            Selection.ShrinkToFit = True
            Selection.Font.Size = 14
            Selection.Font.Bold = True
        End If
        
        If InStr(Cells(row, column).Value, "Participating Items:") > 0 Then
            Cells(row, column).Select
                Selection.Font.Bold = True
                'Selection.Font.Underline = xlUnderlineStyleSingle
                Selection.Font.Size = 14
        End If
        
        If InStr(Cells(row, column).Value, "Non-Participating Items:") > 0 Then
            Cells(row, column).Select
                Selection.Font.Color = -4165632
        End If

            row = row + 1
    Loop
End Sub

 Sub Format_GreenCover()

        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = True
                .ReadingOrder = xlContext
                .MergeCells = True
        End With
End Sub

 Sub Format_GreenCover2()
    Selection.Font.Bold = True
    Selection.ShrinkToFit = True
    Selection.Font.Size = 16
End Sub


 Sub Format_LettingResults()
        Selection.HorizontalAlignment = xlCenter
        Selection.VerticalAlignment = xlCenter
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
        Selection.Font.Size = 20
        Selection.ShrinkToFit = True
End Sub


 Sub Format_ProjectData()
    PROJECT_DATA.Range("G2:G5").Select
    Selection.NumberFormat = "0.0"
    PROJECT_DATA.Range("I2:I5").Select
    Selection.NumberFormat = "0.0"
    PROJECT_DATA.Range("J2:J5").Select
    Selection.NumberFormat = "0"
    PROJECT_DATA.Range("K2:K5").Select
    Selection.NumberFormat = "0.0"
    PROJECT_DATA.Range("N2:N5").Select
    Selection.NumberFormat = "m/d/yyyy"
    PROJECT_DATA.Range("O2:S5").Select
    Selection.NumberFormat = "0.000"
    PROJECT_DATA.Range("W2:Z5").Select
    Selection.NumberFormat = "0"
    PROJECT_DATA.Range("AH2:AH5").Select
    Selection.NumberFormat = "m/d/yyyy"
    PROJECT_DATA.Range("AI2:AJ5").Select
    Selection.NumberFormat = "0.00000"
    PROJECT_DATA.Range("AM2:AM5").Select
    Selection.NumberFormat = "m/d/yyyy h:mm"
End Sub

 Sub Format_SubtotalText_Participating()
    Selection.Font.Bold = True
    Selection.Font.Size = 12
End Sub

 Sub Format_SubtotalAmount_Participating()
    Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Selection.Font.Bold = True
    Selection.ShrinkToFit = True
    Selection.Font.Size = 12
End Sub

 Sub Format_SubtotalText_NonParticipating()
    Selection.Font.Bold = True
    Selection.Font.Size = 12
    Selection.Font.Color = -4165632
End Sub

 Sub Format_SubtotalAmount_NonParticipating()
    Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Selection.Font.Bold = True
    Selection.ShrinkToFit = True
    Selection.Font.Size = 12
    Selection.Font.Color = -4165632
End Sub

Sub FormatComboBox()
    Dim cb As ComboBox
    Set cb = UserForm_LettingResults.cmb_BidderAmount2
    
    cb.AddItem FormatCurrency(1000)

End Sub

Sub Format_PageSetup()
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = "Version " & START.Range("H12").Value
        .LeftFooter = "Printed &D @ &T"
        .CenterFooter = "Page &P of &N"
        .RightFooter = PROJECT_DATA.Range("M2").Value
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .HeaderMargin = Application.InchesToPoints(0.25)
        .FooterMargin = Application.InchesToPoints(0.25)
    End With

End Sub


    Sub Format_QuantitiesThisPeriod()

    Cells.Find(What:="End of Fuel Adjustment Summary", After:=ActiveCell, LookIn:=xlValues _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
     
        Dim rownumber As Integer
        rownumber = ActiveCell.row
        
            row = rownumber
            column = 1
            endrow = rownumber + CountPayItems
            
    Do While row < endrow
        If Cells(row, column + 10).Value = "Yes" Or Cells(row, column + 10).Value = "No" Or Cells(row, column + 10).Value = "Correction Participating" _
        Or Cells(row, column + 10).Value = "Correction Non-Participating" Then

        Range(Cells(row, column), Cells(row, column + 10)).Select              'Format entire row from A to L columns
        Selection.Locked = True
        Selection.HorizontalAlignment = xlCenter
        Selection.VerticalAlignment = xlCenter
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
        Selection.Font.Size = 16
        Selection.ShrinkToFit = True
         
        Cells(row, column + 1).Select                   'Column B - Pay Item Description
        Selection.HorizontalAlignment = xlLeft
        Selection.VerticalAlignment = xlCenter
        Selection.WrapText = True
        Selection.Font.Size = 14
        
        Cells(row, column + 2).Select                   'Current quantity
        Selection.ShrinkToFit = True
        Selection.Font.Size = 16
        Selection.NumberFormat = "0.000"

        Range(Cells(row, column + 3), Cells(row, column + 4)).Select                  'Column E - Quantities / Allowed To Date
        Selection.ShrinkToFit = True
        Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
        Range(Cells(row, column + 5), Cells(row, column + 9)).Select                   'Type
        Selection.Merge
        Selection.WrapText = True
        
        Range(Cells(row, column + 10), Cells(row, column + 11)).Select                   'Participating Status
        Selection.Merge
        Selection.WrapText = True
        Selection.HorizontalAlignment = xlCenter
        Selection.VerticalAlignment = xlCenter
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
        Selection.Font.Size = 14
        
        End If
            row = row + 1
    Loop
    Rows("53:999").EntireRow.AutoFit

End Sub


  Sub Format_FuelSummary()

    Cells.Find(What:="End of CAD Document", After:=ActiveCell, LookIn:=xlValues _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
     
        Dim rownumber As Integer
        rownumber = ActiveCell.row
        
            row = rownumber
            column = 1
            endrow = rownumber + CountPayItems
            
    Do While row < endrow
        If Cells(row, column + 9).Value = "Yes" And Cells(row, column).Value <> "" Or Cells(row, column + 9).Value = "No" And Cells(row, column).Value <> "" Then

        Range(Cells(row, column), Cells(row, column + 9)).Select              'Format entire row from A to L columns
        Selection.Locked = True
        Selection.HorizontalAlignment = xlCenter
        Selection.VerticalAlignment = xlCenter
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
        Selection.Font.Size = 16
        Selection.ShrinkToFit = True
         
        Cells(row, column + 1).Select                   'Column B - Pay Item Description
        Selection.HorizontalAlignment = xlLeft
        Selection.VerticalAlignment = xlCenter
        Selection.WrapText = True
        Selection.Font.Size = 14
        
        Cells(row, column + 3).Select                   'Current quantity
        Selection.ShrinkToFit = True
        Selection.Font.Size = 16
        Selection.NumberFormat = "0.00"
    
        Range(Cells(row, column + 5), Cells(row, column + 8)).Select                   'Type
        Selection.Merge
        'Selection.WrapText = True
        'Selection.EntireRow.AutoFit
        
        End If
            row = row + 1
    Loop
    Rows("53:999").EntireRow.AutoFit

End Sub

Sub Format_MaterialsReport()
    Material_Report.Select
            row = 7
            column = 1
            endrow = CountPayItems
    
    Do While row < endrow
        If Cells(row, column + 7).Value = "Yes" Or Cells(row, column + 7).Value = "No" Then
        Range(Cells(row, column), Cells(row, column + 6)).Select               'Column A to F formatting gridlines
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
        Selection.HorizontalAlignment = xlCenter
        Selection.VerticalAlignment = xlCenter
        Selection.WrapText = True
        Selection.Font.Size = 12
        End If
        
        If Cells(row, column + 7).Value = "Yes" Or Cells(row, column + 7).Value = "No" Then
            Cells(row, column + 2).Select
            Selection.NumberFormat = "#,##0.000"
            Selection.ShrinkToFit = True
        End If
        
        If Cells(row, column + 7).Value = "Yes" Or Cells(row, column + 7).Value = "No" Then
            Cells(row, column + 4).Select
            Cells(row, column + 4).Value = Cells(row, column + 3).Value
            Selection.NumberFormat = "#,##0.000"
        End If
        
        If Cells(row, column + 7).Value = "Yes" Or Cells(row, column + 7).Value = "No" Then
            Range(Cells(row, column + 5), Cells(row, column + 6)).Select
            Selection.ClearContents
            Selection.HorizontalAlignment = xlLeft
            Selection.WrapText = True
        End If
        
        If Cells(row, column + 7).Value = "Yes" Or Cells(row, column + 7).Value = "No" Then
           Cells(row, column + 1).Select
            Selection.HorizontalAlignment = xlLeft
        End If
                   
            row = row + 1
    Loop
    
    
   
     Material_Report.Select
            row = 7
            column = 1
            endrow = CountPayItems
    Do While row < endrow
        If Cells(row, column + 7).Value = "Yes" Or Cells(row, column + 7).Value = "No" Then
            Cells(row, column + 3).Select
            Selection.ClearContents
            Selection.NumberFormat = "#,##0.000"
        End If
            row = row + 1
    Loop
End Sub

Sub Format_902_boxes()
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeLeft).Weight = xlMedium
    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).Weight = xlMedium
    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).Weight = xlMedium
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).Weight = xlMedium
End Sub

Sub Format_BidTabBoxes()
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16775923
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Locked = False
End Sub
