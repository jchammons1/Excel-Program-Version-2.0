Attribute VB_Name = "Module_CAD_Calculate"
Option Explicit
            
    Dim UnitCost As Double
    Dim Participating As String
    Dim Payitem As String
    Dim FuelCode As String
            
    Dim CurrentQuantity As Double
    Dim CurrentCost As Currency
            
    Dim PastQuantity As Double
    Dim PastCost As Currency
            
    Dim TotalQuantity As Double
    Dim TotalCost As Currency
            
    Dim row As Integer
    Dim column As Integer
    Dim endrow As Integer
    

    Dim Active_CAD As String
 

Function CalculateCurrentCost(CurrentQuantity As Double, UnitCost As Double) As Double
    ' Calculate current accumulative cost in Column BN
    CalculateCurrentCost = Round(CurrentQuantity, 3) * Round(UnitCost, 2)
End Function

Function CopyCurrentQuantity(CurrentQuantity As Double) As Double
    ' Copy current quantity to Column BM
    CopyCurrentQuantity = Round(CurrentQuantity, 3)
End Function


Function CalculateTotalCost(PastCost As Currency, CurrentCost As Currency) As Currency
    ' Calculate current total cost in Column BN.
    ' Past cost comes from Column BP of the previous CAD
    CalculateTotalCost = Round(Round(PastCost, 3) + Round(CurrentCost, 3), 2)
End Function

Function CalculateTotalQuantity(PastQuantity As Double, CurrentQuantity As Double) As Double
    ' Add Column D to Column BK
    ' Column BK is copied from the previous CAD
    CalculateTotalQuantity = Round(PastQuantity, 3) + Round(CurrentQuantity, 3)
End Function


Sub Calculate_CostQuantity()
'Application.StatusBar = True
'Application.StatusBar = "Macro is running..."

    ActiveSheet.Select
    ActiveSheet.Unprotect Password:="roadway123"
    Application.ScreenUpdating = False
    Active_CAD = ActiveSheet.Name

'' Calculate cost for each pay item
''
''
            row = 53
            column = 1
            endrow = CountPayItems + 53
   
    Do While row < endrow
   ' Dim ParticipatingStatus As String
    
     '   Select Case ParticipatingStatus = Cells(row, column + 49).Value
      '      Case "Yes", "No", "Correction Participating", "Correction Non-Participating"
        If Cells(row, column + 49).Value = "Yes" Or Cells(row, column + 49).Value = "No" Or Cells(row, column + 49).Value = "Correction Participating" _
        Or Cells(row, column + 49).Value = "Correction Non-Participating" Then
            UnitCost = Cells(row, column + 6).Value
            Participating = Cells(row, column + 49).Value
            CurrentQuantity = Cells(row, column + 3).Value
            PastQuantity = Cells(row, column + 62).Value
            PastCost = Cells(row, column + 63).Value
            TotalQuantity = Cells(row, column + 4).Value
            Payitem = Cells(row, column).Value
                    Cells(row, column + 65).Value = CalculateCurrentCost(CurrentQuantity, UnitCost)                     'Current Cost in Column BN
                    Cells(row, column + 66).Value = CalculateTotalQuantity(PastQuantity, CurrentQuantity)               'Total Quantity in Column BO
                    Cells(row, column + 4).Value = CalculateTotalQuantity(PastQuantity, CurrentQuantity)
                               CurrentCost = Cells(row, column + 65).Value
                    Cells(row, column + 67).Value = CalculateTotalCost(PastCost, CurrentCost)            'Current Total Cost  ' AKA current accumulative amount
                    Cells(row, column + 7).Value = CalculateTotalCost(PastCost, CurrentCost)             'Current Total Cost
                    Cells(row, column + 64).Value = CopyCurrentQuantity(CurrentQuantity)
                    Cells(row, column + 11).Value = "=IF(RC[3]=""NA"",""NA"",CONCAT(RC[3],"" "",RC[4]))"
       '  End Select
        End If
        
 ''
 '' Subtotal calculations
 ''
 ''
            PAYITEMTYPE.Visible = True
            PAYITEMTYPE.Select
            Dim RowType As Range                            'Declare a variable for the Pay Item Type and the range of data from the Data Validation sheet
            For Each RowType In Range("A2:A55")            ' Maximum of 100 types the user can add to the program. 4 default types. 50 User Defined and 50 supplemental types.
                                                            ' The original estimates only go to A2 and A55
                                                            
            Sheets(Active_CAD).Select
            If IsEmpty(RowType) = False Then
            
                If Cells(row, column + 1) = "Subtotal Participating " & RowType.Value & " Items" Then
                    Cells(row, column + 7).Select
                    Cells(row, column + 7).Value = "=SUMIFS(C,C[41]," & _
                                     Chr(34) & _
                                     RowType.Value & _
                                     Chr(34) & _
                                     ",C[42],""Yes"")"
                    Call Format_SubtotalAmount_Participating
                End If
                
                If Cells(row, column + 1) = "Subtotal Non-Participating " & RowType.Value & " Items" Then
                    Cells(row, column + 7).Select
                    Cells(row, column + 7).Value = "=SUMIFS(C,C[41]," & _
                                     Chr(34) & _
                                     RowType.Value & _
                                     Chr(34) & _
                                     ",C[42],""No"")"
                    Call Format_SubtotalAmount_NonParticipating
                End If
               
                
            End If
            Next RowType
            
            PAYITEMTYPE_SUPPLEMENTAL.Visible = True
            PAYITEMTYPE_SUPPLEMENTAL.Select
            Dim RowType2 As Range                            'Declare a variable for the Pay Item Type and the range of data from the Data Validation sheet
            For Each RowType2 In Range("A2:A55")            ' Maximum of 100 types the user can add to the program. 4 default types. 50 User Defined and 50 supplemental types.
                                                            ' The original estimates only go to A2 and A55
                                                            
            Sheets(Active_CAD).Select
            If IsEmpty(RowType2) = False Then
            
                If Cells(row, column + 1) = "Subtotal " & RowType2.Value & " Items" Then
                    Cells(row, column + 7).Select
                    Cells(row, column + 7).Value = "=SUMIFS(C,C[41]," & Chr(34) & RowType2.Value & Chr(34) & ")"
                    Call Format_SubtotalAmount_Participating
                End If

            End If
            Next RowType2
 
 ''
 '' Gas and Diesel calculations
 ''
 ''
            DATA_VALIDATION.Visible = True
            DATA_VALIDATION.Select
            Dim FuelCode As Range
            
            For Each FuelCode In Range("H3:H17")
            Sheets(Active_CAD).Select
            
            If Cells(row, column + 10).Value = FuelCode Then
                Select Case FuelCode
                    Case "E"
                        Cells(row, column + 42).Value = "=IFERROR(IF(ABS((RC[-24]-RC[-26])/RC[-26])>=0.05,(RC[-24]-RC[-26])*0.15*RC[-39],0),0)"
                        Cells(row, column + 43).Value = "=IFERROR(IF(ABS((RC[-24]-RC[-26])/RC[-26])>=0.05,(RC[-24]-RC[-26])*0.29*RC[-40],0),0)"
                        Cells(row, column + 45).Value = "=SUM(RC[-3]:RC[-1])"
                        Cells(row, column + 9).Value = "=RC[36]"
                    Case "GT"
                        Cells(row, column + 42).Value = "=IFERROR(IF(ABS((RC[-24]-RC[-26])/RC[-26])>=0.05,(RC[-24]-RC[-26])*0.40*RC[-39],0),0)"
                        Cells(row, column + 43).Value = "=IFERROR(IF(ABS((RC[-24]-RC[-26])/RC[-26])>=0.05,(RC[-24]-RC[-26])*0.62*RC[-40],0),0)"
                        Cells(row, column + 45).Value = "=SUM(RC[-3]:RC[-1])"
                        Cells(row, column + 9).Value = "=RC[36]"
                    Case "GY"
                        Cells(row, column + 42).Value = "=IFERROR(IF(ABS((RC[-24]-RC[-26])/RC[-26])>=0.05,(RC[-24]-RC[-26])*0.57*RC[-39],0),0)"
                        Cells(row, column + 43).Value = "=IFERROR(IF(ABS((RC[-24]-RC[-26])/RC[-26])>=0.05,(RC[-24]-RC[-26])*0.88*RC[-40],0),0)"
                        Cells(row, column + 45).Value = "=SUM(RC[-3]:RC[-1])"
                        Cells(row, column + 9).Value = "=RC[36]"
                    Case "M"
                        Cells(row, column + 42).Value = "=IFERROR(IF(ABS((RC[-24]-RC[-26])/RC[-26])>=0.05,(RC[-24]-RC[-26])*0.028*RC[-39],0),0)"
                        Cells(row, column + 43).Value = "=IFERROR(IF(ABS((RC[-24]-RC[-26])/RC[-26])>=0.05,(RC[-24]-RC[-26])*0.044*RC[-40],0),0)"
                        Cells(row, column + 45).Value = "=SUM(RC[-3]:RC[-1])"
                        Cells(row, column + 9).Value = "=RC[36]"
                    Case "BA1"
                        Cells(row, column + 42).Value = "=IFERROR(IF(ABS((RC[-24]-RC[-26])/RC[-26])>=0.05,(RC[-24]-RC[-26])*0.78*RC[-39],0),0)"
                        Cells(row, column + 43).Value = "=IFERROR(IF(ABS((RC[-24]-RC[-26])/RC[-26])>=0.05,(RC[-24]-RC[-26])*2.57*RC[-40],0),0)"
                        Cells(row, column + 45).Value = "=SUM(RC[-3]:RC[-1])"
                        Cells(row, column + 9).Value = "=RC[36]"
                    Case "D"
                        Cells(row, column + 42).Value = "=IFERROR(IF(ABS((RC[-24]-RC[-26])/RC[-26])>=0.05,(RC[-24]-RC[-26])*0.15*RC[-39],0),0)"
                        Cells(row, column + 43).Value = "=IFERROR(IF(ABS((RC[-24]-RC[-26])/RC[-26])>=0.05,(RC[-24]-RC[-26])*0.49*RC[-40],0),0)"
                        Cells(row, column + 45).Value = "=SUM(RC[-3]:RC[-1])"
                        Cells(row, column + 9).Value = "=RC[36]"
                    Case "C"
                        Cells(row, column + 42).Value = "=IFERROR(IF(ABS((RC[-24]-RC[-26])/RC[-26])>=0.05,(RC[-24]-RC[-26])*0.15*RC[-39],0),0)"
                        Cells(row, column + 43).Value = "=IFERROR(IF(ABS((RC[-24]-RC[-26])/RC[-26])>=0.05,(RC[-24]-RC[-26])*0.11*RC[-40],0),0)"
                        Cells(row, column + 45).Value = "=SUM(RC[-3]:RC[-1])"
                        Cells(row, column + 9).Value = "=RC[36]"
                    Case "S"
                        Cells(row, column + 42).Value = "=IFERROR(IF(ABS((RC[-24]-RC[-26])/RC[-26])>=0.05,(RC[-24]-RC[-26])*RC[23]*13/1000,0),0)"
                        Cells(row, column + 43).Value = "=IFERROR(IF(ABS((RC[-24]-RC[-26])/RC[-26])>=0.05,(RC[-24]-RC[-26])*RC[22]*11/1000,0),0)"
                        Cells(row, column + 45).Value = "=SUM(RC[-3]:RC[-1])"
                        Cells(row, column + 9).Value = "=RC[36]"
                    Case "A2", "A3", "A4", "A5", "A6"
                        Cells(row, column + 45).Value = "=SUM(RC[-3]:RC[-1])"
                        Cells(row, column + 9).Value = "=RC[36]"
                End Select
            End If
        Next FuelCode
 ''
 '' Asphalt Binder calculations
 ''
 ''
            DATA_VALIDATION.Visible = True
            DATA_VALIDATION.Select
            Dim BinderType As Range
            
            For Each BinderType In Range("I3:I6")
                                                             
            Sheets(Active_CAD).Select
            If Cells(row, column + 21).Value = BinderType And Cells(row, column + 10).Value = "BA1" Or _
            Cells(row, column + 21).Value = BinderType And Cells(row, column + 10).Value = "A6" Then
                Select Case BinderType
                    Case "PG-64-22"
                        Cells(row, column + 44).Value = "=IFERROR(IF(ABS((RC[-12]-RC[-22])/RC[-22])>=0.05,((RC[-12]-RC[-22])*2.375*RC[-24]*RC[-41]),0),0)"
                    Case "PG-67-22"
                        Cells(row, column + 44).Value = "=IFERROR(IF(ABS((RC[-11]-RC[-21])/RC[-21])>=0.05,((RC[-11]-RC[-21])*2.375*RC[-24]*RC[-41]),0),0)"
                    Case "PG-76-22"
                        Cells(row, column + 44).Value = "=IFERROR(IF(ABS((RC[-10]-RC[-20])/RC[-20])>=0.05,((RC[-10]-RC[-20])*2.375*RC[-24]*RC[-41]),0),0)"
                    Case "PG-82-22"
                        Cells(row, column + 44).Value = "=IFERROR(IF(ABS((RC[-9]-RC[-19])/RC[-19])>=0.05,((RC[-9]-RC[-19])*2.375*RC[-24]*RC[-41]),0),0)"
                End Select
            End If
            Next BinderType
            
 ''
 '' Emulsion calculations
 ''
 ''
            DATA_VALIDATION.Visible = True
            DATA_VALIDATION.Select
            Dim EmulsionType As Range
            
            For Each EmulsionType In Range("I7:I12")
                                                              
            Sheets(Active_CAD).Select
            
            If Cells(row, column + 21).Value = EmulsionType And Cells(row, column + 10).Value = "A2" Or _
            Cells(row, column + 21).Value = EmulsionType And Cells(row, column + 10).Value = "A3" Then
                Select Case EmulsionType
                    Case "SS-1"
                        Cells(row, column + 44).Value = "=IFERROR(IF(ABS((RC[-8]-RC[-18])/RC[-18])>=0.05,((RC[-8]-RC[-18])*1.0*RC[-41]),0),0)"
                    Case "CRS-2"
                        Cells(row, column + 44).Value = "=IFERROR(IF(ABS((RC[-7]-RC[-17])/RC[-17])>=0.05,((RC[-7]-RC[-17])*1.0*RC[-41]),0),0)"
                    Case "CRS-2P"
                        Cells(row, column + 44).Value = "=IFERROR(IF(ABS((RC[-6]-RC[-16])/RC[-16])>=0.05,((RC[-6]-RC[-16])*1.0*RC[-41]),0),0)"
                    Case "EA-1"
                        Cells(row, column + 44).Value = "=IFERROR(IF(ABS((RC[-5]-RC[-15])/RC[-15])>=0.05,((RC[-5]-RC[-15])*1.0*RC[-41]),0),0)"
                    Case "CSS-1_UN"
                        Cells(row, column + 44).Value = "=IFERROR(IF(ABS((RC[-4]-RC[-14])/RC[-14])>=0.05,((RC[-4]-RC[-14])*1.0*RC[-41]),0),0)"
                    Case "CSS-1"
                        Cells(row, column + 44).Value = "=IFERROR(IF(ABS((RC[-3]-RC[-13])/RC[-13])>=0.05,((RC[-3]-RC[-13])*1.0*RC[-41]),0),0)"
                End Select
            End If

            If Cells(row, column + 21).Value = EmulsionType And Cells(row, column + 10).Value = "A4" Then
                Select Case EmulsionType
                    Case "SS-1"
                        Cells(row, column + 44).Value = "=IFERROR(IF(ABS((RC[-8]-RC[-18])/RC[-18])>=0.05,((RC[-8]-RC[-18])*0.25*RC[-41]),0),0)"
                    Case "CRS-2"
                        Cells(row, column + 44).Value = "=IFERROR(IF(ABS((RC[-7]-RC[-17])/RC[-17])>=0.05,((RC[-7]-RC[-17])*0.25*RC[-41]),0),0)"
                    Case "CRS-2P"
                        Cells(row, column + 44).Value = "=IFERROR(IF(ABS((RC[-6]-RC[-16])/RC[-16])>=0.05,((RC[-6]-RC[-16])*0.25*RC[-41]),0),0)"
                    Case "EA-1"
                        Cells(row, column + 44).Value = "=IFERROR(IF(ABS((RC[-5]-RC[-15])/RC[-15])>=0.05,((RC[-5]-RC[-15])*0.25*RC[-41]),0),0)"
                    Case "CSS-1_UN"
                        Cells(row, column + 44).Value = "=IFERROR(IF(ABS((RC[-4]-RC[-14])/RC[-14])>=0.05,((RC[-4]-RC[-14])*0.25*RC[-41]),0),0)"
                    Case "CSS-1"
                        Cells(row, column + 44).Value = "=IFERROR(IF(ABS((RC[-3]-RC[-13])/RC[-13])>=0.05,((RC[-3]-RC[-13])*0.25*RC[-41]),0),0)"
                End Select
            End If

            If Cells(row, column + 21).Value = EmulsionType And Cells(row, column + 10).Value = "A5" Then
                Select Case EmulsionType
                    Case "SS-1"
                        Cells(row, column + 44).Value = "=IFERROR(IF(ABS((RC[-8]-RC[-18])/RC[-18])>=0.05,((RC[-8]-RC[-18])*0.50*RC[-41]),0),0)"
                    Case "CRS-2"
                        Cells(row, column + 44).Value = "=IFERROR(IF(ABS((RC[-7]-RC[-17])/RC[-17])>=0.05,((RC[-7]-RC[-17])*0.50*RC[-41]),0),0)"
                    Case "CRS-2P"
                        Cells(row, column + 44).Value = "=IFERROR(IF(ABS((RC[-6]-RC[-16])/RC[-16])>=0.05,((RC[-6]-RC[-16])*0.50*RC[-41]),0),0)"
                    Case "EA-1"
                        Cells(row, column + 44).Value = "=IFERROR(IF(ABS((RC[-5]-RC[-15])/RC[-15])>=0.05,((RC[-5]-RC[-15])*0.50*RC[-41]),0),0)"
                    Case "CSS-1_UN"
                        Cells(row, column + 44).Value = "=IFERROR(IF(ABS((RC[-4]-RC[-14])/RC[-14])>=0.05,((RC[-4]-RC[-14])*0.50*RC[-41]),0),0)"
                    Case "CSS-1"
                        Cells(row, column + 44).Value = "=IFERROR(IF(ABS((RC[-3]-RC[-13])/RC[-13])>=0.05,((RC[-3]-RC[-13])*0.50*RC[-41]),0),0)"
                End Select
            End If
            
            Next EmulsionType
            
            
        row = row + 1
    Loop

   Call Calculate_Totals
   Call Mobilization_Traffic
   Call Over3000_ChangeOrderNeeded
   Call Over10000_ChangeOrderNeeded
   Call Quantities_ThisPeriod
   Call Format_QuantitiesThisPeriod
   Call SetPrintArea_CalculateCAD

    DATA_VALIDATION.Visible = False
    PAYITEMTYPE.Visible = False
    PAYITEMTYPE_SUPPLEMENTAL.Visible = False
    
    Sheets(Active_CAD).Range("G7").Select
    Sheets(Active_CAD).Protect Password:="roadway123"
    Application.ScreenUpdating = True
   ' Application.StatusBar = False
End Sub


 Sub Calculate_Totals()
       
''
 '' Total cost
 ''
 ''
        Range("A52").Select
        Dim ws As Worksheet
        Set ws = ActiveSheet
        
        If Not ws.Cells.Find("PROJECT PARTICIPATING TOTAL", SearchOrder:=xlByRows, SearchDirection:=xlNext) Is Nothing Then
            Cells.Find(What:="PROJECT PARTICIPATING TOTAL", After:=ActiveCell, LookIn _
            :=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
            xlNext, MatchCase:=False, SearchFormat:=False).Activate
                    ActiveCell.Offset(0, 6).Range("A1").Select
                    ActiveCell.FormulaR1C1 = "=SUMIFS(C,C[42],""Yes"")+SUMIFS(C,C[42],""Correction Participating"")"
                        Call Format_SubtotalAmount_Participating
        End If
        
        Range("A52").Select
        If Not ws.Cells.Find("PROJECT NON-PARTICIPATING TOTAL", SearchOrder:=xlByRows, SearchDirection:=xlNext) Is Nothing Then
            Cells.Find(What:="PROJECT NON-PARTICIPATING TOTAL", After:=ActiveCell, LookIn _
            :=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
            xlNext, MatchCase:=False, SearchFormat:=False).Activate
                    ActiveCell.Offset(0, 6).Range("A1").Select
                    ActiveCell.FormulaR1C1 = "=SUMIFS(C,C[42],""No"")+SUMIFS(C,C[42],""Correction Non-Participating"")"
                        Call Format_SubtotalAmount_NonParticipating
        End If
        
        Range("A52").Select
        If Not ws.Cells.Find("FUEL & MATERIAL ADJUSTMENTS PARTICIPATING", SearchOrder:=xlByRows, SearchDirection:=xlNext) Is Nothing Then
            Cells.Find(What:="FUEL & MATERIAL ADJUSTMENTS PARTICIPATING", After:=ActiveCell, LookIn _
            :=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
            xlNext, MatchCase:=False, SearchFormat:=False).Activate
                    ActiveCell.Offset(0, 8).Range("A1").Select
                    ActiveCell.FormulaR1C1 = "=SUMIFS(C,C[40],""Yes"")+SUMIFS(C,C[40],""Correction Participating"")"
                        Call Format_SubtotalAmount_Participating
        End If
        
        Range("A52").Select
        If Not ws.Cells.Find("FUEL & MATERIAL ADJUSTMENTS NON-PARTICIPATING", SearchOrder:=xlByRows, SearchDirection:=xlNext) Is Nothing Then
            Cells.Find(What:="FUEL & MATERIAL ADJUSTMENTS NON-PARTICIPATING", After:=ActiveCell, LookIn _
            :=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
            xlNext, MatchCase:=False, SearchFormat:=False).Activate
                    ActiveCell.Offset(0, 8).Range("A1").Select
                    ActiveCell.FormulaR1C1 = "=SUMIFS(C,C[40],""No"")+SUMIFS(C,C[40],""Correction Non-Participating"")"
                        Call Format_SubtotalAmount_NonParticipating
        End If
       
        Range("A52").Select
        If Not ws.Cells.Find("FUEL & MATERIAL ADJUSTMENTS PARTICIPATING AND NON-PARTICIPATING TOTAL", SearchOrder:=xlByRows, SearchDirection:=xlNext) Is Nothing Then
            Cells.Find(What:="FUEL & MATERIAL ADJUSTMENTS PARTICIPATING AND NON-PARTICIPATING TOTAL", After:=ActiveCell, LookIn _
            :=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
            xlNext, MatchCase:=False, SearchFormat:=False).Activate
                    ActiveCell.Offset(0, 8).Range("A1").Select
                    ActiveCell.FormulaR1C1 = "=SUMIFS(C,C[40],""Yes"")+SUMIFS(C,C[40],""Correction Participating"")+SUMIFS(C,C[40],""No"")+SUMIFS(C,C[40],""Correction Non-Participating"")"
                        Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                        Selection.Font.Bold = True
                        Selection.ShrinkToFit = True
                        Selection.Font.Size = 16
        End If
          
        Range("A52").Select
        If Not ws.Cells.Find("PROJECT PARTICIPATING AND NON-PARTICIPATING TOTAL", SearchOrder:=xlByRows, SearchDirection:=xlNext) Is Nothing Then
            Cells.Find(What:="PROJECT PARTICIPATING AND NON-PARTICIPATING TOTAL", After:=ActiveCell, LookIn _
            :=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
            xlNext, MatchCase:=False, SearchFormat:=False).Activate
                    ActiveCell.Offset(0, 6).Range("A1").Select
                    ActiveCell.FormulaR1C1 = "=SUMIFS(C,C[42],""Yes"")+SUMIFS(C,C[42],""No"")+SUMIFS(C,C[42],""Correction Participating"")+SUMIFS(C,C[42],""Correction Non-Participating"")+R2C212"       'R2C212 = TOTAL FUEL COST TO DATE
                        Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                        Selection.Font.Bold = True
                        Selection.ShrinkToFit = True
                        Selection.Font.Size = 16
        End If
 
End Sub


 Sub Over3000_ChangeOrderNeeded()
' Highlight individual pay items that are $3,000 or more for a CCR
    row = 53
    column = 1
    endrow = CountPayItems + 53
    
    Do While row < endrow

        If IsNumeric(Cells(row, column + 8).Value) And IsNumeric(Cells(row, column + 7).Value) And Cells(row, column + 49).Value = "Yes" _
            Or IsNumeric(Cells(row, column + 8).Value) And IsNumeric(Cells(row, column + 7).Value) And Cells(row, column + 49).Value = "No" Then
            If (Cells(row, column + 7).Value - Cells(row, column + 8).Value) >= 3000 Then
                Cells(row, column + 7).Select
                Selection.Font.Bold = True
                Selection.Interior.Color = 65535
            End If
        End If
        
        
    row = row + 1
    Loop
End Sub


 Sub Over10000_ChangeOrderNeeded()
'Find cell with "PROJECT PARTICIPATING AND NON-PARTICIPATING TOTAL" and then find the cell to the right in Column I
    Cells.Find(What:="PROJECT PARTICIPATING AND NON-PARTICIPATING TOTAL", After:=ActiveCell, LookIn:=xlValues _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Select
    ActiveCell.Offset(0, 6).Range("A1").Select
    ' Determine if the amount is $10,000 more than the contracted amount.  then a CCR is needed.
        If ActiveCell.Value - ActiveCell.Offset(0, 1).Range("A1").Value >= 10000 Then
            Selection.Font.Bold = True
            Selection.Interior.Color = 65535
        End If
End Sub


Sub Quantities_ThisPeriod()
' Find the end of the fuel adjustment
Cells.Find(What:="End of Fuel Adjustment Summary", After:=ActiveCell, LookIn:=xlValues _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Select
    ActiveCell.Offset(2, 0).Range("A1").Select
    ActiveCell.Range("A1:L999").Clear


    ActiveSheet.Select
    Dim row As Integer
    Dim column As Integer
    Dim endrow As Integer
    
            row = 53
            column = 1
            endrow = CountPayItems + 53
   Do While row < endrow
'Calculate the the previous months totals
     '   Dim ParticipatingStatus As String
           ' Select Case ParticipatingStatus = Cells(row, column + 49).Value And Cells(row, column + 3).Value And Cells(row, column + 3).Value
           '     Case "Yes", "No", "Correction Participating", "Correction Non-Participating" And 0 And ""
                

    If Cells(row, column + 49).Value = "Yes" And Cells(row, column + 3).Value <> 0 And Cells(row, column + 3).Value <> "" _
    Or Cells(row, column + 49).Value = "No" And Cells(row, column + 3).Value <> 0 And Cells(row, column + 3).Value <> "" _
    Or Cells(row, column + 49).Value = "Correction Participating" And Cells(row, column + 3).Value <> 0 And Cells(row, column + 3).Value <> "" _
    Or Cells(row, column + 49).Value = "Correction Non-Participating" And Cells(row, column + 3).Value <> 0 And Cells(row, column + 3).Value <> "" Then

            ActiveCell.Value = Cells(row, column).Value
            ActiveCell.Offset(0, 1).Value = Cells(row, column + 1).Value
            ActiveCell.Offset(0, 2).Value = Cells(row, column + 3).Value
            ActiveCell.Offset(0, 3).Value = Cells(row, column + 6).Value
            ActiveCell.Offset(0, 4).Value = Cells(row, column + 3).Value * Cells(row, column + 6).Value
            ActiveCell.Offset(0, 5).Value = Cells(row, column + 48).Value
            ActiveCell.Offset(0, 10).Value = Cells(row, column + 49).Value
            ActiveCell.Offset(1, 0).Select
        End If
      '  End Select
         
        row = row + 1
    Loop
     ActiveCell.Offset(1, 0).Select
     ActiveCell.Value = "End of Monthly Items Summary"
     ActiveCell.Offset(1, 0).Select
End Sub

