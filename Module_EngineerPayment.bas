Attribute VB_Name = "Module_EngineerPayment"
'
''############################################ CREATE THE ENGINEERINGS INVOICE #############################################################
'
'

Sub EngineerPayment()

    Dim EngineeringFundingType As String

    If PROJECT_DATA.Range("AO2").Value = "State" Then
        EngineeringFundingType = "EngPayInv"
        Else
        EngineeringFundingType = "EngPayInvFedFund"
    End If
  
  Application.ScreenUpdating = False
        Sheets(EngineeringFundingType).Visible = True
        Sheets(EngineeringFundingType).Unprotect
        Sheets(EngineeringFundingType).Select

    Sheets(EngineeringFundingType).Copy After:=Sheets(EngineeringFundingType)        'Move after OSARC Estimate sheet'
    Sheets(EngineeringFundingType & " (2)").Select
   
 ' Check the worksheet to see if a name exist and then create a new name
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    Dim ws4 As Worksheet
    Dim ws5 As Worksheet
    Dim ws6 As Worksheet

    Dim SheetName As String
    Dim SheetName2 As String
    Dim SheetName3 As String
    Dim SheetName4 As String
    Dim SheetName5 As String
    
    SheetName = EngineeringFundingType & "_No_1"
    SheetName2 = EngineeringFundingType & "_No_2"
    SheetName3 = EngineeringFundingType & "_No_3"
    SheetName4 = EngineeringFundingType & "_No_4"
    SheetName5 = EngineeringFundingType & "_No_5"


    SheetExists1 = False
    SheetExists2 = False
    SheetExists3 = False
    SheetExists4 = False
    SheetExists5 = False

'
'Check if the Sheet exists
'
'
    With ThisWorkbook
        For Each ws In .Worksheets
            If ws.Name = SheetName Then
                SheetExists1 = True
                Exit For
            End If
        Next
        For Each ws2 In .Worksheets
            If ws2.Name = SheetName2 Then
                SheetExists2 = True
                Exit For
            End If
        Next
        For Each ws3 In .Worksheets
             If ws3.Name = SheetName3 Then
                SheetExists3 = True
                Exit For
            End If
        Next
        For Each ws4 In .Worksheets
            If ws4.Name = SheetName4 Then
                SheetExists4 = True
                Exit For
            End If
        Next
        For Each ws5 In .Worksheets
            If ws5.Name = SheetName5 Then
                SheetExists5 = True
                Exit For
            End If
        Next

'If the sheet doesn't exists, create a new one
        If SheetExists1 = False Then
            Sheets(EngineeringFundingType & " (2)").Name = SheetName
            Sheets(SheetName).Range("B8").Value = 1
        End If
'
'############################### Copy EngPayInv data ############################################################################################################################
'
' Copy data and create new sheet
        If SheetExists2 = False And SheetExists1 = True Then
            Sheets(EngineeringFundingType & " (2)").Name = SheetName2
            Sheets(SheetName2).Range("B8").Value = 2
            Sheets(SheetName2).Range("D12").Value = Sheets(EngineeringFundingType & "_No_1").Range("G31").Value                 'AMOUNT DUE THIS ESTIMATE
            Sheets(SheetName2).Range("A23").Value = Sheets(EngineeringFundingType & "_No_1").Range("A23").Value                 'COPY ROW 23
            Sheets(SheetName2).Range("B23").Value = Sheets(EngineeringFundingType & "_No_1").Range("B23").Value
            Sheets(SheetName2).Range("C23").Value = Sheets(EngineeringFundingType & "_No_1").Range("C23").Value
            Sheets(SheetName2).Range("D23").Value = Sheets(EngineeringFundingType & "_No_1").Range("D23").Value
            Sheets(SheetName2).Range("G23").Value = Sheets(EngineeringFundingType & "_No_1").Range("G23").Value
            
        End If

        If SheetExists3 = False And SheetExists2 = True And SheetExists1 = True Then
            Sheets(EngineeringFundingType & " (2)").Name = SheetName3
            Sheets(SheetName3).Range("B8").Value = 3
            Sheets(SheetName3).Range("D12").Value = Sheets(EngineeringFundingType & "_No_2").Range("G31").Value
            Sheets(SheetName3).Range("A23").Value = Sheets(EngineeringFundingType & "_No_2").Range("A23").Value                 'COPY ROW 23
            Sheets(SheetName3).Range("B23").Value = Sheets(EngineeringFundingType & "_No_2").Range("B23").Value
            Sheets(SheetName3).Range("C23").Value = Sheets(EngineeringFundingType & "_No_2").Range("C23").Value
            Sheets(SheetName3).Range("D23").Value = Sheets(EngineeringFundingType & "_No_2").Range("D23").Value
            Sheets(SheetName3).Range("G23").Value = Sheets(EngineeringFundingType & "_No_2").Range("G23").Value
            Sheets(SheetName3).Range("A24").Value = Sheets(EngineeringFundingType & "_No_2").Range("A24").Value                 'COPY ROW 24
            Sheets(SheetName3).Range("B24").Value = Sheets(EngineeringFundingType & "_No_2").Range("B24").Value
            Sheets(SheetName3).Range("C24").Value = Sheets(EngineeringFundingType & "_No_2").Range("C24").Value
            Sheets(SheetName3).Range("D24").Value = Sheets(EngineeringFundingType & "_No_2").Range("D24").Value
            Sheets(SheetName3).Range("G24").Value = Sheets(EngineeringFundingType & "_No_2").Range("G24").Value
            Sheets(SheetName3).Range("A25").Value = Sheets(EngineeringFundingType & "_No_2").Range("A25").Value                 'COPY ROW 25
            Sheets(SheetName3).Range("B25").Value = Sheets(EngineeringFundingType & "_No_2").Range("B25").Value
            Sheets(SheetName3).Range("C25").Value = Sheets(EngineeringFundingType & "_No_2").Range("C25").Value
            Sheets(SheetName3).Range("D25").Value = Sheets(EngineeringFundingType & "_No_2").Range("D25").Value
            Sheets(SheetName3).Range("G25").Value = Sheets(EngineeringFundingType & "_No_2").Range("G25").Value
            Sheets(SheetName3).Range("A26").Value = Sheets(EngineeringFundingType & "_No_2").Range("A26").Value                 'COPY ROW 26
            Sheets(SheetName3).Range("B26").Value = Sheets(EngineeringFundingType & "_No_2").Range("B26").Value
            Sheets(SheetName3).Range("C26").Value = Sheets(EngineeringFundingType & "_No_2").Range("C26").Value
            Sheets(SheetName3).Range("D26").Value = Sheets(EngineeringFundingType & "_No_2").Range("D26").Value
            Sheets(SheetName3).Range("G26").Value = Sheets(EngineeringFundingType & "_No_2").Range("G26").Value
        End If

        If SheetExists4 = False And SheetExists3 = True And SheetExists2 = True And SheetExists1 = True Then
            Sheets(EngineeringFundingType & " (2)").Name = SheetName4
            Sheets(SheetName4).Range("B8").Value = 4
            Sheets(SheetName4).Range("D12").Value = Sheets(EngineeringFundingType & "_No_3").Range("G31").Value
            Sheets(SheetName4).Range("A23").Value = Sheets(EngineeringFundingType & "_No_3").Range("A23").Value                 'COPY ROW 23
            Sheets(SheetName4).Range("B23").Value = Sheets(EngineeringFundingType & "_No_3").Range("B23").Value
            Sheets(SheetName4).Range("C23").Value = Sheets(EngineeringFundingType & "_No_3").Range("C23").Value
            Sheets(SheetName4).Range("D23").Value = Sheets(EngineeringFundingType & "_No_3").Range("D23").Value
            Sheets(SheetName4).Range("G23").Value = Sheets(EngineeringFundingType & "_No_3").Range("G23").Value
            Sheets(SheetName4).Range("A24").Value = Sheets(EngineeringFundingType & "_No_3").Range("A24").Value                 'COPY ROW 24
            Sheets(SheetName4).Range("B24").Value = Sheets(EngineeringFundingType & "_No_3").Range("B24").Value
            Sheets(SheetName4).Range("C24").Value = Sheets(EngineeringFundingType & "_No_3").Range("C24").Value
            Sheets(SheetName4).Range("D24").Value = Sheets(EngineeringFundingType & "_No_3").Range("D24").Value
            Sheets(SheetName4).Range("G24").Value = Sheets(EngineeringFundingType & "_No_3").Range("G24").Value
            Sheets(SheetName4).Range("A25").Value = Sheets(EngineeringFundingType & "_No_3").Range("A25").Value                 'COPY ROW 25
            Sheets(SheetName4).Range("B25").Value = Sheets(EngineeringFundingType & "_No_3").Range("B25").Value
            Sheets(SheetName4).Range("C25").Value = Sheets(EngineeringFundingType & "_No_3").Range("C25").Value
            Sheets(SheetName4).Range("D25").Value = Sheets(EngineeringFundingType & "_No_3").Range("D25").Value
            Sheets(SheetName4).Range("G25").Value = Sheets(EngineeringFundingType & "_No_3").Range("G25").Value
            Sheets(SheetName4).Range("A26").Value = Sheets(EngineeringFundingType & "_No_3").Range("A26").Value                 'COPY ROW 26
            Sheets(SheetName4).Range("B26").Value = Sheets(EngineeringFundingType & "_No_3").Range("B26").Value
            Sheets(SheetName4).Range("C26").Value = Sheets(EngineeringFundingType & "_No_3").Range("C26").Value
            Sheets(SheetName4).Range("D26").Value = Sheets(EngineeringFundingType & "_No_3").Range("D26").Value
            Sheets(SheetName4).Range("G26").Value = Sheets(EngineeringFundingType & "_No_3").Range("G26").Value
        End If

        If SheetExists5 = False And SheetExists4 = True And SheetExists3 = True And SheetExists2 = True And SheetExists1 = True Then
            Sheets(EngineeringFundingType & " (2)").Name = SheetName5
            Sheets(SheetName5).Range("B8").Value = 5
            Sheets(SheetName5).Range("D12").Value = Sheets(EngineeringFundingType & "_No_4").Range("G31").Value
            Sheets(SheetName5).Range("A23").Value = Sheets(EngineeringFundingType & "_No_4").Range("A23").Value                 'COPY ROW 23
            Sheets(SheetName5).Range("B23").Value = Sheets(EngineeringFundingType & "_No_4").Range("B23").Value
            Sheets(SheetName5).Range("C23").Value = Sheets(EngineeringFundingType & "_No_4").Range("C23").Value
            Sheets(SheetName5).Range("D23").Value = Sheets(EngineeringFundingType & "_No_4").Range("D23").Value
            Sheets(SheetName5).Range("G23").Value = Sheets(EngineeringFundingType & "_No_4").Range("G23").Value
            Sheets(SheetName5).Range("A24").Value = Sheets(EngineeringFundingType & "_No_4").Range("A24").Value                 'COPY ROW 24
            Sheets(SheetName5).Range("B24").Value = Sheets(EngineeringFundingType & "_No_4").Range("B24").Value
            Sheets(SheetName5).Range("C24").Value = Sheets(EngineeringFundingType & "_No_4").Range("C24").Value
            Sheets(SheetName5).Range("D24").Value = Sheets(EngineeringFundingType & "_No_4").Range("D24").Value
            Sheets(SheetName5).Range("G24").Value = Sheets(EngineeringFundingType & "_No_4").Range("G24").Value
            Sheets(SheetName5).Range("A25").Value = Sheets(EngineeringFundingType & "_No_4").Range("A25").Value                 'COPY ROW 25
            Sheets(SheetName5).Range("B25").Value = Sheets(EngineeringFundingType & "_No_4").Range("B25").Value
            Sheets(SheetName5).Range("C25").Value = Sheets(EngineeringFundingType & "_No_4").Range("C25").Value
            Sheets(SheetName5).Range("D25").Value = Sheets(EngineeringFundingType & "_No_4").Range("D25").Value
            Sheets(SheetName5).Range("G25").Value = Sheets(EngineeringFundingType & "_No_4").Range("G25").Value
            Sheets(SheetName5).Range("A26").Value = Sheets(EngineeringFundingType & "_No_4").Range("A26").Value                 'COPY ROW 26
            Sheets(SheetName5).Range("B26").Value = Sheets(EngineeringFundingType & "_No_4").Range("B26").Value
            Sheets(SheetName5).Range("C26").Value = Sheets(EngineeringFundingType & "_No_4").Range("C26").Value
            Sheets(SheetName5).Range("D26").Value = Sheets(EngineeringFundingType & "_No_4").Range("D26").Value
            Sheets(SheetName5).Range("G26").Value = Sheets(EngineeringFundingType & "_No_4").Range("G26").Value
        End If
    End With
    ActiveSheet.Select
    ActiveSheet.Protect

    Sheets(EngineeringFundingType).Visible = False
    Application.ScreenUpdating = True
    ActiveSheet.Range("B7").Select

End Sub





Sub CAD_DataEngineerPayment()
    
    Dim LastCAD_Completed As String
    Dim i As Integer

    i = 1
        Do While SheetExists("CAD_No_" & i)
                    i = i + 1
        
            LastCAD_Completed = "CAD_No_" & i - 1
        
            ActiveSheet.Range("B31").Value = Sheets(LastCAD_Completed).Range("H31")     ' Days working on Project this period
            ActiveSheet.Range("B32").Value = Sheets(LastCAD_Completed).Range("I31")   ' Previous Working Days
            ActiveSheet.Range("B34").Value = Sheets(LastCAD_Completed).Range("J32")      ' Project Percent Completion
            ActiveSheet.Range("B35").Value = Sheets(LastCAD_Completed).Range("J33")      ' Project Elapse Time
                                
        Loop
        

End Sub


'
'#############################################################          CALCULATE ENGINEERING INVOICE             #####################################################

Sub Calculate_EngPayInv()

    ActiveSheet.Select
    ActiveSheet.Unprotect
    Application.ScreenUpdating = False
' WARNING MESSAGE TO SELECT THE APPROPRIATE INVOICE TO CREATE
    If MsgBox("Did you select the option for the 'Invoice Payment for' ", vbYesNo, "Save changes") = vbYes Then

' IF Preliminary then calculate only preliminary of 4.8%
    If Range("B7").Value = "Preliminary" Then
            ActiveSheet.Range("A18").Value = "Preliminary"
            ActiveSheet.Range("A23").Value = "=CONCATENATE(""Invoice #"",R8C2)"
            ActiveSheet.Range("B23").Value = "(Preliminary Engineering)"
            ActiveSheet.Range("C23").Value = Range("D11")
            ActiveSheet.Range("D23").Value = "  X  4.8%  = "
            Range("G23").Value = Round(Range("C23").Value * DATA_VALIDATION.Range("T2").Value, 2)

        ElseIf Range("B7").Value = "Construction @ 25%" Then
           ActiveSheet.Range("A18").Value = "Construction"
            ActiveSheet.Range("A24").Value = "=CONCATENATE(""Invoice #"",R8C2)"
            ActiveSheet.Range("B24").Value = "CONSTRUCTION @ 25%"
            ActiveSheet.Range("C24").Value = Range("D11")
            ActiveSheet.Range("D24").Value = "  X  1.8%  = "
            Range("G24").Value = Round(Range("C24").Value * DATA_VALIDATION.Range("T3").Value, 2)
            
        ElseIf Range("B7").Value = "Construction @ 50%" Then

            ActiveSheet.Range("A18").Value = "Construction"
            ActiveSheet.Range("A25").Value = "=CONCATENATE(""Invoice #"",R8C2)"
            ActiveSheet.Range("B25").Value = "CONSTRUCTION @ 50%"
            ActiveSheet.Range("C25").Value = Range("D11")
            ActiveSheet.Range("D25").Value = "  X  1.8%  = "
            Range("G25").Value = Round(Range("C25").Value * DATA_VALIDATION.Range("T4").Value, 2)
            
          ElseIf Range("B7").Value = "Construction @ 75%" Then
            ActiveSheet.Range("A18").Value = "Construction"
            ActiveSheet.Range("A26").Value = "=CONCATENATE(""Invoice #"",R8C2)"
            ActiveSheet.Range("B26").Value = "CONSTRUCTION @ 75%"
            ActiveSheet.Range("C26").Value = Range("D11")
            ActiveSheet.Range("D26").Value = "  X  1.8%  = "
            Range("G26").Value = Round(Range("C26").Value * DATA_VALIDATION.Range("T5").Value, 2)
          
          ElseIf Range("B7").Value = "Construction @ 100%" Then

            ActiveSheet.Range("A18").Value = "Construction FINAL"
            ActiveSheet.Range("A27").Value = "=CONCATENATE(""Invoice #"",R8C2)"
            ActiveSheet.Range("B27").Value = "CONSTRUCTION @ 100%"
            ActiveSheet.Range("C27").Value = Range("G11")                   ' Changing cells to final here
            ActiveSheet.Range("D27").Value = "  12.0% of the FINAL remaining = "
            Range("G27").Value = "=(R27C3*Data_Validation!R[-21]C[7])-(SUM(R[-4]C:R[-1]C))"
    End If

Call CAD_DataEngineerPayment

Range("C23:C27").Select
Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
Range("G23:G27").Select
Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
ActiveSheet.Protect

    End If
End Sub
