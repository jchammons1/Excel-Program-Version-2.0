Attribute VB_Name = "Module_AverageUnitCost"
 
Public Function DatabaseTempRegion() As Range
 Set DatabaseTempRegion = AvgUnitCost.Range("A2").CurrentRegion
End Function

Public Function DataBaseTempRowCount() As Integer
  DataBaseTempRowCount = DatabaseTempRegion.Rows.Count
End Function

 '*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'*********************************** This section automatically takes the data from the Project Data and updates the unit cost *************>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'
'
'
 Sub AverageUnitCost()
    AvgUnitCost.Visible = True
    AvgUnitCost.Unprotect
    AvgUnitCost.Select
    AvgUnitCost.Range("A2:CS9999").Clear
    
    Database.Visible = True
    Database.Unprotect
    Database.Select
    Database.Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    AvgUnitCost.Select
    AvgUnitCost.Range("A2").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Dim endrow_Temp As Integer
    Dim row As Integer
    Dim column As Integer
    
            row = 2
            column = 1
            endrow_Temp = DataBaseTempRowCount + 1
    Do While row < endrow_Temp

        If Cells(row, column + 51).Value = "NA" Or Cells(row, column + 51).Value = 0 Then
            Cells(row, column + 51) = PROJECT_DATA.Range("AI5").Value         'Latitude
            Cells(row, column + 51).Select
            Selection.NumberFormat = "0.00000"
        End If
        
        If Cells(row, column + 52).Value = "NA" Or Cells(row, column + 52).Value = 0 Then
            Cells(row, column + 52) = PROJECT_DATA.Range("AJ5").Value         'Longitude
            Cells(row, column + 52).Select
            Selection.NumberFormat = "0.00000"
        End If
        If Cells(row, column + 53).Value = "NA" Then
            Cells(row, column + 53) = PROJECT_DATA.Range("D5").Value         'County
        End If
        
        If Cells(row, column + 54).Value = "NA" Then
            Cells(row, column + 54) = PROJECT_DATA.Range("AG5").Value         'District
        End If
        
        If Cells(row, column + 56).Value = "NA" Then
            Cells(row, column + 56) = PROJECT_DATA.Range("AH5").Value         'Bid Date
            Cells(row, column + 56).Select
            Selection.NumberFormat = "m/d/yyyy"
        End If
        
        If Cells(row, column + 57).Value = "NA" Or Cells(row, column + 57).Value = 0 Then
            Cells(row, column + 57) = PROJECT_DATA.Range("AK5").Value         'Project_fk
            Cells(row, column + 57).Select
            Selection.NumberFormat = "0"
        End If
        
        If Cells(row, column + 58).Value = "NA" Then
            Cells(row, column + 58) = PROJECT_DATA.Range("M5").Value         'Project No
        End If
        
    row = row + 1
    Loop
    
    AvgUnitCost.Range("BN2:BN999").Select
     Selection.NumberFormat = "m/d/yyyy h:mm"
    AvgUnitCost.Range("BP2:BP999").Select
     Selection.NumberFormat = "m/d/yyyy h:mm"
    AvgUnitCost.Range("BR2:BR99").Select
     Selection.NumberFormat = "0"
    AvgUnitCost.Select
    Cells.Select
    Selection.Replace What:="'", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceValues
    
    AvgUnitCost.Protect
    AvgUnitCost.Visible = False
    Database.Range("$A$1:$CS$999").AutoFilter
    Database.Range("A1").Select
    Database.Visible = False
    Database.Protect
End Sub

''''####################################################### EXTRACT TO SQL #############################################################'''



Sub Export_UnitCostToSQL()

'WARNING Message
If MsgBox("Are you ready to import the average unit cost data?", vbYesNo, "Save changes") = vbYes Then
    Dim conn As Object
    Dim rs As Object
    Dim strConn As String
    Dim strSQL As String
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range

    ' Set your worksheet and range
    Set ws = ThisWorkbook.Sheets("AvgUnitCost")
    Application.ScreenUpdating = False
    Sheets("AvgUnitCost").Visible = True
    Sheets("AvgUnitCost").Unprotect
    Sheets("AvgUnitCost").Select
    
   ' Set rng = DatabaseTempRegion ' Adjust the range as needed

    ' Create a connection string
    '' strConn = "driver={SQL server};server=roadwaydb6.osarc.local;database=AUC;uid=jhammons;pwd=Osarc123$"
    ' Old Connection string switching to Local database

strConn = _
    "Provider=SQLOLEDB;" & _
    "Server=roadwaydb6.osarc.local;" & _
    "Database=osarc;" & _
    "Trusted_Connection=Yes;"


    ' Create and open the connection
    Set conn = CreateObject("ADODB.Connection")
    conn.Open strConn

    ' Loop through each cell in the range and insert data into SQL
    Sheets("AvgUnitCost").Select
    Sheets("AvgUnitCost").Range("A2").Select
        
        Dim lastRow As Long, i As Long                          ' This sets up to only look at the last populated row and copy it over to SQL
        lastRow = Cells(Rows.Count, "A").End(xlUp).row
        For i = 2 To lastRow


        strSQL = _
        "INSERT INTO AverageUnitCost(Pay_Item,Item_Description,Quantity,Units,Unit_Price,Subtotal,FuelAdjustment_Subtotal,Fuel_Code,Field_1,Fuel_BaseMonth,Fuel_BaseYear,Fuel_CurrentMonth,Fuel_CurrentYear,Base_Gas,Base_Diesel,Current_Gas,Current_Diesel,AC_Percent,Binder_Type,Base_PG_64_22,Base_PG_67_22,Base_PG_76_22,Base_PG_82_22,Base_SS_1,Base_CRS_2,Base_CRS_2P,Base_EA_1,Base_CSS_1UN,Base_CSS_1,Current_PG_64_22,Current_PG_67_22,Current_PG_76_22,Current_PG_82_22,Current_SS_1,Current_CRS_2,Current_CRS_2P,Current_EA_1,Current_CSS_1UN,Current_CSS_1,FuelAdjustment_Gas,FuelAdjustment_Diesel,FuelAdjustment_Binder,FuelAdjustment_Total,Past_FuelAdjustmentSubtotal,Field_2,Type,Participating,SORT_order,Current_InvoiceNo,Past_InvoiceNo,ChangeOrder_No,Latitude,Longitude,County,District,Site,Bid_Date,Project_fk,Project_No,Past_AccumulativeQuantity,Past_AccumulativeCost,Current_PeriodQuantity,Current_PeriodCost,Current_TotalQuantity,Current_TotalCost,Modified," & _
"Modified_By,Created,Created_By,Specification,Field_3,Field_4,Field_5,Field_6,Field_7,Field_8,Field_9,Field_10,Field_11,Field_12,Field_13,Field_14,Field_15,Field_16,Field_17, Field_18,Field_19,Field_20,Field_21,Field_22,Field_23,Field_24,Field_25,Field_26,Field_27,Field_28,Field_29)" & _
        "values('" & ws.Cells(i, 1) & "','" & ws.Cells(i, 2) & "','" & ws.Cells(i, 3) & "','" & ws.Cells(i, 4) & "','" & ws.Cells(i, 5) & "','" & ws.Cells(i, 6) _
        & "','" & ws.Cells(i, 7) & "','" & ws.Cells(i, 8) & "','" & ws.Cells(i, 9) & "','" & ws.Cells(i, 10) & "','" & ws.Cells(i, 11) & "','" & ws.Cells(i, 12) _
        & "','" & ws.Cells(i, 13) & "','" & ws.Cells(i, 14) & "','" & ws.Cells(i, 15) & "','" & ws.Cells(i, 16) & "','" & ws.Cells(i, 17) & "','" & ws.Cells(i, 18) _
        & "','" & ws.Cells(i, 19) & "','" & ws.Cells(i, 20) & "','" & ws.Cells(i, 21) & "','" & ws.Cells(i, 22) & "','" & ws.Cells(i, 23) & "','" & ws.Cells(i, 24) _
        & "','" & ws.Cells(i, 25) & "','" & ws.Cells(i, 26) & "','" & ws.Cells(i, 27) & "','" & ws.Cells(i, 28) & "','" & ws.Cells(i, 29) & "','" & ws.Cells(i, 30) _
        & "','" & ws.Cells(i, 31) & "','" & ws.Cells(i, 32) & "','" & ws.Cells(i, 33) & "','" & ws.Cells(i, 34) & "','" & ws.Cells(i, 35) & "','" & ws.Cells(i, 36) _
        & "','" & ws.Cells(i, 37) & "','" & ws.Cells(i, 38) & "','" & ws.Cells(i, 39) & "','" & ws.Cells(i, 40) & "','" & ws.Cells(i, 41) & "','" & ws.Cells(i, 42) _
        & "','" & ws.Cells(i, 43) & "','" & ws.Cells(i, 44) & "','" & ws.Cells(i, 45) & "','" & ws.Cells(i, 46) & "','" & ws.Cells(i, 47) & "','" & ws.Cells(i, 48) _
        & "','" & ws.Cells(i, 49) & "','" & ws.Cells(i, 50) & "','" & ws.Cells(i, 51) & "','" & ws.Cells(i, 52) & "','" & ws.Cells(i, 53) & "','" & ws.Cells(i, 54) _
        & "','" & ws.Cells(i, 55) & "','" & ws.Cells(i, 56) & "','" & ws.Cells(i, 57) & "','" & ws.Cells(i, 58) & "','" & ws.Cells(i, 59) & "','" & ws.Cells(i, 60) _
        & "','" & ws.Cells(i, 61) & "','" & ws.Cells(i, 62) & "','" & ws.Cells(i, 63) & "','" & ws.Cells(i, 64) & "','" & ws.Cells(i, 65) & "','" & ws.Cells(i, 66) _
        & "','" & ws.Cells(i, 67) & "','" & ws.Cells(i, 68) & "','" & ws.Cells(i, 69) & "','" & ws.Cells(i, 70) & "','" & ws.Cells(i, 71) & "','" & ws.Cells(i, 72) _
        & "','" & ws.Cells(i, 73) & "','" & ws.Cells(i, 74) & "','" & ws.Cells(i, 75) & "','" & ws.Cells(i, 76) & "','" & ws.Cells(i, 77) & "','" & ws.Cells(i, 78) & "','" & ws.Cells(i, 79) _
        & "','" & ws.Cells(i, 80) & "','" & ws.Cells(i, 81) & "','" & ws.Cells(i, 82) & "','" & ws.Cells(i, 83) & "','" & ws.Cells(i, 84) & "','" & ws.Cells(i, 85) _
        & "','" & ws.Cells(i, 86) & "','" & ws.Cells(i, 87) & "','" & ws.Cells(i, 88) & "','" & ws.Cells(i, 89) & "','" & ws.Cells(i, 90) & "','" & ws.Cells(i, 91) _
        & "','" & ws.Cells(i, 92) & "','" & ws.Cells(i, 93) & "','" & ws.Cells(i, 94) & "','" & ws.Cells(i, 95) & "','" & ws.Cells(i, 96) & "','" & ws.Cells(i, 97) _
        & "')"
        
        conn.Execute strSQL
    Next

    ' Clean up
    conn.Close
    Set conn = Nothing
    Set rs = Nothing

    Sheets("AvgUnitCost").Protect
    Sheets("AvgUnitCost").Visible = False
    Application.ScreenUpdating = True
    MsgBox ("You have successfully uploaded " & lastRow - 1 & " records to the database. Ensure this matches the Excel file")
    
    End If
    
    
End Sub

