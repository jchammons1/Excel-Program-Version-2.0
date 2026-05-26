VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_DASHBOARD 
   Caption         =   "DASHBOARD"
   ClientHeight    =   13920
   ClientLeft      =   90
   ClientTop       =   390
   ClientWidth     =   22350
   OleObjectBlob   =   "UserForm_DASHBOARD.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_DASHBOARD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


 Sub BidDate_Button_Click()
    UserForm_BidDatePicker.Show
    txt_BidDate.Value = PROJECT_DATA.Range("AH2").Value
End Sub

'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'*********************************** This section contains the Command Buttons on the forms       **************************************>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################

Sub Button_BaseFuelAdjustmentData_Click()
Application.ScreenUpdating = False
    CAD.Visible = True
    CAD.Select
    CAD.Unprotect Password:="roadway123"
    UserForm_BaseFuelData.Show
End Sub

 Sub Button_BidTab_Click()
    UserForm_DASHBOARD.lblmessage3.Caption = "Processing, please wait....."
    DoEvents
    Call DownloadWAVFile
    Application.ScreenUpdating = False
    BIDTAB.Select
        Call Estimates
    BIDTAB.Select
    BIDTAB.Range("B4").Select
        Call BidTab_Step1
        Call BidTab_Step2
        Call SetPrintArea_Estimate
        Call Flag_618B_803B
        Call Flag_618B_803B_Bidders
        Call ReferenceNo
        Call WEB_BIDLIST
    'Clear processing message
    UserForm_DASHBOARD.lblmessage3.Caption = ""
    DoEvents
    PAYITEMTYPE.Visible = False
    Database.Visible = False
    BIDTAB.Select
    BIDTAB.Protect
    BIDTAB.Range("B4").Select
    Application.ScreenUpdating = True
    Unload Me
End Sub



Private Sub Button_ConstructionType_Click()
    Application.ScreenUpdating = True
    DATA_VALIDATION.Visible = True
    DATA_VALIDATION.Unprotect
    DATA_VALIDATION.Range("Z2") = "Yes"     ' change the drop down selection to supplemental pay items
    Call UpdateTypeComboBox
    DATA_VALIDATION.Protect
    DATA_VALIDATION.Visible = False
    Application.ScreenUpdating = True
End Sub

 Sub Button_ContractorEstimate_Click()
'Exporting the Awarded bidder to the top of the sheet
   Application.ScreenUpdating = False
    CONTRACTOR_ESTIMATE.Select
    CONTRACTOR_ESTIMATE.Unprotect
    Range("A20:AB999").Clear
    Range("B20") = BIDTAB.Range("DA18:DB18").Value
    Range("B20").Select
        With Selection
            Selection.HorizontalAlignment = xlCenter
            Selection.Font.Size = 12
            Selection.Font.Bold = True
            .ShrinkToFit = True
        End With
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 13434879
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        
        Call Estimates
        
    CONTRACTOR_ESTIMATE.Select
    CONTRACTOR_ESTIMATE.Range("B5").Select
        Call SetPrintArea_Estimate
        Call Flag_618B_803B
    CONTRACTOR_ESTIMATE.Protect

    Call AverageUnitCost
    
    Application.ScreenUpdating = True
    CONTRACTOR_ESTIMATE.Select
    Unload Me
End Sub

Private Sub Button_CreateSuppCorrection_Click()
    UserForm_EditSuppListBox.Show
End Sub



 Sub Button_DeletePayItem_Click()
' Delete pay item from the Database sheet
    If MsgBox("Do you want to delete this pay item?", vbYesNo, "Save Changes") = vbYes Then
        Database.Unprotect
        Database.Visible = True
            Call DeleteRow_Database(ListBox_Database.ListIndex)
        Database.Protect
        Database.Visible = False
    End If
End Sub

Sub Button_EngineersEstimate_Click()
    Application.ScreenUpdating = False
    ENGINEERS_ESTIMATE.Select
    ENGINEERS_ESTIMATE.Unprotect
        Call Estimates
    ENGINEERS_ESTIMATE.Select
    ENGINEERS_ESTIMATE.Range("B5").Select
        Call SetPrintArea_Estimate
        Call Flag_618B_803B
    ENGINEERS_ESTIMATE.Protect
    Application.ScreenUpdating = True
End Sub


 Sub Button_Form900_Click()
    Application.ScreenUpdating = False
    Database.Unprotect
    Database.Visible = True
        Call SortPayItems
    Form900.Select
    Form900.Unprotect
        Call Form_900
    Database.Protect
    Database.Visible = False
    Form900.Select
    Form900.Range("B10").Select
    Form900.Protect
    Application.ScreenUpdating = True
End Sub

Private Sub Button_Form900_Alternate_Click()
    Application.ScreenUpdating = False
    Database.Unprotect
    Database.Visible = True
        Call SortPayItems
    Form900.Select
    Form900.Unprotect
        Call Form_900Alternate
    Database.Protect
    Database.Visible = False
    Form900.Select
    Form900.Range("B10").Select
    Form900.Protect
    Application.ScreenUpdating = True
End Sub

 Sub Button_Form902_Click()
    Application.ScreenUpdating = False
    Database.Unprotect
    Database.Visible = True
        Call SortPayItems
    Form902.Select
    Form902.Unprotect Password:="roadway123"
    Range("B4:F7").Clear
    Range("A13:CZ999").Clear
    Range("G2").Clear
    Range("E3:G3").Clear
        Call Form_902
    Database.Protect
    Database.Visible = False
    Form902.Select
    Form902.Range("A13").Select
    Form902.Protect Password:="roadway123"
    Application.ScreenUpdating = True
End Sub

 Sub Button_GreenProjectCover_Click()
    Application.ScreenUpdating = False
    GreenCover.Select
    GreenCover.Unprotect
    Range("P8:T8").Clear
    Range("D15:G15").Clear
    Range("Q18:V18").Clear
    Range("P19:V19").Clear
    Range("O20:Q20").Clear
    Range("P21:V21").Clear
    Range("N22:V22").Clear
    Range("P23").Clear
    Range("P24").Clear
    Range("U24").Clear
        Call Form_GreenCover
    GreenCover.Select
    GreenCover.Protect
    Application.ScreenUpdating = True
    Range("P8").Select
End Sub

Private Sub Button_GreenProjectCoverNEW_Click()
    Application.ScreenUpdating = False
    GreenCoverNEW.Select
    GreenCoverNEW.Unprotect
    Range("E16:I16").Clear
    Range("E59:H59").Clear
    Range("F26:K26").Clear
    Range("E27:K27").Clear
    Range("D28:F28").Clear
    Range("D29:K29").Clear
    
    Range("C30:K30").Clear
    Range("E31").Clear
    Range("E32").Clear
    Range("K32").Clear
        Call Form_GreenCoverNEW
    GreenCoverNEW.Select
    GreenCoverNEW.Protect
    Application.ScreenUpdating = True
    Range("E16").Select
End Sub

 Sub Button_LettingResults_Click()
    Call Form_LettingResults
End Sub

Private Sub Button_Materials_Click()
 Call MaterialsReport
 Unload Me
End Sub

Private Sub Button_MDOT2017_Click()
' 2017 MDOT PAY ITEMS
    Application.ScreenUpdating = False
    DATA_VALIDATION.Visible = True
    DATA_VALIDATION.Unprotect
    DATA_VALIDATION.Range("Y2") = "MDOT"
    DATA_VALIDATION.Protect
    DATA_VALIDATION.Visible = False
    txt_SpecificationYearSelected = DATA_VALIDATION.Range("Y2")
        If DATA_VALIDATION.Range("Y2") = "MDOT" Then
            ' Filter RowSource when checkbox is checked
            cmb_PayItemSearch.RowSource = "PayItemListMDOT!A2:A9999" ' Adjust the range as needed
        End If
    Application.ScreenUpdating = True
End Sub

Private Sub Button_MusicOff_Click()
    Application.ScreenUpdating = True
    DATA_VALIDATION.Visible = True
    DATA_VALIDATION.Unprotect
    DATA_VALIDATION.Range("AA2") = "Off"
    DATA_VALIDATION.Protect
    DATA_VALIDATION.Visible = False
    Application.ScreenUpdating = True
End Sub

Private Sub Button_MusicOn_Click()
    Application.ScreenUpdating = True
    DATA_VALIDATION.Visible = True
    DATA_VALIDATION.Unprotect
    DATA_VALIDATION.Range("AA2") = "On"
    DATA_VALIDATION.Protect
    DATA_VALIDATION.Visible = False
    Application.ScreenUpdating = True
End Sub

 Sub Button_OSARCEstimate_Click()
    Application.ScreenUpdating = False
    OSARC_ESTIMATE.Select
    OSARC_ESTIMATE.Protect
        Call Estimates
    OSARC_ESTIMATE.Select
    OSARC_ESTIMATE.Range("B5").Select
        Call SetPrintArea_Estimate
        Call Flag_618B_803B
    OSARC_ESTIMATE.Protect
    Application.ScreenUpdating = True
    Unload Me
End Sub

Private Sub Button_PhaseYes_Click()
    Application.ScreenUpdating = False
    DATA_VALIDATION.Visible = True
    DATA_VALIDATION.Unprotect
    DATA_VALIDATION.Range("Z2") = "Yes"
    ''Button_PhaseYes.Value = DATA_VALIDATION.Range("Z2")
    DATA_VALIDATION.Protect
    DATA_VALIDATION.Visible = False
    Application.ScreenUpdating = True
        If DATA_VALIDATION.Range("Z2") = "Yes" Then
            ' Filter RowSource when checkbox is checked
            cmb_Type.RowSource = "PAYITEMTYPE_SUPPLEMENTAL!A2:A51" ' Adjust the range as needed
        End If
End Sub

Private Sub Button_PreConstructionTypes_Click()
    Application.ScreenUpdating = True
    DATA_VALIDATION.Visible = True
    DATA_VALIDATION.Unprotect
    DATA_VALIDATION.Range("Z2") = "No"     ' change the drop down selection to supplemental pay items
    Call UpdateTypeComboBox
    DATA_VALIDATION.Protect
    DATA_VALIDATION.Visible = False
    Application.ScreenUpdating = True
End Sub

Sub Button_Process_Click()
   
' Retrieve values from Pay Item List sheet and inputs the values on the form
Dim PayItemNo As Variant
Dim PayItemDescription As Variant
Dim Unit As Variant
Dim FuelCode As Variant
Dim Participating As Variant
Dim SortOrder As Variant
Dim PayItemListSelected As String         ' Added this variable to lookup based on the pay item list selected by the user.
                                        ' Data_Validation Sheet Range("Y2") = selected pay item list


If DATA_VALIDATION.Range("Y2") = "2004" Then
    ' Filter the range for pay item list used below
            PayItemListSelected = "PayItemList"
        ElseIf DATA_VALIDATION.Range("Y2") = "2024" Then
            PayItemListSelected = "PayItemList2024"
        ElseIf DATA_VALIDATION.Range("Y2") = "MDOT" Then
            PayItemListSelected = "PayItemListMDOT"
        Else
End If
        
PayItemNo_Lookup = Application.VLookup(cmb_PayItemSearch.Value, Sheets(PayItemListSelected).Range("A1:I9999"), 2, False)
PayItemDescription_Lookup = Application.VLookup(cmb_PayItemSearch.Value, Sheets(PayItemListSelected).Range("A1:I9999"), 3, False)
Unit_Lookup = Application.VLookup(cmb_PayItemSearch.Value, Sheets(PayItemListSelected).Range("A1:I9999"), 4, False)
FuelCode_Lookup = Application.VLookup(cmb_PayItemSearch.Value, Sheets(PayItemListSelected).Range("A1:I9999"), 5, False)
Participating_Lookup = Application.VLookup(cmb_PayItemSearch.Value, Sheets(PayItemListSelected).Range("A1:I9999"), 6, False)
Type_Lookup = Application.VLookup(cmb_PayItemSearch.Value, Sheets(PayItemListSelected).Range("A1:I9999"), 7, False)
SpecYear_Lookup = Application.VLookup(cmb_PayItemSearch.Value, Sheets(PayItemListSelected).Range("A1:I9999"), 8, False)
SortOrder_Lookup = Application.VLookup(cmb_PayItemSearch.Value, Sheets(PayItemListSelected).Range("A1:I9999"), 9, False)

    If IsError(PayItemNo_Lookup) Then
            txt_PayItemNo.Value = cmb_PayItemSearch.Value
            Else
            txt_PayItemNo.Value = Application.VLookup(cmb_PayItemSearch.Value, Sheets(PayItemListSelected).Range("A1:I9999"), 2, False)
    End If
    If IsError(PayItemDescription_Lookup) Then
            txt_PayItemDescription.Value = "Manually complete all the fields to add this new pay item.  The pay item is not a selection"
            Else
            txt_PayItemDescription.Value = Application.VLookup(cmb_PayItemSearch.Value, Sheets(PayItemListSelected).Range("A1:I9999"), 3, False)
    End If
    If IsError(Unit_Lookup) Then
            txt_Unit.Value = ""
            Else
            txt_Unit.Value = Application.VLookup(cmb_PayItemSearch.Value, Sheets(PayItemListSelected).Range("A1:I9999"), 4, False)
    End If
    If IsError(FuelCode_Lookup) Then
            cmb_FuelCode.Value = "NA"
            Else
            cmb_FuelCode.Value = Application.VLookup(cmb_PayItemSearch.Value, Sheets(PayItemListSelected).Range("A1:I9999"), 5, False)
    End If
    If IsError(Participating_Lookup) Then
            cmb_Participating.Value = "Yes"
            Else
            cmb_Participating.Value = Application.VLookup(cmb_PayItemSearch.Value, Sheets(PayItemListSelected).Range("A1:I9999"), 6, False)
    End If
    If IsError(Type_Lookup) Then
            cmb_Type.Value = ""
            Else
            cmb_Type.Value = Application.VLookup(cmb_PayItemSearch.Value, Sheets(PayItemListSelected).Range("A1:I9999"), 7, False)
    End If
        If IsError(SpecYear_Lookup) Then
            txt_SpecYear.Value = "1900"
            Else
            txt_SpecYear.Value = Application.VLookup(cmb_PayItemSearch.Value, Sheets(PayItemListSelected).Range("A1:I9999"), 8, False)
    End If
    If IsError(SortOrder_Lookup) Then
            txt_SortOrder.Value = "9999"
            Else
            txt_SortOrder.Value = Application.VLookup(cmb_PayItemSearch.Value, Sheets(PayItemListSelected).Range("A1:I9999"), 9, False)
    End If

' Avoids a debug error by checking to see if the user entered data first
    If txt_Quantity.Value = "" Or txt_UnitPrice.Value = "" Then
        MsgBox "The quantity and/or unit price is blank.  A value is required.  No data is saved. Please try again."
    Else
        With Me
            txt_Subtotal.Value = txt_Quantity.Value * txt_UnitPrice.Value
        End With
    End If
End Sub

 Sub Button_AddPayItem_Click()
' Declare a variable range to lookup the next blank row in the Database
    Database.Unprotect
    Database.Visible = True
    
    Dim NextRow_Database As Range
    
' Check form for blank values
If Me.txt_PayItemNo.Value = "" Or Me.txt_Quantity.Value = "" Or Me.txt_UnitPrice.Value = "" Or Me.txt_Unit.Value = "" _
    Or Me.txt_Subtotal.Value = "" Or Me.cmb_FuelCode.Value = "" Or Me.cmb_Type.Value = "" Or Me.cmb_Participating.Value = "" _
    Or Me.txt_SortOrder.Value = "" Then
    MsgBox "Please ensure all fields are populated.  The data has not been saved"
    Exit Sub
End If

'Find the next blank row in the Database sheet
    Set NextRow_Database = Database.Cells(Rows.Count, 2).End(xlUp).Offset(1, -1)
'Send the data from the form to the Database sheet
'All other cells set to "NA" because in order to operarate properly, each cell needs a value
    With NextRow_Database
        .Value = Me.txt_PayItemNo.Value
        .Offset(0, 1).Value = Me.txt_PayItemDescription.Value
        .Offset(0, 2).Value = Format(Me.txt_Quantity.Value, "#,##0.000")
        .Offset(0, 3).Value = Me.txt_Unit.Value
        .Offset(0, 4).Value = Format(Me.txt_UnitPrice.Value, "Currency")
        .Offset(0, 5).Value = Format(Me.txt_Subtotal.Value, "Currency")
        .Offset(0, 6).Value = 0
        .Offset(0, 7).Value = Me.cmb_FuelCode.Value
        .Offset(0, 8).Value = "NA"
        .Offset(0, 9).Value = "NA"
        .Offset(0, 10).Value = 0
        .Offset(0, 11).Value = "NA"
        .Offset(0, 12).Value = 0
        .Offset(0, 13).Value = 0
        .Offset(0, 14).Value = 0
        .Offset(0, 15).Value = 0
        .Offset(0, 16).Value = 0
        .Offset(0, 17).Value = 0
        .Offset(0, 18).Value = "NA"
        .Offset(0, 19).Value = 0
        .Offset(0, 20).Value = 0                'Fuel adjustment data entry values = 0
        .Offset(0, 21).Value = 0
        .Offset(0, 22).Value = 0
        .Offset(0, 23).Value = 0
        .Offset(0, 24).Value = 0
        .Offset(0, 25).Value = 0
        .Offset(0, 26).Value = 0
        .Offset(0, 27).Value = 0
        .Offset(0, 28).Value = 0
        .Offset(0, 29).Value = 0
        .Offset(0, 30).Value = 0
        .Offset(0, 31).Value = 0
        .Offset(0, 32).Value = 0
        .Offset(0, 33).Value = 0
        .Offset(0, 34).Value = 0
        .Offset(0, 35).Value = 0
        .Offset(0, 36).Value = 0
        .Offset(0, 37).Value = 0
        .Offset(0, 38).Value = 0
        .Offset(0, 39).Value = 0
        .Offset(0, 40).Value = 0
        .Offset(0, 41).Value = 0
        .Offset(0, 42).Value = 0
        .Offset(0, 43).Value = 0
        .Offset(0, 44).Value = 0
        .Offset(0, 45).Value = Me.cmb_Type.Value
        .Offset(0, 46).Value = Me.cmb_Participating.Value
        .Offset(0, 47).Value = Me.txt_SortOrder.Value
        .Offset(0, 48).Value = "NA"
        .Offset(0, 49).Value = "NA"
        .Offset(0, 50).Value = "NA"
        .Offset(0, 51).Value = 0
        .Offset(0, 52).Value = 0
        .Offset(0, 53).Value = "NA"
        .Offset(0, 54).Value = "NA"
        .Offset(0, 55).Value = "NA"
        .Offset(0, 56).Value = "NA"
        .Offset(0, 57).Value = "NA"
        .Offset(0, 58).Value = "NA"
        .Offset(0, 59).Value = 0
        .Offset(0, 60).Value = 0
        .Offset(0, 61).Value = 0
        .Offset(0, 62).Value = 0
        .Offset(0, 63).Value = 0
        .Offset(0, 64).Value = 0
        .Offset(0, 65).Value = Now                                                             ' Modified Date Time
        .Offset(0, 66).Value = ActiveWorkbook.BuiltinDocumentProperties("Last Author")         ' Modified by
        .Offset(0, 67).Value = Now                                                             ' Created Date Time
        .Offset(0, 68).Value = ActiveWorkbook.BuiltinDocumentProperties("Last Author")         ' Created by
        .Offset(0, 69).Value = Me.txt_SpecYear.Value                                           ' Specification Year
        .Offset(0, 70).Value = "NA"
        .Offset(0, 71).Value = "NA"
        .Offset(0, 72).Value = "NA"
        .Offset(0, 73).Value = "NA"
        .Offset(0, 74).Value = "NA"
        .Offset(0, 75).Value = "NA"
        .Offset(0, 76).Value = "NA"
        .Offset(0, 77).Value = "NA"
        .Offset(0, 78).Value = "NA"
        .Offset(0, 79).Value = "NA"
        .Offset(0, 80).Value = "NA"
        .Offset(0, 81).Value = "NA"
        .Offset(0, 82).Value = "NA"
        .Offset(0, 83).Value = "NA"
        .Offset(0, 84).Value = "NA"
        .Offset(0, 85).Value = "NA"
        .Offset(0, 86).Value = "NA"
        .Offset(0, 87).Value = "NA"
        .Offset(0, 88).Value = "NA"
        .Offset(0, 89).Value = "NA"
        .Offset(0, 90).Value = "NA"
        .Offset(0, 91).Value = "NA"
        .Offset(0, 92).Value = "NA"
        .Offset(0, 93).Value = "NA"
        .Offset(0, 94).Value = "NA"
        .Offset(0, 95).Value = "NA"
        .Offset(0, 96).Value = "NA"
        
    End With
'Update the list box after clicking the Add Pay Item Button
    Call AddPayItemToListBox_Database
'Clear the values for the next user pay item entry'
    Call cmb_PayItemSearch_Change
    cmb_PayItemSearch.Value = ""
'After the pay item is added, then set the cursor back to the Pay Item search field
    UserForm_DASHBOARD.cmb_PayItemSearch.SetFocus
    Database.Protect
    Database.Visible = False
End Sub

Sub AddPayItemToListBox_Database()
' Get the data range from the Database sheet
  Database.Unprotect
   ' Database.Visible = True
    
    Dim Database_Range As Range
    Set Database_Range = GetDatabase_Range
' Link the data to the list box.  Setup list box headings and columns widths
    With ListBox_Database
        .RowSource = Database_Range.Address(external:=True)
        .ColumnCount = Database_Range.Columns.Count
        .ColumnWidths = "85;500;60,60;70;85;60;60;85;50;60;60"
        .ColumnHeads = True
        .ListIndex = 0
    End With
Database.Protect
End Sub


'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'*********************************** This section contains Buttons on the Dashboard to open other forms *********************************>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################

 Sub Button_CloseProjectData_Click()
    Unload Me
End Sub

 Sub Button_CreateSupplementalTypeForm_Click()
    UserForm_SupplementalType.Show
End Sub

 Sub Button_CreateTypeForm_Click()
    UserForm_EditTypeListBox.Show
End Sub

 Sub Button_EditPayItemForm_Click()
'Modify a pay item and open Edit Pay Item User Form
    Database.Unprotect
    Database.Visible = True
        Call EditRow_Database
    Database.Protect
    Database.Visible = False
End Sub

Sub EditRow_Database()
'Shows the edit form for the pay item selected in the list box
    Dim frm As New UserForm_EditPayItem
    frm.currentRow = ListBox_Database.ListIndex
    frm.Show vbModal
End Sub

 Sub Button_EstimateDate_Click()
    UserForm_EstimateDate.Show
    txt_EstimateDate.Value = PROJECT_DATA.Range("N2").Value
    PROJECT_DATA.Protect
    PROJECT_DATA.Visible = False
End Sub

 Sub Button_OpenDocumentHeading_Click()
    UserForm_DocumentHeading.Show
End Sub

Private Sub Button_Specification2004_Click()
    Application.ScreenUpdating = False
    DATA_VALIDATION.Visible = True
    DATA_VALIDATION.Unprotect
    DATA_VALIDATION.Range("Y2") = 2004
    txt_SpecificationYearSelected = DATA_VALIDATION.Range("Y2")
    DATA_VALIDATION.Protect
    DATA_VALIDATION.Visible = False
    Application.ScreenUpdating = True
        If DATA_VALIDATION.Range("Y2") = 2004 Then
            ' Filter RowSource when checkbox is checked
            cmb_PayItemSearch.RowSource = "PayItemList!A2:A500" ' Adjust the range as needed
        End If
End Sub

Private Sub Button_Specification2024_Click()
    Application.ScreenUpdating = False
    DATA_VALIDATION.Visible = True
    DATA_VALIDATION.Unprotect
    DATA_VALIDATION.Range("Y2") = 2024
    DATA_VALIDATION.Protect
    DATA_VALIDATION.Visible = False
    txt_SpecificationYearSelected = DATA_VALIDATION.Range("Y2")
        If DATA_VALIDATION.Range("Y2") = 2024 Then
            ' Filter RowSource when checkbox is checked
            cmb_PayItemSearch.RowSource = "PayItemList2024!A2:A500" ' Adjust the range as needed
        End If
    Application.ScreenUpdating = True
End Sub

Sub Button_SQL_Download_Click()
    Call Export_UnitCostToSQL
End Sub

Private Sub Button_UpdateUnitCost_Click()
    Call DatabaseUnitCostUpdate
End Sub

Private Sub Button_PrintCAD_Click()
    CAD.Visible = True
    CAD.Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    CAD.Visible = False
    START.Select
End Sub




Private Sub ButtonDeleteMonthlyFuelAdjustment_Click()

End Sub

Private Sub Buttong_AddCADSignature_Click()
    UserForm_CAD_Signature.Show
End Sub

'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'*********************************** This section resets the Pay Item Form when a new pay item is selected*******************************>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################

 Sub cmb_PayItemSearch_Change()
    With Me
        txt_PayItemNo.Value = ""
        txt_Subtotal.Value = ""
        txt_Unit.Value = ""
        txt_Quantity.Value = ""
        txt_UnitPrice.Value = ""
        cmb_Participating.Value = ""
        cmb_Type.Value = ""
        cmb_FuelCode.Value = ""
        txt_SortOrder.Value = ""
        txt_PayItemDescription.Value = ""
        txt_SpecYear.Value = ""
    End With
End Sub





 Sub Create_CAD_Click()
    ' Set the message
    UserForm_DASHBOARD.lblMessage.Caption = "Processing, please wait..."
    DoEvents
        Call DownloadWAVFile
    Application.ScreenUpdating = False
    CAD.Visible = True
    CAD.Unprotect Password:="roadway123"
       ' Call Open_MonthlyFuelData
        Call Update_CAD
    ' Update message
            UserForm_DASHBOARD.lblMessage.Caption = "Almost done processing..."
                    DoEvents
    CAD.Visible = True
    CAD.Unprotect Password:="roadway123"
    CAD.Select
        Call Create_NewCAD
    CAD.Protect Password:="roadway123"
    CAD.Visible = False
    Database.Protect
    Database.Visible = False
    DATA_VALIDATION.Visible = False
    PAYITEMTYPE.Visible = False
    PAYITEMTYPE_SUPPLEMENTAL.Visible = False
    ActiveSheet.Protect Password:="roadway123"
    Application.ScreenUpdating = True
    ActiveSheet.Range("G7").Select
    ' Clear the message
            UserForm_DASHBOARD.lblMessage.Caption = ""
            DoEvents
    MsgBox "CAD Invoice No. " & ActiveSheet.Name & " is ready for processing"
    Unload Me
End Sub

 Sub Create_EngineerPayment_Click()
    Call EngineerPayment
    Unload Me
End Sub

 Sub CreateModify_CAD_Click()
        ' Set the message
    UserForm_DASHBOARD.lblmessage2.Caption = "Processing, please wait..."
            DoEvents
    Call Update_CAD
    UserForm_DASHBOARD.lblmessage2.Caption = ""
            DoEvents
    MsgBox "CAD Template updated. Review the CAD template sheet to ensure the pay item updates are correct"
End Sub
Sub Update_CAD()
    Application.ScreenUpdating = False
    CAD.Visible = True
    CAD.Unprotect Password:="roadway123"
    CAD.Select
        Call Organize_CAD
    CAD.Range("G7").Select
    CAD.Protect Password:="roadway123"
    CAD.Visible = False
    Application.ScreenUpdating = True
    'Unload Me
End Sub


Sub Delete900_Click()
    Call Delete_900
End Sub

Sub Delete902_Click()
    Call Delete_902
End Sub

Private Sub DeleteBidTab_Click()
    BIDTAB.Select
    BIDTAB.Unprotect
    BIDTAB.Range("DA21:DZ9999").Clear
    BIDTAB.Range("B4:F4").ClearContents
    Delete_Estimates
End Sub

Private Sub DeleteCAD002_Click()
    Call Delete_CAD
End Sub

Private Sub DeleteContractorsEstimate_Click()
    CONTRACTOR_ESTIMATE.Select
    CONTRACTOR_ESTIMATE.Unprotect
    CONTRACTOR_ESTIMATE.Range("B20").ClearContents
    Delete_Estimates
End Sub

Private Sub DeleteDatabasePayItems_Click()
    Call Delete_PayItems
End Sub

Private Sub DeleteEngineersEstimate_Click()
    ENGINEERS_ESTIMATE.Select
    Delete_Estimates
End Sub

Private Sub DeleteEnvPayInv_Click()
   Call Delete_EngineersPayment
End Sub

Private Sub DeleteGreenCover_Click()
    Call Delete_GreenCover
End Sub

Private Sub DeleteLettingResults_Click()
    Call Delete_LettingResults
End Sub

Private Sub DeleteOSARC_Estimate_Click()
    OSARC_ESTIMATE.Select
    Delete_Estimates
End Sub

Private Sub DeletePayItemTypes_Click()
    Call Delete_PayItemTypes
End Sub

Private Sub DeletePayItemTypesSupp_Click()
 Call Delete_PayItemTypesSupplemental
End Sub

Private Sub DeleteProjectData_Click()
    Call Delete_ProjectData
End Sub

Private Sub DeleteWebBidList_Click()
    Call Delete_WebBidList
End Sub




Private Sub DeleteMaterialsReport_Click()
    Call Delete_MaterialsReport
End Sub



Private Sub Hide_DefaultSheets_Click()
    BIDTAB.Visible = False
    AvgUnitCost.Visible = False
    CONTRACTOR_ESTIMATE.Visible = False
    ENGINEERS_ESTIMATE.Visible = False
    GreenCover.Visible = False
    OSARC_ESTIMATE.Visible = False
    WEBBIDLIST.Visible = False
    Material_Report.Visible = False
    Form900.Visible = False
    Form902.Visible = False
    LettingResults.Visible = False
End Sub

 Sub Save_CADdata_Click()
'Invoice Tab for the CAD template data
Application.ScreenUpdating = False
    CAD.Unprotect Password:="roadway123"
    CAD.Visible = True
    CAD.Select
        CAD.Range("B3").Value = txt_VendorNo.Value
        CAD.Range("B5").Value = txt_ContractorNameAddress.Value
        CAD.Range("G6").Value = txt_Surety.Value
        CAD.Range("B6").Value = txt_Email.Value
        CAD.Range("B13").Value = txt_Retainage.Value
    PROJECT_DATA.Unprotect
    PROJECT_DATA.Visible = True
        PROJECT_DATA.Range("J2").Value = txt_WorkingDaysUpdate.Value
    PROJECT_DATA.Protect
    PROJECT_DATA.Visible = False
    
    DATA_VALIDATION.Visible = True
    DATA_VALIDATION.Unprotect
    DATA_VALIDATION.Range("Z2") = "Yes"     ' change the drop down selection to supplemental pay items
    DATA_VALIDATION.Protect
    DATA_VALIDATION.Visible = False

    CAD.Protect Password:="roadway123"
    CAD.Select
    CAD.Visible = False
    
    Application.ScreenUpdating = True
    START.Select
End Sub


 Sub Save_EngineerPaymentData_Click()
    Application.ScreenUpdating = False
    EngPayInv.Unprotect
    EngPayInv.Visible = True
    EngPayInv.Select
    PROJECT_DATA.Unprotect
    PROJECT_DATA.Visible = True

        EngPayInv.Range("B5").Value = txt_CEVendorNo.Value
        EngPayInv.Range("D5").Value = txt_CEVendorAddress.Value
        EngPayInv.Range("D11").Value = txt_ContractAmount.Value
        EngPayInv.Range("C18").Value = txt_PaidTo.Value
        EngPayInv.Range("B20").Value = txt_AgreementDate.Value
        EngPayInv.Range("B21").Value = txt_ApprovedSAEngineer.Value
        PROJECT_DATA.Range("AO2").Value = cmb_EngineeringFunding.Value
    EngPayInv.Protect
    EngPayInv.Select
    EngPayInv.Visible = False
    
    EngPayInvFedFund.Unprotect
    EngPayInvFedFund.Visible = True
    EngPayInvFedFund.Select
        EngPayInvFedFund.Range("B5").Value = txt_CEVendorNo.Value
        EngPayInvFedFund.Range("D5").Value = txt_CEVendorAddress.Value
        EngPayInvFedFund.Range("D11").Value = txt_ContractAmount.Value
        EngPayInvFedFund.Range("C18").Value = txt_PaidTo.Value
        EngPayInvFedFund.Range("B20").Value = txt_AgreementDate.Value
        EngPayInvFedFund.Range("B21").Value = txt_ApprovedSAEngineer.Value
        PROJECT_DATA.Range("AO2").Value = cmb_EngineeringFunding.Value
    EngPayInvFedFund.Protect
    EngPayInvFedFund.Select
    EngPayInvFedFund.Visible = False
    PROJECT_DATA.Protect
    PROJECT_DATA.Visible = False
    Application.ScreenUpdating = True
    START.Select
    
End Sub

 Sub Save_MonthlyFuelAdjustment_Click()
    CAD.Visible = True
    CAD.Select
    CAD.Unprotect Password:="roadway123"
    UserForm_MonthlyFuelData.Show vbModeless
End Sub

 Sub Open_MonthlyFuelData()
    UserForm_MonthlyFuelData.Show
End Sub

'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'*********************************** This section contains the Save functions for the Project Data**************************************>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
 Sub SaveModify_EngineersEstimate_Click()
'Project Data saved to Project Data table/sheet
    Application.ScreenUpdating = False
    PROJECT_DATA.Unprotect
    PROJECT_DATA.Visible = True
    PROJECT_DATA.Select
        With Me
        'LEFT Side of the Project Data Form
            Range("A2").Value = Me.txt_RoadName.Value
            Range("B2").Value = Me.txt_RoadName2.Value
            Range("C2").Value = Me.txt_StructureNo.Value
            Range("D2").Value = Me.cmb_County.Value
            Range("E2").Value = Me.cmb_CharacterofWork.Value
            Range("F2").Value = Me.txt_CharacterofWorkNotes.Value
            Range("G2").Value = Me.txt_RoadwayWidth.Value
            Range("H2").Value = Me.cmb_SurfaceType.Value
            Range("I2").Value = Me.txt_SurfaceWidth.Value
            Range("J2").Value = Me.txt_WorkingDays.Value
            Range("K2").Value = Me.txt_EngineerPercent.Value
        'RIGHT side of the Project Data Form
            Range("L2").Value = Me.cmb_ProjectType.Value
            Range("M2").Value = Me.txt_ProjectNo.Value
            Range("N2").Value = Me.txt_EstimateDate.Value
            Range("O2").Value = Me.txt_RoadwayLength.Value
            Range("P2").Value = Me.txt_BridgeLength.Value
            Range("Q2").Value = Range("O2").Value + Range("P2").Value
                txt_ProjectLength.Value = Range("Q2").Value
            Range("R2").Value = Me.txt_ExceptionLength.Value
            Range("S2").Value = Range("Q2").Value + Range("R2").Value  '*****This field is calculated and read only on the form*****
                txt_GrossProjectLength = Range("S2").Value
            Range("T2").Value = Me.txt_PreparedbyName.Value
            Range("U2").Value = Me.txt_PreparedbyTitle.Value
        'BOTTOM of the Project Data Form
            Range("V2").Value = Me.txt_Form900.Value    ' Form 900 statement
            Range("W2").Value = Me.txt_DistrictA.Value     ' Board President District
            Range("X2").Value = Me.txt_DistrictB.Value       '
            Range("Y2").Value = Me.txt_DistrictC.Value       '
            Range("Z2").Value = Me.txt_DistrictD.Value       '
            Range("AA2").Value = Me.txt_DistrictE.Value       '
            Range("AB2").Value = Me.txt_SupervisorA.Value       ' Board President Name
            Range("AC2").Value = Me.txt_SupervisorB.Value        ' B
            Range("AD2").Value = Me.txt_SupervisorC.Value        ' C
            Range("AE2").Value = Me.txt_SupervisorD.Value        ' D
            Range("AF2").Value = Me.txt_SupervisorE.Value        ' E
                Call Bottom_ProjectFormData
        'Default non-user added data to the ProjectData table / sheet
            Range("AL2").Value = "Engineers Estimate"
            Range("AM2").Value = Now
            Range("AN2").Value = ActiveWorkbook.BuiltinDocumentProperties("Last Author")
        End With
        
        Call Format_ProjectData
        
    ENGINEERS_ESTIMATE.Protect
    PROJECT_DATA.Protect
    PROJECT_DATA.Visible = False
    Application.ScreenUpdating = True
    ENGINEERS_ESTIMATE.Select
End Sub


 Sub Save_OSARCEstimate_Click()
'Project Data saved to Project Data table/sheet
    Application.ScreenUpdating = False
    PROJECT_DATA.Unprotect
    PROJECT_DATA.Visible = True
    PROJECT_DATA.Select
        With Me
        'LEFT Side of the Project Data Form
            Range("A3").Value = Me.txt_RoadName.Value
            Range("B3").Value = Me.txt_RoadName2.Value
            Range("C3").Value = Me.txt_StructureNo.Value
            Range("D3").Value = Me.cmb_County.Value
            Range("E3").Value = Me.cmb_CharacterofWork.Value
            Range("F3").Value = Me.txt_CharacterofWorkNotes.Value
            Range("G3").Value = Me.txt_RoadwayWidth.Value
            Range("H3").Value = Me.cmb_SurfaceType.Value
            Range("I3").Value = Me.txt_SurfaceWidth.Value
            Range("J3").Value = Me.txt_WorkingDays.Value
            Range("K3").Value = Me.txt_EngineerPercent.Value
        'RIGHT side of the Project Data Form
            Range("L3").Value = Me.cmb_ProjectType.Value
            Range("M3").Value = Me.txt_ProjectNo.Value
            Range("N3").Value = Me.txt_EstimateDate.Value
            Range("O3").Value = Me.txt_RoadwayLength.Value
            Range("P3").Value = Me.txt_BridgeLength.Value
            Range("Q3").Value = Range("O3").Value + Range("P3").Value
                txt_ProjectLength.Value = Range("Q3").Value
            Range("R3").Value = Me.txt_ExceptionLength.Value
            Range("S3").Value = Range("Q3").Value + Range("R3").Value  '*****This field is calculated and read only on the form*****
                txt_GrossProjectLength = Range("S3").Value
            Range("T3").Value = Me.txt_PreparedbyName.Value
            Range("U3").Value = Me.txt_PreparedbyTitle.Value
        'BOTTOM of the Project Data Form
            Call Bottom_ProjectFormData
        'Default non-user added data to the ProjectData table / sheet
            Range("AL3").Value = "OSARC Estimate"
            Range("AM3").Value = Now
            Range("AN3").Value = ActiveWorkbook.BuiltinDocumentProperties("Last Author")
        End With
        
        Call Format_ProjectData
        
    OSARC_ESTIMATE.Protect
    PROJECT_DATA.Protect
    PROJECT_DATA.Visible = False
    Application.ScreenUpdating = True
    OSARC_ESTIMATE.Select
End Sub

 Sub Save_BIDTABEstimate_Click()
'Project Data saved to Project Data table/sheet
    Application.ScreenUpdating = False
    PROJECT_DATA.Unprotect
    PROJECT_DATA.Visible = True
    PROJECT_DATA.Select
        With Me
        'LEFT Side of the Project Data Form
            Range("A4").Value = Me.txt_RoadName.Value
            Range("B4").Value = Me.txt_RoadName2.Value
            Range("C4").Value = Me.txt_StructureNo.Value
            Range("D4").Value = Me.cmb_County.Value
            Range("E4").Value = Me.cmb_CharacterofWork.Value
            Range("F4").Value = Me.txt_CharacterofWorkNotes.Value
            Range("G4").Value = Me.txt_RoadwayWidth.Value
            Range("H4").Value = Me.cmb_SurfaceType.Value
            Range("I4").Value = Me.txt_SurfaceWidth.Value
            Range("J4").Value = Me.txt_WorkingDays.Value
            Range("K4").Value = Me.txt_EngineerPercent.Value
        'RIGHT side of the Project Data Form
            Range("L4").Value = Me.cmb_ProjectType.Value
            Range("M4").Value = Me.txt_ProjectNo.Value
            Range("N4").Value = Me.txt_EstimateDate.Value
            Range("O4").Value = Me.txt_RoadwayLength.Value
            Range("P4").Value = Me.txt_BridgeLength.Value
            Range("Q4").Value = Range("O4").Value + Range("P4").Value
                txt_ProjectLength.Value = Range("Q4").Value
            Range("R4").Value = Me.txt_ExceptionLength.Value
            Range("S4").Value = Range("Q4").Value + Range("R4").Value  '*****This field is calculated and read only on the form*****
                txt_GrossProjectLength = Range("S4").Value
            Range("T4").Value = Me.txt_PreparedbyName.Value
            Range("U4").Value = Me.txt_PreparedbyTitle.Value
        'BOTTOM of the Project Data Form
            Call Bottom_ProjectFormData
        'Default non-user added data to the ProjectData table / sheet
            Range("AL4").Value = "Bidtab Estimate"
            Range("AM4").Value = Now
            Range("AN4").Value = ActiveWorkbook.BuiltinDocumentProperties("Last Author")
        End With
        
        Call Format_ProjectData
        
    BIDTAB.Protect
    PROJECT_DATA.Protect
    PROJECT_DATA.Visible = False
    Application.ScreenUpdating = True
    BIDTAB.Select
End Sub

 Sub Save_CONTRACTOREstimate_Click()
'Project Data saved to Project Data table/sheet
    Application.ScreenUpdating = False
    PROJECT_DATA.Unprotect
    PROJECT_DATA.Visible = True
    PROJECT_DATA.Select
        With Me
        'LEFT Side of the Project Data Form
            Range("A5").Value = Me.txt_RoadName.Value
            Range("B5").Value = Me.txt_RoadName2.Value
            Range("C5").Value = Me.txt_StructureNo.Value
            Range("D5").Value = Me.cmb_County.Value
            Range("E5").Value = Me.cmb_CharacterofWork.Value
            Range("F5").Value = Me.txt_CharacterofWorkNotes.Value
            Range("G5").Value = Me.txt_RoadwayWidth.Value
            Range("H5").Value = Me.cmb_SurfaceType.Value
            Range("I5").Value = Me.txt_SurfaceWidth.Value
            Range("J5").Value = Me.txt_WorkingDays.Value
            Range("K5").Value = Me.txt_EngineerPercent.Value
        'RIGHT side of the Project Data Form
            Range("L5").Value = Me.cmb_ProjectType.Value
            Range("M5").Value = Me.txt_ProjectNo.Value
            Range("N5").Value = Me.txt_EstimateDate.Value
            Range("O5").Value = Me.txt_RoadwayLength.Value
            Range("P5").Value = Me.txt_BridgeLength.Value
            Range("Q5").Value = Range("O5").Value + Range("P5").Value
                txt_ProjectLength.Value = Range("Q5").Value
            Range("R5").Value = Me.txt_ExceptionLength.Value
            Range("S5").Value = Range("Q5").Value + Range("R5").Value  '*****This field is calculated and read only on the form*****
                txt_GrossProjectLength = Range("S5").Value
            Range("T5").Value = Me.txt_PreparedbyName.Value
            Range("U5").Value = Me.txt_PreparedbyTitle.Value
        'BOTTOM of the Project Data Form
            Call Bottom_ProjectFormData
        'Default non-user added data to the ProjectData table / sheet
            Range("AL5").Value = "Contractors Estimate"
            Range("AM5").Value = Now
            Range("AN5").Value = ActiveWorkbook.BuiltinDocumentProperties("Last Author")
        End With
        
        Call Format_ProjectData
        
    CONTRACTOR_ESTIMATE.Protect
    PROJECT_DATA.Protect
    PROJECT_DATA.Visible = False
    Application.ScreenUpdating = True
    CONTRACTOR_ESTIMATE.Select
End Sub

'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'*********************************** This section contains the Load functions for the Project Data**************************************>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################

 Sub Load_EngineersEstimate_Click()
    'LEFT Side of the Project Data Form
        txt_RoadName.Value = PROJECT_DATA.Range("A2").Value
        txt_RoadName2.Value = PROJECT_DATA.Range("B2").Value
        txt_StructureNo.Value = PROJECT_DATA.Range("C2").Value
        cmb_County.Value = PROJECT_DATA.Range("D2").Value
        cmb_CharacterofWork.Value = PROJECT_DATA.Range("E2").Value
        txt_CharacterofWorkNotes.Value = PROJECT_DATA.Range("F2").Value
        txt_RoadwayWidth.Value = PROJECT_DATA.Range("G2").Value
        cmb_SurfaceType.Value = PROJECT_DATA.Range("H2").Value
        txt_SurfaceWidth.Value = PROJECT_DATA.Range("I2").Value
        txt_WorkingDays.Value = PROJECT_DATA.Range("J2").Value
        txt_EngineerPercent.Value = PROJECT_DATA.Range("K2").Value
    'RIGHT Side of the Project Data Form
        cmb_ProjectType.Value = PROJECT_DATA.Range("L2").Value
        txt_ProjectNo = PROJECT_DATA.Range("M2").Value
        txt_EstimateDate.Value = PROJECT_DATA.Range("N2").Value
        txt_RoadwayLength.Value = PROJECT_DATA.Range("O2").Value
        txt_BridgeLength.Value = PROJECT_DATA.Range("P2").Value
        txt_ProjectLength.Value = PROJECT_DATA.Range("Q2").Value
        txt_ExceptionLength.Value = PROJECT_DATA.Range("R2").Value
        txt_GrossProjectLength.Value = PROJECT_DATA.Range("S2").Value '*****This field is calculated and read only on the form*****
        txt_PreparedbyName.Value = PROJECT_DATA.Range("T2").Value
        txt_PreparedbyTitle.Value = PROJECT_DATA.Range("U2").Value
    'BOTTOM of the Project Data Form
        txt_Form900.Value = PROJECT_DATA.Range("V2").Value   ' Form 900 statement
        txt_DistrictA.Value = PROJECT_DATA.Range("W2").Value    ' Board President District
        txt_DistrictB.Value = PROJECT_DATA.Range("X2").Value      '
        txt_DistrictC.Value = PROJECT_DATA.Range("Y2").Value      '
        txt_DistrictD.Value = PROJECT_DATA.Range("Z2").Value      '
        txt_DistrictE.Value = PROJECT_DATA.Range("AA2").Value      '
        txt_SupervisorA.Value = PROJECT_DATA.Range("AB2").Value      ' Board President Name
        txt_SupervisorB.Value = PROJECT_DATA.Range("AC2").Value       ' B
        txt_SupervisorC.Value = PROJECT_DATA.Range("AD2").Value       ' C
        txt_SupervisorD.Value = PROJECT_DATA.Range("AE2").Value        ' D
        txt_SupervisorE.Value = PROJECT_DATA.Range("AF2").Value       ' E
End Sub

 Sub Load_OSARCEstimate_Click()
    'LEFT Side of the Project Data Form
        txt_RoadName.Value = PROJECT_DATA.Range("A3").Value
        txt_RoadName2.Value = PROJECT_DATA.Range("B3").Value
        txt_StructureNo.Value = PROJECT_DATA.Range("C3").Value
        cmb_County.Value = PROJECT_DATA.Range("D3").Value
        cmb_CharacterofWork.Value = PROJECT_DATA.Range("E3").Value
        txt_CharacterofWorkNotes.Value = PROJECT_DATA.Range("F3").Value
        txt_RoadwayWidth.Value = PROJECT_DATA.Range("G3").Value
        cmb_SurfaceType.Value = PROJECT_DATA.Range("H3").Value
        txt_SurfaceWidth.Value = PROJECT_DATA.Range("I3").Value
        txt_WorkingDays.Value = PROJECT_DATA.Range("J3").Value
        txt_EngineerPercent.Value = PROJECT_DATA.Range("K3").Value
    'RIGHT Side of the Project Data Form
        cmb_ProjectType.Value = PROJECT_DATA.Range("L3").Value
        txt_ProjectNo = PROJECT_DATA.Range("M3").Value
        txt_EstimateDate.Value = PROJECT_DATA.Range("N3").Value
        txt_RoadwayLength.Value = PROJECT_DATA.Range("O3").Value
        txt_BridgeLength.Value = PROJECT_DATA.Range("P3").Value
        txt_ProjectLength.Value = PROJECT_DATA.Range("Q3").Value
        txt_ExceptionLength.Value = PROJECT_DATA.Range("R3").Value
        txt_GrossProjectLength.Value = PROJECT_DATA.Range("S3").Value '*****This field is calculated and read only on the form*****
        txt_PreparedbyName.Value = PROJECT_DATA.Range("T3").Value
        txt_PreparedbyTitle.Value = PROJECT_DATA.Range("U3").Value
    'BOTTOM of the Project Data Form
        DistrictSelection.Value = PROJECT_DATA.Range("AG3").Value   ' District
        txt_BidDate.Value = PROJECT_DATA.Range("AH3").Value ' Bid Date
        txt_latitude.Value = PROJECT_DATA.Range("AI3").Value ' Latitude
        txt_longitude.Value = PROJECT_DATA.Range("AJ3").Value ' Longitude
        txt_Project_fk.Value = PROJECT_DATA.Range("AK3").Value ' pk_Projects
End Sub

 Sub Load_BIDTABEstimate_Click()
    'LEFT Side of the Project Data Form
        txt_RoadName.Value = PROJECT_DATA.Range("A4").Value
        txt_RoadName2.Value = PROJECT_DATA.Range("B4").Value
        txt_StructureNo.Value = PROJECT_DATA.Range("C4").Value
        cmb_County.Value = PROJECT_DATA.Range("D4").Value
        cmb_CharacterofWork.Value = PROJECT_DATA.Range("E4").Value
        txt_CharacterofWorkNotes.Value = PROJECT_DATA.Range("F4").Value
        txt_RoadwayWidth.Value = PROJECT_DATA.Range("G4").Value
        cmb_SurfaceType.Value = PROJECT_DATA.Range("H4").Value
        txt_SurfaceWidth.Value = PROJECT_DATA.Range("I4").Value
        txt_WorkingDays.Value = PROJECT_DATA.Range("J4").Value
        txt_EngineerPercent.Value = PROJECT_DATA.Range("K4").Value
    'RIGHT Side of the Project Data Form
        cmb_ProjectType.Value = PROJECT_DATA.Range("L4").Value
        txt_ProjectNo = PROJECT_DATA.Range("M4").Value
        txt_EstimateDate.Value = PROJECT_DATA.Range("N4").Value
        txt_RoadwayLength.Value = PROJECT_DATA.Range("O4").Value
        txt_BridgeLength.Value = PROJECT_DATA.Range("P4").Value
        txt_ProjectLength.Value = PROJECT_DATA.Range("Q4").Value
        txt_ExceptionLength.Value = PROJECT_DATA.Range("R4").Value
        txt_GrossProjectLength.Value = PROJECT_DATA.Range("S4").Value '*****This field is calculated and read only on the form*****
        txt_PreparedbyName.Value = PROJECT_DATA.Range("T4").Value
        txt_PreparedbyTitle.Value = PROJECT_DATA.Range("U4").Value
    'BOTTOM of the Project Data Form
        DistrictSelection.Value = PROJECT_DATA.Range("AG4").Value   ' District
        txt_BidDate.Value = PROJECT_DATA.Range("AH4").Value ' Bid Date
        txt_latitude.Value = PROJECT_DATA.Range("AI4").Value ' Latitude
        txt_longitude.Value = PROJECT_DATA.Range("AJ4").Value ' Longitude
        txt_Project_fk.Value = PROJECT_DATA.Range("AK4").Value ' pk_Projects
End Sub


 Sub Load_CONTRACTOREstimate_Click()
    'LEFT Side of the Project Data Form
        txt_RoadName.Value = PROJECT_DATA.Range("A5").Value
        txt_RoadName2.Value = PROJECT_DATA.Range("B5").Value
        txt_StructureNo.Value = PROJECT_DATA.Range("C5").Value
        cmb_County.Value = PROJECT_DATA.Range("D5").Value
        cmb_CharacterofWork.Value = PROJECT_DATA.Range("E5").Value
        txt_CharacterofWorkNotes.Value = PROJECT_DATA.Range("F5").Value
        txt_RoadwayWidth.Value = PROJECT_DATA.Range("G5").Value
        cmb_SurfaceType.Value = PROJECT_DATA.Range("H5").Value
        txt_SurfaceWidth.Value = PROJECT_DATA.Range("I5").Value
        txt_WorkingDays.Value = PROJECT_DATA.Range("J5").Value
        txt_EngineerPercent.Value = PROJECT_DATA.Range("K5").Value
    'RIGHT Side of the Project Data Form
        cmb_ProjectType.Value = PROJECT_DATA.Range("L5").Value
        txt_ProjectNo = PROJECT_DATA.Range("M5").Value
        txt_EstimateDate.Value = PROJECT_DATA.Range("N5").Value
        txt_RoadwayLength.Value = PROJECT_DATA.Range("O5").Value
        txt_BridgeLength.Value = PROJECT_DATA.Range("P5").Value
        txt_ProjectLength.Value = PROJECT_DATA.Range("Q5").Value
        txt_ExceptionLength.Value = PROJECT_DATA.Range("R5").Value
        txt_GrossProjectLength.Value = PROJECT_DATA.Range("S5").Value '*****This field is calculated and read only on the form*****
        txt_PreparedbyName.Value = PROJECT_DATA.Range("T5").Value
        txt_PreparedbyTitle.Value = PROJECT_DATA.Range("U5").Value
    'BOTTOM of the Project Data Form
        DistrictSelection.Value = PROJECT_DATA.Range("AG5").Value   ' District
        txt_BidDate.Value = PROJECT_DATA.Range("AH5").Value ' Bid Date
        txt_latitude.Value = PROJECT_DATA.Range("AI5").Value ' Latitude
        txt_longitude.Value = PROJECT_DATA.Range("AJ5").Value ' Longitude
        txt_Project_fk.Value = PROJECT_DATA.Range("AK5").Value ' pk_Projects
End Sub

'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'*********************************** This section contains the Initialize / Active functions for the forms ******************************>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'
' Initialize values from Project Data table/sheet into UserForm_ProjectData
'
 Sub UserForm_Initialize()
    'LEFT Side of the Project Data Form
        txt_RoadName.Value = PROJECT_DATA.Range("A2").Value
        txt_RoadName2.Value = PROJECT_DATA.Range("B2").Value
        txt_StructureNo.Value = PROJECT_DATA.Range("C2").Value
        cmb_County.Value = PROJECT_DATA.Range("D2").Value
        cmb_CharacterofWork.Value = PROJECT_DATA.Range("E2").Value
        txt_CharacterofWorkNotes.Value = PROJECT_DATA.Range("F2").Value
        txt_RoadwayWidth.Value = PROJECT_DATA.Range("G2").Value
        cmb_SurfaceType.Value = PROJECT_DATA.Range("H2").Value
        txt_SurfaceWidth.Value = PROJECT_DATA.Range("I2").Value
        txt_WorkingDays.Value = PROJECT_DATA.Range("J2").Value
        txt_EngineerPercent.Value = PROJECT_DATA.Range("K2").Value
    'RIGHT Side of the Project Data Form
        cmb_ProjectType.Value = PROJECT_DATA.Range("L2").Value
        txt_ProjectNo = PROJECT_DATA.Range("M2").Value
        txt_EstimateDate.Value = PROJECT_DATA.Range("N2").Value
        txt_RoadwayLength.Value = PROJECT_DATA.Range("O2").Value
        txt_BridgeLength.Value = PROJECT_DATA.Range("P2").Value
        txt_ProjectLength.Value = PROJECT_DATA.Range("Q2").Value
        txt_ExceptionLength.Value = PROJECT_DATA.Range("R2").Value
        txt_GrossProjectLength.Value = PROJECT_DATA.Range("S2").Value '*****This field is calculated and read only on the form*****
        txt_PreparedbyName.Value = PROJECT_DATA.Range("T2").Value
        txt_PreparedbyTitle.Value = PROJECT_DATA.Range("U2").Value
    'BOTTOM of the Project Data Form
        txt_Form900.Value = PROJECT_DATA.Range("V2").Value   ' Form 900 statement
        txt_DistrictA.Value = PROJECT_DATA.Range("W2").Value    ' Board President District
        txt_DistrictB.Value = PROJECT_DATA.Range("X2").Value      '
        txt_DistrictC.Value = PROJECT_DATA.Range("Y2").Value      '
        txt_DistrictD.Value = PROJECT_DATA.Range("Z2").Value      '
        txt_DistrictE.Value = PROJECT_DATA.Range("AA2").Value      '
        txt_SupervisorA.Value = PROJECT_DATA.Range("AB2").Value      ' Board President Name
        txt_SupervisorB.Value = PROJECT_DATA.Range("AC2").Value       ' B
        txt_SupervisorC.Value = PROJECT_DATA.Range("AD2").Value       ' C
        txt_SupervisorD.Value = PROJECT_DATA.Range("AE2").Value        ' D
        txt_SupervisorE.Value = PROJECT_DATA.Range("AF2").Value       ' E
    'Invoice Tab for the CAD template data
        txt_VendorNo.Value = CAD.Range("B3").Value
        txt_ContractorNameAddress.Value = CAD.Range("B5").Value
        txt_Surety.Value = CAD.Range("G6").Value
        txt_Email.Value = CAD.Range("B6").Value
        txt_Retainage.Value = CAD.Range("B13").Value
        txt_WorkingDaysUpdate.Value = PROJECT_DATA.Range("J2").Value
        cmb_EngineeringFunding.Value = PROJECT_DATA.Range("AO2").Value
    'Invoice Tab for the Engineers Payment data
        txt_CEVendorNo.Value = EngPayInv.Range("B5").Value
        txt_CEVendorAddress.Value = EngPayInv.Range("D5").Value
        txt_ContractAmount.Value = EngPayInv.Range("D11").Value
        txt_PaidTo.Value = EngPayInv.Range("C18").Value
        txt_AgreementDate.Value = EngPayInv.Range("B20").Value
        txt_ApprovedSAEngineer.Value = EngPayInv.Range("B21").Value
        cmb_EngineeringFunding.Value = PROJECT_DATA.Range("AO2").Value
     ' Specification Year Use
        DATA_VALIDATION.Unprotect
        DATA_VALIDATION.Range("Y2").Value = "2004"
        txt_SpecificationYearSelected.Value = DATA_VALIDATION.Range("Y2").Value
    ' Call update Type selection based on the construction status
        Call UpdateTypeComboBox
        DATA_VALIDATION.Protect
End Sub

Sub UserForm_Activate()
' Initialize the Pay Item Form list box from the database
    Call AddPayItemToListBox_Database
End Sub


Sub ListBox_Database_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'Double click on a row in the list box and populates the form with these values from the Database'
        Me.cmb_PayItemSearch.Value = ""
        Me.txt_PayItemNo.Value = Me.ListBox_Database.List(Me.ListBox_Database.ListIndex, 0)
        Me.txt_PayItemDescription.Value = Me.ListBox_Database.List(Me.ListBox_Database.ListIndex, 1)
        Me.txt_Quantity.Value = Me.ListBox_Database.List(Me.ListBox_Database.ListIndex, 2)
        Me.txt_Unit.Value = Me.ListBox_Database.List(Me.ListBox_Database.ListIndex, 3)
        Me.txt_UnitPrice.Value = Me.ListBox_Database.List(Me.ListBox_Database.ListIndex, 4)
        Me.txt_Subtotal.Value = Me.ListBox_Database.List(Me.ListBox_Database.ListIndex, 5)
        Me.cmb_FuelCode.Value = Me.ListBox_Database.List(Me.ListBox_Database.ListIndex, 7)
        Me.cmb_Type.Value = Me.ListBox_Database.List(Me.ListBox_Database.ListIndex, 45)
        Me.cmb_Participating.Value = Me.ListBox_Database.List(Me.ListBox_Database.ListIndex, 46)
        Me.txt_SortOrder.Value = Me.ListBox_Database.List(Me.ListBox_Database.ListIndex, 47)
        Me.txt_SpecYear.Value = Me.ListBox_Database.List(Me.ListBox_Database.ListIndex, 69)
End Sub



Sub Bottom_ProjectFormData()
            Range("AG2").Value = DistrictSelection.Value  ' District
            Range("AH2").Value = txt_BidDate.Value  ' Bid Date
            Range("AI2").Value = txt_latitude.Value ' Latitude
            Range("AJ2").Value = txt_longitude.Value ' Longitude
            Range("AK2").Value = txt_Project_fk.Value ' pk_Projects
            
            Range("AG3").Value = DistrictSelection.Value  ' District
            Range("AH3").Value = txt_BidDate.Value  ' Bid Date
            Range("AI3").Value = txt_latitude.Value ' Latitude
            Range("AJ3").Value = txt_longitude.Value ' Longitude
            Range("AK3").Value = txt_Project_fk.Value ' pk_Projects
            
            Range("AG4").Value = DistrictSelection.Value  ' District
            Range("AH4").Value = txt_BidDate.Value  ' Bid Date
            Range("AI4").Value = txt_latitude.Value ' Latitude
            Range("AJ4").Value = txt_longitude.Value ' Longitude
            Range("AK4").Value = txt_Project_fk.Value ' pk_Projects
            
            Range("AG5").Value = DistrictSelection.Value  ' District
            Range("AH5").Value = txt_BidDate.Value  ' Bid Date
            Range("AI5").Value = txt_latitude.Value ' Latitude
            Range("AJ5").Value = txt_longitude.Value ' Longitude
            Range("AK5").Value = txt_Project_fk.Value ' pk_Projects
End Sub

Sub UpdateTypeComboBox()
        Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DATA_VALIDATION") ' Change to your sheet name
    
    Select Case ws.Range("Z2").Value ' Change A1 to your cell
        Case "No"
            Me.cmb_Type.RowSource = "PAYITEMTYPE!A2:A55" ' Change to your range name
        Case "Yes"
            Me.cmb_Type.RowSource = "PAYITEMTYPE_SUPPLEMENTAL!A2:A51"
        ' Add more cases as needed
    End Select
End Sub

