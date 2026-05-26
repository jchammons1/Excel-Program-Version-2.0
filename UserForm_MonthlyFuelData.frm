VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_MonthlyFuelData 
   Caption         =   "CURRENT MONTH Fuel Adjustment Data Form"
   ClientHeight    =   13050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10320
   OleObjectBlob   =   "UserForm_MonthlyFuelData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_MonthlyFuelData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=== Your "Anyone" share link (the :x: link you provided)
Private Const FUELPRICES_SHARELINK As String = _
"https://osarc.sharepoint.com/:x:/s/OSARCEstimateInvoiceProgram/IQC88MFKJi9IRrY58bCNn0zBAbWPGLmDfLTQLtpFh6MPonE?e=cOOVrN"

'Convert SharePoint Copy Link into a direct-download URL.
'Method: replace everything after ? with download=1 [1](https://www.sharepointdiary.com/2020/05/sharepoint-online-link-to-document-download-instead-of-open.html)[2](https://pennstate.service-now.com/sp?id=kb_article_view&sysparm_article=KB0016080&sys_kb_id=87881ea99755e5103dbab577f053afae&spa=1)
Public Function SharePointDirectDownloadUrl(ByVal shareUrl As String) As String
    Dim p As Long
    p = InStr(1, shareUrl, "?", vbTextCompare)
    If p > 0 Then
        SharePointDirectDownloadUrl = Left$(shareUrl, p) & "download=1"
    Else
        SharePointDirectDownloadUrl = shareUrl & "?download=1"
    End If
End Function

'Download a URL to %TEMP% and return the local file path
Public Function DownloadToTemp(ByVal fileUrl As String, ByVal localFileName As String) As String
    Dim localPath As String
    localPath = Environ$("TEMP") & "\" & localFileName

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Option(6) = True ' WinHttpRequestOption_EnableRedirects
    http.Open "GET", fileUrl, False
    http.Send

    If http.Status <> 200 Then
        Err.Raise vbObjectError + 710, , "Download failed. HTTP Status: " & http.Status
    End If

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 'binary
    stm.Open
    stm.Write http.ResponseBody
    stm.SaveToFile localPath, 2 'overwrite
    stm.Close

    DownloadToTemp = localPath
End Function

'Find a column index in a 2D array by header name (row 1)
Private Function FindHeaderCol(ByVal data As Variant, ByVal headerName As String) As Long
    Dim c As Long
    For c = 1 To UBound(data, 2)
        If StrComp(CStr(data(1, c)), headerName, vbTextCompare) = 0 Then
            FindHeaderCol = c
            Exit Function
        End If
    Next c
    Err.Raise vbObjectError + 711, , "Header not found: " & headerName
End Function

'Return the row index that matches Month + Year; 0 if not found
Private Function FindFuelRow(ByVal data As Variant, ByVal MonthSelected As String, ByVal YearSelected As Long) As Long
    Dim colMonth As Long, colYear As Long
    colMonth = FindHeaderCol(data, "Month")
    colYear = FindHeaderCol(data, "Year")

    Dim r As Long
    For r = 2 To UBound(data, 1)
        If StrComp(CStr(data(r, colMonth)), MonthSelected, vbTextCompare) = 0 Then
            'Year in your export appears numeric; handle "2022.0" style safely [4](https://osarc.sharepoint.com/sites/OSARCEstimateInvoiceProgram/Lists/Fuel_Prices)
            If CLng(Val(CStr(data(r, colYear)))) = CLng(YearSelected) Then
                FindFuelRow = r
                Exit Function
            End If
        End If
    Next r

    FindFuelRow = 0
End Function

'Load the SharePoint Fuel_Prices.xlsx into a 2D array (UsedRange)
Public Function LoadFuelPricesData() As Variant
    Dim dlUrl As String, localPath As String
    dlUrl = SharePointDirectDownloadUrl(FUELPRICES_SHARELINK)
    localPath = DownloadToTemp(dlUrl, "Fuel_Prices.xlsx") ' workbook name [3](https://osarc.sharepoint.com/sites/OSARCEstimateInvoiceProgram/Fuel%20Prices/Fuel_Prices.xlsx)

    Dim wb As Workbook, ws As Worksheet
    Set wb = Workbooks.Open(localPath, ReadOnly:=True)
    Set ws = wb.Worksheets(1)

    LoadFuelPricesData = ws.UsedRange.Value2

    wb.Close SaveChanges:=False
End Function

'Get a value from the data array by header name for a given row
Private Function ValueByHeader(ByVal data As Variant, ByVal rowIndex As Long, ByVal headerName As String) As Variant
    Dim c As Long
    c = FindHeaderCol(data, headerName)
    ValueByHeader = data(rowIndex, c)
End Function





Private Sub Button_FuelDamages_Click()
'######################################################### FUEL ADJUSTMENT CALCULATIONS ##############################################
' Adding monthly Fuel Adjustment Data to the Active CAD sheet
If MsgBox("Do you want to discontinue fuel adjustments? This will delete the base and current fuel adjustments on this invoice and subsequent invoices", vbYesNo, "Save changes") = vbYes Then

    Application.ScreenUpdating = False
       Database.Visible = True
       Database.Unprotect
       CAD.Visible = True
       CAD.Select
       CAD.Unprotect
       
        With Me     ' I am not sure which types of the materials are associated with the pay item codes below
            'Base Data
                Range("GF2") = "Damages"
                Range("GF3") = "Damages"
                Range("GF4") = 0
                Range("GF5") = 0
                Range("GF6") = 0
                Range("GF7") = 0
                Range("GF8") = 0
                Range("GF9") = 0
                Range("GF10") = 0
                Range("GF11") = 0
                Range("GF12") = 0
                Range("GF13") = 0
                Range("GF14") = 0
                Range("GF15") = 0
         'Current Data
                Range("GG2") = "Damages"
                Range("GG3") = "Damages"
                Range("GG4") = 0
                Range("GG5") = 0
                Range("GG6") = 0
                Range("GG7") = 0
                Range("GG8") = 0
                Range("GG9") = 0
                Range("GG10") = 0
                Range("GG11") = 0
                Range("GG12") = 0
                Range("GG13") = 0
                Range("GG14") = 0
                Range("GG15") = 0
        End With
       
        
        With Me     ' Refreshes and shows the changed values
                Range("GG2") = Me.cmb_CurrentMonth.Value
                Range("GG3") = Me.cmb_CurrentYear.Value
                Range("GG4") = Me.txt_CurrentGasoline.Value
                Range("GG5") = Me.txt_CurrentDiesel.Value
                Range("GG6") = Me.txt_CurrentPG6422.Value
                Range("GG7") = Me.txt_CurrentPG6722.Value
                Range("GG8") = Me.txt_CurrentPG7622.Value
                Range("GG9") = Me.txt_CurrentPG8222.Value
                Range("GG10") = Me.txt_CurrentSS1.Value
                Range("GG11") = Me.txt_CurrentCRS2.Value
                Range("GG12") = Me.txt_CurrentCRS2P.Value
                Range("GG13") = Me.txt_CurrentEA1.Value
                Range("GG14") = Me.txt_CurrentCSS1_Undiluted.Value
                Range("GG15") = Me.txt_CurrentCSS1.Value
        End With
        
        CAD.Range("GF4:GH15").Select
        Selection.NumberFormat = "0.0000"
       
    ' Adds the current fuel adjustment data to the Database
    ' Populates the row that matches the fuel code
        Database.Visible = True
        Database.Select
        Database.Unprotect
                row = 2
                column = 1
                endrow = CountPayItems - 46
        Do While row < endrow
            If Cells(row, column + 7).Value = "E" Or Cells(row, column + 7).Value = "GY" Or Cells(row, column + 7).Value = "GT" Or Cells(row, column + 7).Value = "M" _
            Or Cells(row, column + 7).Value = "B" Or Cells(row, column + 7).Value = "D" Or Cells(row, column + 7).Value = "C" Or Cells(row, column + 7).Value = "S" _
            Or Cells(row, column + 7).Value = "BA1" Or Cells(row, column + 7).Value = "A1" Or Cells(row, column + 7).Value = "A2" Or Cells(row, column + 7).Value = "A3" _
            Or Cells(row, column + 7).Value = "A4" Or Cells(row, column + 7).Value = "A5" Or Cells(row, column + 7).Value = "A6" Then
                            
            ' Base
                Cells(row, column + 9) = CAD.Range("GF2").Value         'Fuel Base Month
                Cells(row, column + 10) = CAD.Range("GF3").Value        'Fuel Base Year
                    Cells(row, column + 10).Select
                    Selection.NumberFormat = "0"
                Cells(row, column + 13) = CAD.Range("GF4").Value         'Base Gas
                Cells(row, column + 14) = CAD.Range("GF5").Value         'Base Diesel
                Cells(row, column + 19) = CAD.Range("GF6").Value         'Base PG-64-22
                Cells(row, column + 20) = CAD.Range("GF7").Value         'Base PG-67-22
                Cells(row, column + 21) = CAD.Range("GF8").Value         'Base PG-76-22
                Cells(row, column + 22) = CAD.Range("GF9").Value         'Base PG-82-22
                Cells(row, column + 23) = CAD.Range("GF10").Value         'Base SS-1
                Cells(row, column + 24) = CAD.Range("GF11").Value         'Base CRS-2
                Cells(row, column + 25) = CAD.Range("GF12").Value         'Base CRS-2P
                Cells(row, column + 26) = CAD.Range("GF13").Value         'Base EA-1
                Cells(row, column + 27) = CAD.Range("GF14").Value         'Base CSS-1 undiluted
                Cells(row, column + 28) = CAD.Range("GF15").Value         'Base CSS-1
                
            ' Current
                Cells(row, column + 11) = Database.Range("GG2").Value         'Fuel Current Month
                Cells(row, column + 12) = Database.Range("GG3").Value        'Fuel Current Year
                    Cells(row, column + 12).Select
                    Selection.NumberFormat = "0"
                Cells(row, column + 15) = Database.Range("GG4").Value         'Current Gas
                Cells(row, column + 16) = Database.Range("GG5").Value         'Current Diesel
                Cells(row, column + 29) = Database.Range("GG6").Value         'Current PG-64-22
                Cells(row, column + 30) = Database.Range("GG7").Value         'Current PG-67-22
                Cells(row, column + 31) = Database.Range("GG8").Value         'Current PG-76-22
                Cells(row, column + 32) = Database.Range("GG9").Value         'Current PG-82-22
                Cells(row, column + 33) = Database.Range("GG10").Value         'Current SS-1
                Cells(row, column + 34) = Database.Range("GG11").Value         'Current CRS-2
                Cells(row, column + 35) = Database.Range("GG12").Value         'Current CRS-2P
                Cells(row, column + 36) = Database.Range("GG13").Value         'Current EA-1
                Cells(row, column + 37) = Database.Range("GG14").Value         'Current CSS-1 undiluted
                Cells(row, column + 38) = Database.Range("GG15").Value         'Current CSS-1
            End If
        row = row + 1
        Loop
        
        Database.Protect
        Database.Visible = False
        CAD.Select
        CAD.Protect
        CAD.Visible = False

        'Application.ScreenUpdating = True
        ActiveSheet.Range("A2").Select
    End If
    Unload Me
End Sub

Private Sub Button_ImportCurrentFuel_Click()

    Dim MonthSelected As String
    Dim YearSelected As Long

    MonthSelected = cmb_CurrentMonth.Value
    YearSelected = CLng(cmb_CurrentYear.Value)

    Dim data As Variant
    Dim matchRow As Long

    On Error GoTo ErrHandler

    'Pull latest workbook from SharePoint (Anyone link) and read to array
    data = LoadFuelPricesData()

    'Find the matching row
    matchRow = FindFuelRow(data, MonthSelected, YearSelected)

    If matchRow > 0 Then
        txt_CurrentGasoline.Value = ValueByHeader(data, matchRow, "Gasoline")
        txt_CurrentDiesel.Value = ValueByHeader(data, matchRow, "Diesel")
        txt_CurrentPG6422.Value = ValueByHeader(data, matchRow, "PG6422")
        txt_CurrentPG6722.Value = ValueByHeader(data, matchRow, "PG6722")
        txt_CurrentPG7622.Value = ValueByHeader(data, matchRow, "PG7622")
        txt_CurrentPG8222.Value = ValueByHeader(data, matchRow, "PG8222")
        txt_CurrentSS1.Value = ValueByHeader(data, matchRow, "SS1")
        txt_CurrentCRS2.Value = ValueByHeader(data, matchRow, "RS2C_CRS2")
        txt_CurrentCRS2P.Value = ValueByHeader(data, matchRow, "CRS2P")
        txt_CurrentEA1.Value = ValueByHeader(data, matchRow, "EA1_ERP1_AEP")
        txt_CurrentCSS1_Undiluted.Value = ValueByHeader(data, matchRow, "CSS1_Undiluted")
        txt_CurrentCSS1.Value = ValueByHeader(data, matchRow, "CSS1_1H")
    Else
        MsgBox "The current month Current fuel data selected is not currently available. Check back later.", vbExclamation
    End If

    Exit Sub

ErrHandler:
    MsgBox "Fuel import failed: " & Err.Description, vbCritical
End Sub

 Sub Button_SaveMonthlyFuelAdjustment_Click()
'######################################################### FUEL ADJUSTMENT CALCULATIONS ##############################################
' Adding monthly Fuel Adjustment Data to the Active CAD sheet

' Set the message.  After changin the below Do while code to Select Case for the fuel code, the processing message is not needed because it is a lot faster
    UserForm_MonthlyFuelData.lblMessage.Caption = "Processing, please wait..."
    DoEvents
    
    Call DownloadWAVFile
    
Application.ScreenUpdating = False
   Database.Unprotect
   CAD.Visible = True
   CAD.Select
   CAD.Unprotect Password:="roadway123"
   
If Me.cmb_CurrentMonth.Value = "" Or Me.cmb_CurrentYear.Value = "" Or Me.txt_CurrentGasoline.Value = "" Or Me.txt_CurrentDiesel.Value = "" Or _
Me.txt_CurrentPG6422.Value = "" Or Me.txt_CurrentPG6722.Value = "" Or Me.txt_CurrentPG7622.Value = "" Or Me.txt_CurrentPG8222.Value = "" Or Me.txt_CurrentSS1.Value = "" _
Or Me.txt_CurrentCRS2.Value = "" Or Me.txt_CurrentCRS2P.Value = "" Or Me.txt_CurrentEA1.Value = "" Or Me.txt_CurrentCSS1_Undiluted.Value = "" _
Or Me.txt_CurrentCSS1.Value = "" Then

MsgBox "Please ensure all fields are populated or contain a value of 0.  User results did not save."
Exit Sub
End If
    
    With Me     ' I am not sure which types of the materials are associated with the pay item codes below
            Range("GG2") = Me.cmb_CurrentMonth.Value
            Range("GG3") = Me.cmb_CurrentYear.Value
            Range("GG4") = Me.txt_CurrentGasoline.Value
            Range("GG5") = Me.txt_CurrentDiesel.Value
            Range("GG6") = Me.txt_CurrentPG6422.Value
            Range("GG7") = Me.txt_CurrentPG6722.Value
            Range("GG8") = Me.txt_CurrentPG7622.Value
            Range("GG9") = Me.txt_CurrentPG8222.Value
            Range("GG10") = Me.txt_CurrentSS1.Value
            Range("GG11") = Me.txt_CurrentCRS2.Value
            Range("GG12") = Me.txt_CurrentCRS2P.Value
            Range("GG13") = Me.txt_CurrentEA1.Value
            Range("GG14") = Me.txt_CurrentCSS1_Undiluted.Value
            Range("GG15") = Me.txt_CurrentCSS1.Value
    End With
    
    CAD.Range("GF4:GH15").Select
    Selection.NumberFormat = "0.0000"
   
' Adds the current fuel adjustment data to the Database
' Populates the row that matches the fuel code

' Since I am using Option Explicit at the Top of this code.  The variables now must be defined for row, column, endrow,
    Dim row As Long
    Dim column As Long
    Dim endrow As Long

    
    Database.Visible = True
    Database.Select
    Database.Unprotect
            row = 2
            column = 1
            endrow = CountPayItems - 46
            
    Do While row < endrow
    
    Dim FuelCode As String
    
    FuelCode = Cells(row, column + 7).Value
    
        Select Case Cells(row, column + 7).Value
            Case "E", "GY", "GT", "M", "B", "D", "C", "S", "BA1", "A1", "A2", "A3", "A4", "A5", "A6"
            
        ''"E" Or Cells(row, column + 7).Value = "GY" Or Cells(row, column + 7).Value = "GT" Or Cells(row, column + 7).Value = "M" _
        Or Cells(row, column + 7).Value = "B" Or Cells(row, column + 7).Value = "D" Or Cells(row, column + 7).Value = "C" Or Cells(row, column + 7).Value = "S" _
        Or Cells(row, column + 7).Value = "BA1" Or Cells(row, column + 7).Value = "A1" Or Cells(row, column + 7).Value = "A2" Or Cells(row, column + 7).Value = "A3" _
        Or Cells(row, column + 7).Value = "A4" Or Cells(row, column + 7).Value = "A5" Or Cells(row, column + 7).Value = "A6" Then
            
            'Current
            Cells(row, column + 11) = CAD.Range("GG2").Value         'Fuel Current Month
            Cells(row, column + 12) = CAD.Range("GG3").Value        'Fuel Current Year
                Cells(row, column + 12).Select
                Selection.NumberFormat = "0"
            Cells(row, column + 15) = CAD.Range("GG4").Value         'Current Gas
            Cells(row, column + 16) = CAD.Range("GG5").Value         'Current Diesel
            Cells(row, column + 29) = CAD.Range("GG6").Value         'Current PG-64-22
            Cells(row, column + 30) = CAD.Range("GG7").Value         'Current PG-67-22
            Cells(row, column + 31) = CAD.Range("GG8").Value         'Current PG-76-22
            Cells(row, column + 32) = CAD.Range("GG9").Value         'Current PG-82-22
            Cells(row, column + 33) = CAD.Range("GG10").Value         'Current SS-1
            Cells(row, column + 34) = CAD.Range("GG11").Value         'Current CRS-2
            Cells(row, column + 35) = CAD.Range("GG12").Value         'Current CRS-2P
            Cells(row, column + 36) = CAD.Range("GG13").Value         'Current EA-1
            Cells(row, column + 37) = CAD.Range("GG14").Value         'Current CSS-1 undiluted
            Cells(row, column + 38) = CAD.Range("GG15").Value         'Current CSS-1
        End Select
        
        'End If
        
    ' Show which row is processing
        UserForm_MonthlyFuelData.lblMessage.Caption = "Processing: " & row & " of " & endrow
        DoEvents
        
    row = row + 1
    Loop
    
    'Clear caption
    UserForm_MonthlyFuelData.lblMessage.Caption = ""
    DoEvents
    
    Database.Protect
    Database.Visible = False
    CAD.Select
    CAD.Protect Password:="roadway123"
    CAD.Visible = False
    Application.ScreenUpdating = True
    ActiveSheet.Range("A2").Select
    Unload Me
End Sub



Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        ' Hide the sheet.  When you X close the CAD stays visible if this is not here.
        CAD.Visible = xlSheetHidden
    End If
End Sub

 Sub UserForm_Initialize()
    CAD.Select
    With Me     'Load Values from the monthly fuel data located on the Active Sheet CAD in Column GG
        Me.cmb_CurrentMonth.Value = Range("GG2")
        Me.cmb_CurrentYear.Value = Range("GG3")
        Me.txt_CurrentGasoline.Value = Range("GG4")
        Me.txt_CurrentDiesel.Value = Range("GG5")
        Me.txt_CurrentPG6422.Value = Range("GG6")
        Me.txt_CurrentPG6722.Value = Range("GG7")
        Me.txt_CurrentPG7622.Value = Range("GG8")
        Me.txt_CurrentPG8222.Value = Range("GG9")
        Me.txt_CurrentSS1.Value = Range("GG10")
        Me.txt_CurrentCRS2.Value = Range("GG11")
        Me.txt_CurrentCRS2P.Value = Range("GG12")
        Me.txt_CurrentEA1.Value = Range("GG13")
        Me.txt_CurrentCSS1_Undiluted.Value = Range("GG14")
        Me.txt_CurrentCSS1.Value = Range("GG15")
    End With
End Sub




