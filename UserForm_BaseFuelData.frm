VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_BaseFuelData 
   Caption         =   "BASE FUEL ADJUSTMENT Data Form"
   ClientHeight    =   13560
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   10200
   OleObjectBlob   =   "UserForm_BaseFuelData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_BaseFuelData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'
'
'############################################### BASE FUEL ADJUSTMENT ####################################################
'
'

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

Private Sub Button_ImportBaseFuel_Click()

    Dim MonthSelected As String
    Dim YearSelected As Long

    MonthSelected = cmb_BaseMonth.Value
    YearSelected = CLng(cmb_BaseYear.Value)

    Dim data As Variant
    Dim matchRow As Long

    On Error GoTo ErrHandler

    'Pull latest workbook from SharePoint (Anyone link) and read to array
    data = LoadFuelPricesData()

    'Find the matching row
    matchRow = FindFuelRow(data, MonthSelected, YearSelected)

    If matchRow > 0 Then
        txt_Gasoline.Value = ValueByHeader(data, matchRow, "Gasoline")
        txt_Diesel.Value = ValueByHeader(data, matchRow, "Diesel")
        txt_PG6422.Value = ValueByHeader(data, matchRow, "PG6422")
        txt_PG6722.Value = ValueByHeader(data, matchRow, "PG6722")
        txt_PG7622.Value = ValueByHeader(data, matchRow, "PG7622")
        txt_PG8222.Value = ValueByHeader(data, matchRow, "PG8222")
        txt_SS1.Value = ValueByHeader(data, matchRow, "SS1")
        txt_CRS2.Value = ValueByHeader(data, matchRow, "RS2C_CRS2")
        txt_CRS2P.Value = ValueByHeader(data, matchRow, "CRS2P")
        txt_EA1.Value = ValueByHeader(data, matchRow, "EA1_ERP1_AEP")
        txt_CSS1_Undiluted.Value = ValueByHeader(data, matchRow, "CSS1_Undiluted")
        txt_CSS1.Value = ValueByHeader(data, matchRow, "CSS1_1H")
    Else
        MsgBox "The current month base fuel data selected is not currently available. Check back later.", vbExclamation
    End If

    Exit Sub

ErrHandler:
    MsgBox "Fuel import failed: " & Err.Description, vbCritical
End Sub



Private Sub Button_SaveBaseFuelAdjustment_Click()
'######################################################### FUEL ADJUSTMENT CALCULATIONS ##############################################
' Set the message.  After changin the below Do while code to Select Case for the fuel code, the processing message is not needed because it is a lot faster
    UserForm_BaseFuelData.lblMessage.Caption = "Processing, please wait..."
    DoEvents
    
    Call DownloadWAVFile

' Adding Base Fuel Adjustment Data to the initial CAD sheet
Application.ScreenUpdating = False
    CAD.Visible = True
    CAD.Select
    CAD.Unprotect
   
If Me.cmb_BaseMonth.Value = "" Or Me.cmb_BaseYear.Value = "" Or Me.txt_Gasoline.Value = "" Or Me.txt_Diesel.Value = "" Or _
    Me.txt_PG6422.Value = "" Or Me.txt_PG6722.Value = "" Or Me.txt_PG7622.Value = "" Or Me.txt_PG8222.Value = "" Or Me.txt_SS1.Value = "" _
    Or Me.txt_CRS2.Value = "" Or Me.txt_CRS2P.Value = "" Or Me.txt_EA1.Value = "" Or Me.txt_CSS1_Undiluted.Value = "" _
    Or Me.txt_CSS1.Value = "" Then

MsgBox "Please ensure all fields are populated or contain a value of 0.  User results did not save."
Exit Sub
End If

' This data saved on the CAD form will not be directly used for calculation.  instead this will be used as a reference to what is updated in the Database.
' Serves as the values shown on the Base Fuel Data user form for the user to know the base fuel values.
    With Me
            Range("GF2") = Me.cmb_BaseMonth.Value
            Range("GF3") = Me.cmb_BaseYear.Value
            Range("GF4") = Me.txt_Gasoline.Value
            Range("GF5") = Me.txt_Diesel.Value
            Range("GF6") = Me.txt_PG6422.Value
            Range("GF7") = Me.txt_PG6722.Value
            Range("GF8") = Me.txt_PG7622.Value
            Range("GF9") = Me.txt_PG8222.Value
            Range("GF10") = Me.txt_SS1.Value
            Range("GF11") = Me.txt_CRS2.Value
            Range("GF12") = Me.txt_CRS2P.Value
            Range("GF13") = Me.txt_EA1.Value
            Range("GF14") = Me.txt_CSS1_Undiluted.Value
            Range("GF15") = Me.txt_CSS1.Value
    End With
 
' Adds the base fuel adjustment data to the Database
' The CAD template must be updated after this data is added or updated
    Database.Visible = True
    Database.Select
    Database.Unprotect
    
' Since I am using Option Explicit at the Top of this code.  The variables now must be defined for row, column, endrow,
    Dim row As Long
    Dim column As Long
    Dim endrow As Long

' Start on 2nd row.  Populate the rows with "NA" that exist
            row = 2
            column = 1
            endrow = CountPayItems - 46
    Do While row < endrow
    
    Dim FuelCode As String
    FuelCode = Cells(row, column + 7).Value
    
    Select Case Cells(row, column + 7).Value
          Case "E", "GY", "GT", "M", "B", "D", "C", "S", "BA1", "A1", "A2", "A3", "A4", "A5", "A6"
    
        'If Cells(row, column + 7).Value = "E" Or Cells(row, column + 7).Value = "GY" Or Cells(row, column + 7).Value = "GT" Or Cells(row, column + 7).Value = "M" _
        'Or Cells(row, column + 7).Value = "B" Or Cells(row, column + 7).Value = "D" Or Cells(row, column + 7).Value = "C" Or Cells(row, column + 7).Value = "S" _
       ' Or Cells(row, column + 7).Value = "BA1" Or Cells(row, column + 7).Value = "A1" Or Cells(row, column + 7).Value = "A2" Or Cells(row, column + 7).Value = "A3" _
        'Or Cells(row, column + 7).Value = "A4" Or Cells(row, column + 7).Value = "A5" Or Cells(row, column + 7).Value = "A6" Then
     
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
     '   End If
     
       End Select
    ' Show which row is processing
        UserForm_BaseFuelData.lblMessage.Caption = "Processing: " & row & " of " & endrow
        DoEvents
    row = row + 1
    Loop

    'Clear caption
    UserForm_MonthlyFuelData.lblMessage.Caption = ""
    DoEvents
    
    Database.Select
    Database.Range("A2").Select
    Database.Range("$A$1:$CS$999").AutoFilter
    Database.Protect
    Database.Visible = False
    CAD.Select
    CAD.Range("GF4:GG15").Select
    Selection.NumberFormat = "0.0000"
    CAD.Protect
    CAD.Visible = False
    Application.ScreenUpdating = True
    
    Unload Me
End Sub

 Sub UserForm_Initialize()
    CAD.Visible = True
    CAD.Select
    CAD.Unprotect
    
    With Me     'Load Values from the BASE FUEL DATA located on the CAD template in Column GF
        Me.cmb_BaseMonth.Value = Range("GF2")
        Me.cmb_BaseYear.Value = Range("GF3")
        Me.txt_Gasoline.Value = Range("GF4")
        Me.txt_Diesel.Value = Range("GF5")
        Me.txt_PG6422.Value = Range("GF6")
        Me.txt_PG6722.Value = Range("GF7")
        Me.txt_PG7622.Value = Range("GF8")
        Me.txt_PG8222.Value = Range("GF9")
        Me.txt_SS1.Value = Range("GF10")
        Me.txt_CRS2.Value = Range("GF11")
        Me.txt_CRS2P.Value = Range("GF12")
        Me.txt_EA1.Value = Range("GF13")
        Me.txt_CSS1_Undiluted.Value = Range("GF14")
        Me.txt_CSS1.Value = Range("GF15")
    End With
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Check if the form is being closed by the user (CloseMode = 0)
    If CloseMode = vbFormControlMenu Then
        ' Lock the sheet
        CAD.Protect Password:="roadway123"
            ' Hide the sheet.  When you X close the CAD stays visible if this is not here.
        CAD.Visible = xlSheetHidden
    End If
End Sub
