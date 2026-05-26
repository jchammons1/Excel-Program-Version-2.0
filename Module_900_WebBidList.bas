Attribute VB_Name = "Module_900_WebBidList"

Option Explicit
    Dim FORM_SELECTED As String
    
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'*********************************** This section contains Webbidlist that creates a document for the OSARC website *********************>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
    

 Sub WEB_BIDLIST()

    Application.ScreenUpdating = False
    WEBBIDLIST.Select
    WEBBIDLIST.Unprotect
    FORM_SELECTED = ActiveSheet.Name
    Range("A11:CZ999").Select
    Range("A11:CZ999").Clear                'Note: There is data in Range "AB1:AB3" used to populate this sheet
    Range("C5:C8").ClearContents
    Range("C9:I9").ClearContents
    Range("C10:I10").ClearContents
    
    WEBBIDLIST.Range("C5").Value = PROJECT_DATA.Range("D2").Value    'County
    WEBBIDLIST.Range("C6").Value = PROJECT_DATA.Range("M2").Value    'Copy Project No
    WEBBIDLIST.Range("C7").Value = PROJECT_DATA.Range("A2").Value   'Road Name #1
    WEBBIDLIST.Range("C8").Value = PROJECT_DATA.Range("B2").Value    'Road Name #2
    WEBBIDLIST.Range("C9").Value = BIDTAB.Range("B4").Value         'BID TAB  location
    WEBBIDLIST.Range("C10").Value = "=CONCATENATE(BIDTAB!R12C5,"" miles of "",BIDTAB!R10C2,"" "",BIDTAB!R11C1)"     ' Project Description
     
' Copying pay items and pasting to this starting location
    Range("B11").Select
    Call OrganizePayItems_900_WebBidList
    
    WEBBIDLIST.Select
    ActiveCell.Offset(0, 8).Range("A1").Select
    ActiveCell.FormulaR1C1 = "THE_END"
     
'   Retrieve Roadway Participating and Non-Particpating  from the Project Pay Item database sheet.  Copy to the 900 sheet
    Call Format_WebBidList
 
'Setting the print area to find the row with "TRANSFER FUNDS TO THE PROJECT'S ENGINEERING FUND" so that the page prints to the last row
  Dim ws As Worksheet
  Dim LastRow As Long
  Set ws = WEBBIDLIST
    LastRow = ws.Cells.Find("***THE_END***", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row      ' find the last row with formatting, to be included in print range
    ws.PageSetup.PrintArea = ws.Range("A1:I" & LastRow).Address

'Protecting and finishing up
    WEBBIDLIST.Select
    Rows("5:999").EntireRow.AutoFit
    WEBBIDLIST.Protect
    Range("A5").Select
End Sub


'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'*********************************** This section contains Form 900 that creates a document for CE to send to the newspaper *************>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################


 Sub Form_900()
    FORM_SELECTED = ActiveSheet.Name
    Range("A19:CZ999").Clear                 'Note: There is data in Range "AB1:AB3" used to populate this sheet
    Range("A10:I16").Clear
    
    Call Heading_900
    
    ' Copying pay items and pasting to this starting location
    Range("B19").Select
    Call OrganizePayItems_900_WebBidList
    
    Form900.Select
        ' Format columns
        Call Format_900
     
    ActiveCell.Offset(3, -6).Range("A1").Select
    ActiveCell.FormulaR1C1 = "PROJECT NO. "
        Selection.Font.Bold = True
        Selection.ShrinkToFit = True
        Selection.Font.Size = 12
    ActiveCell.Offset(0, 0).Range("A1:B1").Select
    Selection.Merge
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=PROJECT_DATA!R2C13"          ' Project No
        Selection.Font.Size = 12
        
    ActiveCell.Offset(1, -2).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=PROJECT_DATA!R2C4"
    ActiveCell.Offset(0, 0).Range("A1:B1").Select
    Selection.Merge
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "COUNTY"
        Selection.Font.Bold = True
        Selection.ShrinkToFit = True
        Selection.Font.Size = 12
    
    ActiveCell.Offset(2, -3).Range("A1").Select
    ActiveCell.FormulaR1C1 = "NOTICE TO CONTRACTORS:"
        Selection.Font.Underline = xlUnderlineStyleSingle
        Selection.Font.Bold = True
        Selection.Font.Size = 12
    
    ActiveCell.Offset(2, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "CONTRACT TIME:"
        Selection.Font.Bold = True
        Selection.Font.Size = 12
        Selection.Font.Underline = xlUnderlineStyleSingle
    ActiveCell.Offset(0, 1).Range("A1").Select          ' Working days data entry cell
    ActiveCell.FormulaR1C1 = "=PROJECT_DATA!R2C10"
    Selection.NumberFormat = "0"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Working Days"

    ActiveCell.Offset(2, -3).Range("A1").Select
    ActiveCell.FormulaR1C1 = "The award, if made, will be made to the lowest qualified bidder on the basis of published quantities."
    ActiveCell.Offset(0, 0).Range("A1:I1").Select
    Selection.Merge
        With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = True
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    
    ActiveCell.Offset(2, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "The Board of Supervisors hereby notifies all bidders that it will affirmatively ensure that in any contract entered into pursuant to this advertisement; minority business enterprise will be afforded full opportunity to submit bids in response to this invitation and will not be discriminated against on the grounds of race, color, or national origin in consideration for an award."
    ActiveCell.Offset(0, 0).Range("A1:I4").Select
    Selection.Merge
        With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
   
    ActiveCell.Offset(2, 0).Range("A1").Select
    ''ActiveCell.FormulaR1C1 = Range("AB2").Value                  replace this with the direct concatenate
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(""PLANS AND SPECIFICATIONS are on file in the Office of the Chancery Clerk of "",PROJECT_DATA!R2C4,"" County, "",PROJECT_DATA!R2C20,"" , the LSBP Engineer's office and the Office of the State Aid Engineer, 412 E. Woodrow Wilson Avenue, Jackson, Mississippi."","" This project shall be constructed in accordance with the latest edition of the Mississippi Standard Spec" & _
        "ifications for State Aid Road and Bridge Construction."")" & _
        ""
    ActiveCell.Offset(0, 0).Range("A1:I5").Select
    Selection.Merge
        With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 13434879
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    Selection.Locked = False
    
    '
    '
    '*******************if statement for notifying CE or reviewer to check for a cost for plans for >= 20,000 population counties.  Updated 5-1-24
    '
    '
    Dim PopulationLookup As String
    
    PopulationLookup = Application.VLookup(PROJECT_DATA.Range("D2").Value, DATA_VALIDATION.Range("Q2:S83"), 3, False)
    If PopulationLookup < 20000 Then
    
        ActiveCell.Offset(2, 0).Range("A1").Select
        ActiveCell.FormulaR1C1 = _
            "The plan can be purchased at this location, ***__________________________________________****, at a cost of **** $________________ ****** for the PLANS and a cost of **** $________________ ****** for the CONTRACT DOCUMENTS. The population of this county is less than 20,000."
        ActiveCell.Offset(0, 0).Range("A1:I3").Select
        Selection.Merge
            With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
            End With
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 15204327
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
            End With
            Selection.Locked = False
            Else
        End If
    
    
    ActiveCell.Offset(2, 0).Range("A1").Select              'Form 900 statement from the Project Data form and PROJECT_DATA sheet
    ActiveCell.FormulaR1C1 = "=PROJECT_DATA!R2C22"
    ActiveCell.Offset(0, 0).Range("A1:I5").Select
    Selection.Merge
        With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 13434879
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    Selection.Locked = False
    
    ActiveCell.Offset(2, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(""Certified check or bid bond for five percent (5%) of the total bid, made payable to "",PROJECT_DATA!R2C4,"" County and the State of Mississippi must accompany each proposal."")"
    ActiveCell.Offset(0, 0).Range("A1:I3").Select
    Selection.Merge
        With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    
    ActiveCell.Offset(2, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Bidders are hereby notified that any proposal accompanied by letters qualifying in any manner the condition under which the proposal is tendered will be considered an irregular bid and such proposal will not be considered in making the award."
    ActiveCell.Offset(0, 0).Range("A1:I4").Select
    Selection.Merge
        With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
  
    ActiveCell.Offset(3, 0).Range("A1").Select
    ActiveCell.Offset(0, 5).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(PROJECT_DATA!R2C28,"", President"")"
    
    ActiveCell.Offset(0, 0).Range("A1:D1").Select
    Selection.Merge
        With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = True
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 13434879
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    Selection.Locked = False
    
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(PROJECT_DATA!R2C4,"" County Board of Supervisors"")"
    ActiveCell.Offset(0, 0).Range("A1:D1").Select
    Selection.Merge
        With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = True
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.Locked = False
    
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "THE_END"
    

 
'Setting the print area to find the row with "The End" so that the page prints to the last row
  Dim ws As Worksheet
  Dim LastRow As Long
  Set ws = Form900
    LastRow = ws.Cells.Find("***THE_END***", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row      ' find the last row with formatting, to be included in print range
    ws.PageSetup.PrintArea = ws.Range("A1:I" & LastRow).Address

'Protecting and finishing up
    PAYITEMTYPE.Visible = False
    Form900.Select
    Rows("10:999").EntireRow.AutoFit
    Call Format_PageSetup
    Form900.Protect
End Sub



 Sub OrganizePayItems_900_WebBidList()
    PAYITEMTYPE.Visible = True
    PAYITEMTYPE.Select
    Dim RowType As Range                         ' Declare a variable for the Pay Item Type and the range of data from the Data Validation sheet
    For Each RowType In Range("A2:A55")          ' Maximum of 50 types the user can add to the program.  4 default types
       
        If IsEmpty(RowType) = False Then
        
    ' Select pay items on Database
    Database.Visible = True
    Database.Select
    Database.Range("$A$1:$CS$999").AutoFilter
    Range("A2").Select
    
' Filter pay items to type'
    ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=46, Criteria1:=RowType.Value
    Range("A2").Select
    Selection.End(xlDown).Select
            If Selection <> "" Then
                Range("B2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
            ' Pasting copied participating and non-participating pay items
                Sheets(FORM_SELECTED).Select
                ActiveCell.FormulaR1C1 = RowType.Value & " ITEMS:"
                ActiveCell.Offset(0, 6).Range("A1").Value = "QUANTITY"
                ActiveCell.Offset(0, 7).Range("A1").Value = "UNIT"
                ActiveCell.Offset(1, -1).Range("A1").Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            ' Return to the database and copy the quantity and units
                Database.Select
                Range("C2:D2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
            ' Pasting copied Participating Pay Items to the Engineer's Estimate'
                Sheets(FORM_SELECTED).Select
                ActiveCell.Offset(0, 7).Range("A1").Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                    If Selection.End(xlDown).Value <> "" Then
                        Selection.End(xlDown).Select
                    Else
                    End If
                ActiveCell.Offset(2, -6).Range("A1").Select
                End If
        End If
    Next RowType

    ' Project number to right header
        With Sheets(FORM_SELECTED)
            Call Format_PageSetup
        End With
        
    Database.Range("$A$1:$CZ$999").AutoFilter
End Sub






 Sub Form_900Alternate()
    FORM_SELECTED = ActiveSheet.Name
    Range("A19:CZ999").Clear                 'Note: There is data in Range "AB1:AB3" used to populate this sheet
    Range("A10:I16").Clear
    
    Call Heading_900
    
    ' Copying pay items and pasting to this starting location
    Range("B19").Select
    Call OrganizePayItems_900Alternate
    
    'Determine if there are any Non-Participating Pay Items.  If not, then skip the Call
            Database.Select
            Database.Range("$A$1:$CS$999").AutoFilter
                ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=47, Criteria1:="No"
                Range("A2").Select
                Selection.End(xlDown).Select
            If Selection <> "" Then
                Form902.Select
                ActiveCell.Offset(1, 0).Range("A1").Select
                Call OrganizePayItems_900NP_Alternate
                Database.Range("$A$1:$CS$999").AutoFilter
                Else
                Form902.Select
                Database.Range("$A$1:$CS$999").AutoFilter
            End If
            
    Form900.Select
        ' Format columns
        Call Format_900
     
    ActiveCell.Offset(3, -6).Range("A1").Select
    ActiveCell.FormulaR1C1 = "PROJECT NO. "
        Selection.Font.Bold = True
        Selection.ShrinkToFit = True
        Selection.Font.Size = 12
    ActiveCell.Offset(0, 0).Range("A1:B1").Select
    Selection.Merge
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=PROJECT_DATA!R2C13"          ' Project No
        Selection.Font.Size = 12
        
    ActiveCell.Offset(1, -2).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=PROJECT_DATA!R2C4"
    ActiveCell.Offset(0, 0).Range("A1:B1").Select
    Selection.Merge
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "COUNTY"
        Selection.Font.Bold = True
        Selection.ShrinkToFit = True
        Selection.Font.Size = 12
    
    ActiveCell.Offset(2, -3).Range("A1").Select
    ActiveCell.FormulaR1C1 = "NOTICE TO CONTRACTORS:"
        Selection.Font.Underline = xlUnderlineStyleSingle
        Selection.Font.Bold = True
        Selection.Font.Size = 12
    
    ActiveCell.Offset(2, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "CONTRACT TIME:"
        Selection.Font.Bold = True
        Selection.Font.Size = 12
        Selection.Font.Underline = xlUnderlineStyleSingle
    ActiveCell.Offset(0, 1).Range("A1").Select          ' Working days data entry cell
    ActiveCell.FormulaR1C1 = "=PROJECT_DATA!R2C10"
    Selection.NumberFormat = "0"
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Working Days"

    ActiveCell.Offset(2, -3).Range("A1").Select
    ActiveCell.FormulaR1C1 = "The award, if made, will be made to the lowest qualified bidder on the basis of published quantities."
    ActiveCell.Offset(0, 0).Range("A1:I1").Select
    Selection.Merge
        With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = True
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    
    ActiveCell.Offset(2, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "The Board of Supervisors hereby notifies all bidders that it will affirmatively ensure that in any contract entered into pursuant to this advertisement; minority business enterprise will be afforded full opportunity to submit bids in response to this invitation and will not be discriminated against on the grounds of race, color, or national origin in consideration for an award."
    ActiveCell.Offset(0, 0).Range("A1:I4").Select
    Selection.Merge
        With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
   
    ActiveCell.Offset(2, 0).Range("A1").Select
    ''ActiveCell.FormulaR1C1 = Range("AB2").Value                  replace this with the direct concatenate
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(""PLANS AND SPECIFICATIONS are on file in the Office of the Chancery Clerk of "",PROJECT_DATA!R2C4,"" County, "",PROJECT_DATA!R2C20,"" , the LSBP Engineer's office and the Office of the State Aid Engineer, 412 E. Woodrow Wilson Avenue, Jackson, Mississippi."","" This project shall be constructed in accordance with the latest edition of the Mississippi Standard Spec" & _
        "ifications for State Aid Road and Bridge Construction."")" & _
        ""
    ActiveCell.Offset(0, 0).Range("A1:I5").Select
    Selection.Merge
        With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 13434879
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    Selection.Locked = False
    
    '
    '
    '*******************if statement for notifying CE or reviewer to check for a cost for plans for >= 20,000 population counties.  Updated 5-1-24
    '
    '
    Dim PopulationLookup As String
    
    PopulationLookup = Application.VLookup(PROJECT_DATA.Range("D2").Value, DATA_VALIDATION.Range("Q2:S83"), 3, False)
    If PopulationLookup < 20000 Then
    
        ActiveCell.Offset(2, 0).Range("A1").Select
        ActiveCell.FormulaR1C1 = _
            "The plan can be purchased at this location, ***__________________________________________****, at a cost of **** $________________ ****** for the PLANS and a cost of **** $________________ ****** for the CONTRACT DOCUMENTS. The population of this county is less than 20,000."
        ActiveCell.Offset(0, 0).Range("A1:I3").Select
        Selection.Merge
            With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
            End With
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 15204327
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
            End With
            Selection.Locked = False
            Else
        End If
    
    
    ActiveCell.Offset(2, 0).Range("A1").Select              'Form 900 statement from the Project Data form and PROJECT_DATA sheet
    ActiveCell.FormulaR1C1 = "=PROJECT_DATA!R2C22"
    ActiveCell.Offset(0, 0).Range("A1:I5").Select
    Selection.Merge
        With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 13434879
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    Selection.Locked = False
    
    ActiveCell.Offset(2, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(""Certified check or bid bond for five percent (5%) of the total bid, made payable to "",PROJECT_DATA!R2C4,"" County and the State of Mississippi must accompany each proposal."")"
    ActiveCell.Offset(0, 0).Range("A1:I3").Select
    Selection.Merge
        With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    
    ActiveCell.Offset(2, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Bidders are hereby notified that any proposal accompanied by letters qualifying in any manner the condition under which the proposal is tendered will be considered an irregular bid and such proposal will not be considered in making the award."
    ActiveCell.Offset(0, 0).Range("A1:I4").Select
    Selection.Merge
        With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
  
    ActiveCell.Offset(3, 0).Range("A1").Select
    ActiveCell.Offset(0, 5).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(PROJECT_DATA!R2C28,"", President"")"
    
    ActiveCell.Offset(0, 0).Range("A1:D1").Select
    Selection.Merge
        With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = True
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 13434879
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    Selection.Locked = False
    
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(PROJECT_DATA!R2C4,"" County Board of Supervisors"")"
    ActiveCell.Offset(0, 0).Range("A1:D1").Select
    Selection.Merge
        With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = True
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.Locked = False
    
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "THE_END"
    

 
'Setting the print area to find the row with "The End" so that the page prints to the last row
  Dim ws As Worksheet
  Dim LastRow As Long
  Set ws = Form900
    LastRow = ws.Cells.Find("***THE_END***", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row      ' find the last row with formatting, to be included in print range
    ws.PageSetup.PrintArea = ws.Range("A1:I" & LastRow).Address

'Protecting and finishing up
    PAYITEMTYPE.Visible = False
    Form900.Select
    Rows("10:999").EntireRow.AutoFit
    Call Format_PageSetup
    Form900.Protect
End Sub



 Sub OrganizePayItems_900Alternate()
    PAYITEMTYPE.Visible = True
    PAYITEMTYPE.Select
    Dim RowType As Range                         ' Declare a variable for the Pay Item Type and the range of data from the Data Validation sheet
    For Each RowType In Range("A2:A55")          ' Maximum of 50 types the user can add to the program.  4 default types
       
        If IsEmpty(RowType) = False Then
        
    ' Select pay items on Database
    Database.Visible = True
    Database.Select
    Database.Range("$A$1:$CS$999").AutoFilter
    Range("A2").Select
    
' Filter pay items to Participating Items'
    ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=47, Criteria1:="Yes"
' Filter pay items to type'
    ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=46, Criteria1:=RowType.Value
    Range("A2").Select
    Selection.End(xlDown).Select
            If Selection <> "" Then
                Range("B2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
            ' Pasting copied participating and non-participating pay items
                Sheets(FORM_SELECTED).Select
                ActiveCell.FormulaR1C1 = RowType.Value & " PARTICIPATING ITEMS:"
                ActiveCell.Offset(1, -1).Range("A1").Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            ' Return to the database and copy the quantity and units
                Database.Select
                Range("C2:D2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
            ' Pasting copied Participating Pay Items to the Engineer's Estimate'
                Sheets(FORM_SELECTED).Select
                ActiveCell.Offset(0, 7).Range("A1").Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                    If Selection.End(xlDown).Value <> "" Then
                        Selection.End(xlDown).Select
                    Else
                    End If
                ActiveCell.Offset(2, -6).Range("A1").Select
                End If
        End If
    Next RowType

    ' Project number to right header
        With Sheets(FORM_SELECTED)
            Call Format_PageSetup
        End With
        
    Database.Range("$A$1:$CZ$999").AutoFilter
End Sub


Sub OrganizePayItems_900NP_Alternate()
    PAYITEMTYPE.Visible = True
    PAYITEMTYPE.Select
    Dim RowType As Range                         ' Declare a variable for the Pay Item Type and the range of data from the Data Validation sheet
    For Each RowType In Range("A2:A55")          ' Maximum of 50 types the user can add to the program.  4 default types
       
        If IsEmpty(RowType) = False Then
        
    ' Select pay items on Database
    Database.Visible = True
    Database.Select
    Database.Range("$A$1:$CS$999").AutoFilter
    Range("A2").Select
    
' Filter pay items to Participating Items'
    ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=47, Criteria1:="No"
' Filter pay items to type'
    ActiveSheet.Range("$A$1:$CS$999").AutoFilter Field:=46, Criteria1:=RowType.Value
    Range("A2").Select
    Selection.End(xlDown).Select
            If Selection <> "" Then
                Range("B2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
            ' Pasting copied participating and non-participating pay items
                Sheets(FORM_SELECTED).Select
                ActiveCell.FormulaR1C1 = RowType.Value & " NON-PARTICIPATING ITEMS:"
                ActiveCell.Offset(1, -1).Range("A1").Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            ' Return to the database and copy the quantity and units
                Database.Select
                Range("C2:D2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
            ' Pasting copied Participating Pay Items to the Engineer's Estimate'
                Sheets(FORM_SELECTED).Select
                ActiveCell.Offset(0, 7).Range("A1").Select
                     Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                    If Selection.End(xlDown).Value <> "" Then
                        Selection.End(xlDown).Select
                    Else
                    End If
                ActiveCell.Offset(2, -6).Range("A1").Select
                End If
        End If
    Next RowType

    ' Project number to right header
        With Sheets(FORM_SELECTED)
            Call Format_PageSetup
        End With
        
    Database.Range("$A$1:$CZ$999").AutoFilter
End Sub

Sub Heading_900()
' Cell Range "AB1" concatenate data from the PROJECT_DATA sheet
    Range("AB1").Select
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(""Sealed bids will be received by the Board of Supervisors of "",PROJECT_DATA!R2C4,"" County"","" Mississippi at the "","" ***LOCATION *** "","" ***CITY *** "","" Mississippi, "",""until "","" ***??:??_AM/PM *** "",""on the "","" ***DAY*** "",""day of "","" ***MONTH*** "","" ***YEAR*** "",""and shortly thereafter publicly"","" opened for the construction of " & _
        """,TEXT(PROJECT_DATA!R2C17,""0.000""),"" miles of "","" "",PROJECT_DATA!R2C5,"" "","" on the "",PROJECT_DATA!R2C1,"" "",PROJECT_DATA!R2C2,"" being known as "",PROJECT_DATA!R2C12,"" Project No. "",PROJECT_DATA!R2C13,"" in "",PROJECT_DATA!R2C4,"" County"","" Mississippi."")" & _
        ""
    Selection.Copy
    Range("A10").Select
          Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A10:I14").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = True
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
            With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 13434879
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
    Selection.Locked = False
    
    ' Heading
    Range("A17").Select
        ActiveCell.FormulaR1C1 = "PRINCIPAL ITEMS OF WORK ARE APPROXIMATELY AS FOLLOWS"
    Range("A17:I17").Select
        Selection.Merge
        Selection.Font.Size = 12
        Selection.Font.Bold = True
        Selection.Font.Underline = xlUnderlineStyleSingle
        Selection.HorizontalAlignment = xlCenter
End Sub
