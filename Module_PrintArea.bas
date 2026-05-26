Attribute VB_Name = "Module_PrintArea"
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'*********************************** This section sets the print area for ??? document types       **************************************>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################


 Sub SetPrintArea_Estimate()
   
'Setting the print area to find the last row of data for the Estimates
    Dim ws As Worksheet
    Dim lastRow As Long
    Set ws = ActiveSheet
    Dim EstimateNameSelected As String
    EstimateNameSelected = ActiveSheet.Name

            Select Case EstimateNameSelected        'Evaluate Sheet Name
                Case "ENGINEER", "OSARC", "CONTRACTOR"
                    lastRow = ws.Cells.Find("PROJECT PARTICIPATING AND NON-PARTICIPATING TOTAL", SearchOrder:=xlByRows, SearchDirection:=xlNext).row
                    ws.PageSetup.PrintArea = ws.Range("A1:F" & lastRow).Address
                Case "BIDTAB"
                    lastRow = ws.Cells.Find("APPROVED:", SearchOrder:=xlByRows, SearchDirection:=xlNext).row
                    ws.PageSetup.PrintArea = ws.Range("A1:DF" & lastRow).Address
                Case "CAD_Template"
                    lastRow = ws.Cells.Find("End of Fuel Adjustment Summary", SearchOrder:=xlByRows, SearchDirection:=xlNext).row
                    ws.PageSetup.PrintArea = ws.Range("A1:L" & lastRow).Address
            End Select
            
    Call Format_PageSetup
End Sub


 Sub SetPrintArea_CalculateCAD()
   
'Setting the print area to find the last row of data for the Estimates
    Dim ws As Worksheet
    Dim lastRow As Long
    Set ws = ActiveSheet
    Dim EstimateNameSelected As String
    EstimateNameSelected = ActiveSheet.Name

    ActiveSheet.Range("A53").Select
    
            Select Case EstimateNameSelected        'Evaluate Sheet Name
                Case EstimateNameSelected
                    lastRow = ws.Cells.Find("End of Monthly Items Summary", SearchOrder:=xlByRows, SearchDirection:=xlNext).row
                    ws.PageSetup.PrintArea = ws.Range("A1:L" & lastRow).Address
            End Select
            
End Sub


