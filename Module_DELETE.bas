Attribute VB_Name = "Module_DELETE"
  Option Explicit
    Dim ESTIMATE_Deleted As String '*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'*********************************** This section contains all the delete buttons *******************************************************>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
Sub Delete_PayItems()
    ' Delete pay item from the Database sheet
    If MsgBox("Do you want to delete this all the pay item?", vbYesNo, "Save Changes") = vbYes Then
        Database.Unprotect
        Database.Visible = True
        Database.Range("A2:ACS9999").Clear
        Database.Protect
        Database.Visible = False
        
        AvgUnitCost.Unprotect
        AvgUnitCost.Visible = True
        AvgUnitCost.Range("A2:CS9999").Clear
        AvgUnitCost.Protect
        AvgUnitCost.Visible = False
    End If
End Sub
Sub Delete_Estimates()
    ESTIMATE_Deleted = ActiveSheet.Name
    
    If MsgBox("Do you want to delete this estimate sheets data?", vbYesNo, "Save Changes") = vbYes Then
        Sheets(ESTIMATE_Deleted).Select
        Sheets(ESTIMATE_Deleted).Unprotect
        Sheets(ESTIMATE_Deleted).Range("A21:CZ9999").Clear
        Sheets(ESTIMATE_Deleted).Protect
        Sheets(ESTIMATE_Deleted).Select
    End If
End Sub

Sub Delete_ProjectData()
    If MsgBox("Do you want to delete the project data?", vbYesNo, "Save Changes") = vbYes Then
        PROJECT_DATA.Unprotect
        PROJECT_DATA.Visible = True
        PROJECT_DATA.Select
        PROJECT_DATA.Range("A2:AN5").ClearContents
            With PROJECT_DATA
            'LEFT Side of the Project Data Form
                Range("K2").Value = 12#
            'RIGHT side of the Project Data Form
            'BOTTOM of the Project Data Form
                Range("V2").Value = "If applicable, please include additional Form 900 statement here"
            'Default non-user added data to the ProjectData table / sheet
            End With
        PROJECT_DATA.Protect
        PROJECT_DATA.Visible = False
        START.Select
    End If
End Sub
 
Sub Delete_900()
    If MsgBox("Do you want to clear the 900 contents?", vbYesNo, "DELETE 900 Sheet Data") = vbYes Then
        Form900.Select
        Form900.Unprotect
        Range("A19:CZ999").Clear
        Range("A9:I16").Clear
        Form900.Protect
        Range("A10").Select
    End If
End Sub

Sub Delete_902()
    If MsgBox("Do you want to clear the 902 contents?", vbYesNo, "DELETE 902 Sheet Data") = vbYes Then
        Application.ScreenUpdating = False
        Form902.Select
        Form902.Unprotect Password:="roadway123"
        Range("A13:CZ999").Clear
        Range("A4:F7").Clear
        Range("E3:G3").ClearContents
        Form902.Protect Password:="roadway123"
        Application.ScreenUpdating = True
        Range("A11").Select
    End If
End Sub

Sub Delete_GreenCover()
    GreenCover.Select
    GreenCover.Unprotect
'Left Side
    GreenCover.Range("D15:G15").ClearContents
    GreenCover.Range("C17").ClearContents
    GreenCover.Range("C19").ClearContents
    GreenCover.Range("E17").ClearContents
    GreenCover.Range("E19").ClearContents
    GreenCover.Range("G17:H17").ClearContents
    GreenCover.Range("G19:H19").ClearContents
    GreenCover.Range("J17").ClearContents
    GreenCover.Range("J19").ClearContents
    GreenCover.Range("I21").ClearContents
    GreenCover.Range("D21:G21").ClearContents
'Right Side
    GreenCover.Range("P8:T8").ClearContents
    GreenCover.Range("N17:P17").ClearContents
    GreenCover.Range("Q17").ClearContents
    GreenCover.Range("Q18:V18").ClearContents
    GreenCover.Range("P19:V19").ClearContents
    GreenCover.Range("O20:Q20").ClearContents
    GreenCover.Range("S20:V20").ClearContents
    GreenCover.Range("P21:V21").ClearContents
    GreenCover.Range("N22:V22").ClearContents
    GreenCover.Range("P23").ClearContents
    GreenCover.Range("P24").ClearContents
    GreenCover.Range("U24").ClearContents
    GreenCover.Range("O25:V25").ClearContents
    GreenCover.Range("O26:V26").ClearContents
    GreenCover.Range("O27:V27").ClearContents
    GreenCover.Protect
    
    GreenCoverNEW.Select
    GreenCoverNEW.Unprotect
'Top
    GreenCover.Range("E16:I16").ClearContents
    GreenCover.Range("C25").ClearContents
    GreenCover.Range("F25").ClearContents
    GreenCover.Range("F26:K26").ClearContents
    GreenCover.Range("E27:K27").ClearContents
    GreenCover.Range("H28:K28").ClearContents
    GreenCover.Range("E29:K29").ClearContents
    GreenCover.Range("C30:K30").ClearContents
    GreenCover.Range("E31").ClearContents
    GreenCover.Range("E31").ClearContents
    GreenCover.Range("J32").ClearContents
    GreenCover.Range("D33:K33").ClearContents
    GreenCover.Range("D34:K34").ClearContents
    GreenCover.Range("D35:K35").ClearContents
'Bottom
    GreenCover.Range("E59:H59").ClearContents
    GreenCover.Range("C61:D61").ClearContents
    GreenCover.Range("C63:D63").ClearContents
    GreenCover.Range("H61:I61").ClearContents
    GreenCover.Range("H63:D63").ClearContents
    GreenCover.Range("E65:H65").ClearContents
    GreenCover.Range("F61").ClearContents
    GreenCover.Range("F63").ClearContents
    GreenCover.Range("K61").ClearContents
    GreenCover.Range("K63").ClearContents
    GreenCover.Range("J65").ClearContents
    GreenCoverNEW.Protect
    GreenCover.Protect
End Sub



Sub Delete_WebBidList()
    Application.ScreenUpdating = False
    WEBBIDLIST.Select
    WEBBIDLIST.Unprotect
    Range("A11:CZ9999").Clear
    Range("C5:C8").ClearContents
    Range("C9:I9").ClearContents
    Range("C10:I10").ClearContents
    WEBBIDLIST.Protect
    Application.ScreenUpdating = True
    Range("C5").Select
End Sub

Sub Delete_EngineersPayment()
    If MsgBox("Do you want to clear data on the Engineer's Payment Invoice Template (EngPayInv) sheet? ", vbYesNo, "Save changes") = vbYes Then
        Application.ScreenUpdating = False
        EngPayInv.Visible = True
        EngPayInv.Unprotect
        EngPayInv.Select
        Range("B5").ClearContents
        Range("D5:G5").ClearContents
        Range("B7").ClearContents
        Range("B8").ClearContents
        Range("B9").ClearContents
        Range("D9").ClearContents
        Range("D11").ClearContents
        Range("G11").ClearContents
        Range("D12").ClearContents
        Range("A18").ClearContents
        Range("C18:F18").ClearContents
        Range("B20").ClearContents
        Range("B21").ClearContents
        Range("B31").ClearContents
        Range("B31").ClearContents
        Range("B32").ClearContents
        Range("B34").ClearContents
        Range("G34").ClearContents
        EngPayInv.Protect
        EngPayInv.Visible = False
        
        EngPayInvFedFund.Visible = True
        EngPayInvFedFund.Unprotect
        EngPayInvFedFund.Select
        Range("B5").ClearContents
        Range("D5:G5").ClearContents
        Range("B7").ClearContents
        Range("B8").ClearContents
        Range("B9").ClearContents
        Range("D9").ClearContents
        Range("D11").ClearContents
        Range("G11").ClearContents
        Range("D12").ClearContents
        Range("A18").ClearContents
        Range("C18:F18").ClearContents
        Range("B20").ClearContents
        Range("B21").ClearContents
        Range("B31").ClearContents
        Range("B31").ClearContents
        Range("B32").ClearContents
        Range("B34").ClearContents
        Range("G34").ClearContents
        EngPayInvFedFund.Protect
        EngPayInvFedFund.Visible = False
        Application.ScreenUpdating = True
        
     End If
     START.Select
End Sub

Sub Delete_PayItemTypes()
    PAYITEMTYPE.Unprotect
    PAYITEMTYPE.Visible = True
    PAYITEMTYPE.Select
    PAYITEMTYPE.Range("A6:A55").ClearContents
    PAYITEMTYPE.Protect
    PAYITEMTYPE.Visible = False
    START.Select
End Sub

Sub Delete_PayItemTypesSupplemental()
    PAYITEMTYPE_SUPPLEMENTAL.Unprotect
    PAYITEMTYPE_SUPPLEMENTAL.Visible = True
    PAYITEMTYPE_SUPPLEMENTAL.Select
    PAYITEMTYPE_SUPPLEMENTAL.Range("A2:A51").ClearContents
    PAYITEMTYPE_SUPPLEMENTAL.Protect
    PAYITEMTYPE_SUPPLEMENTAL.Visible = False
    START.Select
End Sub

Sub Delete_LettingResults()
    If MsgBox("Do you want to clear the contents Letting Results and Bid Tab Letting Results Sheet?", vbYesNo, "DELETE Letting Results Data") = vbYes Then
        Application.ScreenUpdating = False
        LettingResults.Select
        LettingResults.Unprotect
        Range("C4:C6").Clear
        Range("B9:D999").Clear
        Range("J1:L999").Clear
    
        BIDTAB.Select
        BIDTAB.Unprotect
        Range("DB3:DD15").ClearContents
        Range("DF12:DF13").ClearContents
        Range("DA21:DZ9999").Clear
        
        BIDTAB.Protect
        LettingResults.Protect
        Application.ScreenUpdating = True
        Range("C4").Select
    Else
    End If
End Sub


Sub Delete_CAD()
' Clears the CAD and fuel adjustment on the template
    If MsgBox("Do you want to delete the CAD template data and fuel adjustment data?", vbYesNo, "DELETE CAD Data") = vbYes Then
        Application.ScreenUpdating = False
        CAD.Visible = True
        CAD.Select
        CAD.Unprotect Password:="roadway123"
            Range("A53:CV9999").Clear
            Range("GF2") = "Select Month"
            Range("GF3") = "Select Year"
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
            Range("GG2") = "Select Month"
            Range("GG3") = "Select Year"
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
          
            Range("EZ2").ClearContents
            Range("GY2").ClearContents
            Range("HA2").ClearContents
            
            Range("G30").ClearContents
            
            Range("B3:D3").ClearContents
            Range("B5:H5").ClearContents
            Range("G6:K6").ClearContents
            Range("B6:D6").ClearContents
        
        CAD.Select
        CAD.Protect Password:="roadway123"
        CAD.Visible = False
        Application.ScreenUpdating = False
    End If
End Sub


Sub Delete_MaterialsReport()
    If MsgBox("Do you want to delete the Materials Report?", vbYesNo, "DELETE Letting Results Data") = vbYes Then
        Material_Report.Unprotect
        Material_Report.Select
        Material_Report.Range("A7:H9999").Clear
        Material_Report.Protect
    End If
End Sub

