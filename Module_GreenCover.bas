Attribute VB_Name = "Module_GreenCover"
 Sub Form_GreenCover()

' Copying the project_data values to the Green Cover Sheet'
    GreenCover.Range("T8").Value = PROJECT_DATA.Range("D2").Value                ' Copy County Name
    GreenCover.Range("T8").Select
        Call Format_GreenCover2
    GreenCover.Range("Q17").Value = "PROJECT NO " & PROJECT_DATA.Range("M2").Value               ' Copy Project Number
    GreenCover.Range("Q17").Select
            Selection.Font.Bold = True
            Selection.Font.Size = 16
    GreenCover.Range("G15").Value = PROJECT_DATA.Range("D2").Value
    GreenCover.Range("G15").Select
        Call Format_GreenCover2
    GreenCover.Range("O20").Value = PROJECT_DATA.Range("D2").Value
    GreenCover.Range("O20").Select
    
    Range("P8:T8").Select
        Call Format_GreenCover
        Call Format_GreenCover2
    Range("D15:G15").Select
        Call Format_GreenCover
        Call Format_GreenCover2
    Range("O20:Q20").Select
        Call Format_GreenCover
    
    GreenCover.Range("Q18").Value = PROJECT_DATA.Range("A2").Value             ' Copy Road Name
    Range("Q18:V18").Select
        Call Format_GreenCover
    
    GreenCover.Range("P19").Value = PROJECT_DATA.Range("B2").Value               ' SECOND ROAD NAME
    Range("P19:V19").Select
        Call Format_GreenCover
            
    GreenCover.Range("P21").Value = PROJECT_DATA.Range("E2").Value             ' Copy Type of Construction
    Range("P21:V21").Select
        Call Format_GreenCover
        
    GreenCover.Range("N22").Value = PROJECT_DATA.Range("F2").Value                   ' SECOND PROJECT DESCRIPTION
    Range("N22:V22").Select
        Call Format_GreenCover
                
    GreenCover.Range("P24").Value = PROJECT_DATA.Range("S2").Value        ' Gross Length
    GreenCover.Range("P24").Select
        Call Format_GreenCover
        Selection.NumberFormat = "0.000"
    
    GreenCover.Range("P23").Value = PROJECT_DATA.Range("Q2").Value      ' Net length
    GreenCover.Range("P23").Select
        Call Format_GreenCover
        Selection.NumberFormat = "0.000"

    GreenCover.Range("U24").Value = PROJECT_DATA.Range("J2").Value       ' Working Days
    GreenCover.Range("U24").Select
        Call Format_GreenCover
        Selection.NumberFormat = "0"
        
    
        GreenCover.Range("D21").Value = PROJECT_DATA.Range("AB2").Value       ' Board President Name
        GreenCover.Range("D21").Select
            Call Format_GreenCover2
        GreenCover.Range("C17").Value = PROJECT_DATA.Range("AC2").Value       ' B
        GreenCover.Range("C17").Select
            Call Format_GreenCover2
        GreenCover.Range("C19").Value = PROJECT_DATA.Range("AD2").Value       ' C
        GreenCover.Range("C19").Select
            Call Format_GreenCover2
        GreenCover.Range("G17").Value = PROJECT_DATA.Range("AE2").Value       ' D
        GreenCover.Range("G17").Select
            Call Format_GreenCover2
        GreenCover.Range("G19").Value = PROJECT_DATA.Range("AF2").Value       ' E
        GreenCover.Range("G19").Select
            Call Format_GreenCover2
            
        GreenCover.Range("I21").Value = PROJECT_DATA.Range("W2").Value       ' Board President District
        GreenCover.Range("I21").Select
            Call Format_GreenCover2
            Selection.NumberFormat = "0"
        GreenCover.Range("E17").Value = PROJECT_DATA.Range("X2").Value       '
        GreenCover.Range("E17").Select
            Call Format_GreenCover2
            Selection.NumberFormat = "0"
        GreenCover.Range("E19").Value = PROJECT_DATA.Range("Y2").Value       '
        GreenCover.Range("E19").Select
            Call Format_GreenCover2
            Selection.NumberFormat = "0"
        GreenCover.Range("J17").Value = PROJECT_DATA.Range("Z2").Value       '
        GreenCover.Range("J17").Select
            Call Format_GreenCover2
            Selection.NumberFormat = "0"
        GreenCover.Range("J19").Value = PROJECT_DATA.Range("AA2").Value       '
        GreenCover.Range("J19").Select
            Call Format_GreenCover2
            Selection.NumberFormat = "0"

End Sub


 Sub Form_GreenCoverNEW()

' Copying the project_data values to the Green Cover Sheet'
    GreenCoverNEW.Range("E16").Value = PROJECT_DATA.Range("D2").Value                ' Copy County Name
    GreenCoverNEW.Range("E16").Select
        Call Format_GreenCover2
    GreenCoverNEW.Range("F25").Value = "PROJECT NO " & PROJECT_DATA.Range("M2").Value               ' Copy Project Number
    GreenCoverNEW.Range("F25").Select
            Selection.Font.Bold = True
            Selection.Font.Size = 16
    GreenCoverNEW.Range("E59").Value = PROJECT_DATA.Range("D2").Value
    GreenCoverNEW.Range("E59").Select
        Call Format_GreenCover2
    GreenCoverNEW.Range("D28").Value = PROJECT_DATA.Range("D2").Value
    GreenCoverNEW.Range("D28").Select
    
    Range("E16:I16").Select
        Call Format_GreenCover
        Call Format_GreenCover2
    Range("E59:H59").Select
        Call Format_GreenCover
        Call Format_GreenCover2
    Range("D28:F28").Select
        Call Format_GreenCover
    
    GreenCoverNEW.Range("F26").Value = PROJECT_DATA.Range("A2").Value             ' Copy Road Name
    Range("F26:K26").Select
        Call Format_GreenCover
    
    GreenCoverNEW.Range("E27").Value = PROJECT_DATA.Range("B2").Value               ' SECOND ROAD NAME
    Range("E27:K27").Select
        Call Format_GreenCover
            
    GreenCoverNEW.Range("E29").Value = PROJECT_DATA.Range("E2").Value             ' Copy Type of Construction
    Range("E29:K29").Select
        Call Format_GreenCover
        
    GreenCoverNEW.Range("C30").Value = PROJECT_DATA.Range("F2").Value                   ' SECOND PROJECT DESCRIPTION
    Range("C30:K30").Select
        Call Format_GreenCover
                
    GreenCoverNEW.Range("E31").Value = PROJECT_DATA.Range("S2").Value        ' Gross Length
    GreenCoverNEW.Range("E31").Select
        Call Format_GreenCover
        Selection.NumberFormat = "0.000"
    
    GreenCoverNEW.Range("E32").Value = PROJECT_DATA.Range("Q2").Value      ' Net length
    GreenCoverNEW.Range("E32").Select
        Call Format_GreenCover
        Selection.NumberFormat = "0.000"

    GreenCoverNEW.Range("J32").Value = PROJECT_DATA.Range("J2").Value       ' Working Days
    GreenCoverNEW.Range("J32").Select
        Call Format_GreenCover
        Selection.NumberFormat = "0"
        
    
        GreenCoverNEW.Range("E65").Value = PROJECT_DATA.Range("AB2").Value       ' Board President Name
        GreenCoverNEW.Range("E65").Select
            Call Format_GreenCover2
        GreenCoverNEW.Range("C61").Value = PROJECT_DATA.Range("AC2").Value       ' B
        GreenCoverNEW.Range("C61").Select
            Call Format_GreenCover2
        GreenCoverNEW.Range("C63").Value = PROJECT_DATA.Range("AD2").Value       ' C
        GreenCoverNEW.Range("C63").Select
            Call Format_GreenCover2
        GreenCoverNEW.Range("H61").Value = PROJECT_DATA.Range("AE2").Value       ' D
        GreenCoverNEW.Range("H62").Select
            Call Format_GreenCover2
        GreenCoverNEW.Range("H63").Value = PROJECT_DATA.Range("AF2").Value       ' E
        GreenCoverNEW.Range("H63").Select
            Call Format_GreenCover2
            
        GreenCoverNEW.Range("J65").Value = PROJECT_DATA.Range("W2").Value       ' Board President District
        GreenCoverNEW.Range("J65").Select
            Call Format_GreenCover2
            Selection.NumberFormat = "0"
        GreenCoverNEW.Range("F61").Value = PROJECT_DATA.Range("X2").Value       '
        GreenCoverNEW.Range("F61").Select
            Call Format_GreenCover2
            Selection.NumberFormat = "0"
        GreenCoverNEW.Range("F63").Value = PROJECT_DATA.Range("Y2").Value       '
        GreenCoverNEW.Range("F63").Select
            Call Format_GreenCover2
            Selection.NumberFormat = "0"
        GreenCoverNEW.Range("K61").Value = PROJECT_DATA.Range("Z2").Value       '
        GreenCoverNEW.Range("K61").Select
            Call Format_GreenCover2
            Selection.NumberFormat = "0"
        GreenCoverNEW.Range("K63").Value = PROJECT_DATA.Range("AA2").Value       '
        GreenCoverNEW.Range("K63").Select
            Call Format_GreenCover2
            Selection.NumberFormat = "0"

End Sub


