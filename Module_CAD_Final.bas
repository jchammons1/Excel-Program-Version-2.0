Attribute VB_Name = "Module_CAD_Final"
'
'
'
'############################################ FINAL CAD-002 begin#############################################################
'
'
'

Sub Final_CAD()
' Allows the user to override the calculation of the traffic maintenance and make it equal to 1
    ActiveSheet.Select
    ActiveSheet.Unprotect Password:="roadway123"
    Application.ScreenUpdating = False
 ''WARNING Message
    If MsgBox("Do you want to create the final INVOICE for this project? This action is ONLY used to create the last invoice. DO NOT use this function unless this is your last invoice.", vbYesNo, "Save changes") = vbYes Then
     
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ActiveSheet.Range("A53").Select

        If Not ws.Cells.Find("S-618-A", SearchOrder:=xlByRows, SearchDirection:=xlNext) Is Nothing Then
'Find cell with "S-618-A"                   ############# TRAFFIC MAINTENANCE ###################
            Cells.Find(What:="S-618-A", After:=ActiveCell, LookIn:=xlValues _
                , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False).Activate
                
            ActiveCell.Select
            ActiveCell.Offset(0, 3).Range("A1").Select
            ActiveCell.Value = 1 - ActiveCell.Offset(0, 1).Range("A1").Value   ' Current Period
                Selection.NumberFormat = "0.000"
            ActiveCell.Offset(0, 1).Range("A1").Select
            ActiveCell.Value = 1                                               ' Allowed to Date
                Selection.NumberFormat = "0.000"
            ActiveCell.Offset(0, 3).Range("A1").Select
            ActiveCell.Value = ActiveCell.Offset(0, -3).Range("A1").Value * ActiveCell.Offset(0, -1).Range("A1").Value      ' Current Accumulative amount
            
            ActiveCell.Offset(0, 58).Range("A1").Value = ActiveCell.Value
            ActiveCell.Offset(0, 57).Range("A1").Value = ActiveCell.Offset(0, -4).Range("A1").Value
            ActiveCell.Offset(0, 59).Range("A1").Value = ActiveCell.Offset(0, -3).Range("A1").Value
            ActiveCell.Offset(0, 60).Range("A1").Value = ActiveCell.Offset(0, -3).Range("A1").Value * ActiveCell.Offset(0, -1).Range("A1").Value
            
        End If
        
        
    Range("C7").Select
    Range("C7").Value = "(Progress - FINAL)"
    Selection.Font.Bold = True
    Range("C49").Select
    Range("C49").Value = "(Progress - FINAL)"
    Selection.Font.Bold = True
            
    Range("J10").Value = Range("FB3").Value     ' Total participating items to 2 decimals
    Range("J11").Value = Range("FC3").Value     ' Total non-participating items to 2 decimals
    
    Range("H13").Select                         ' Current retainage
    ActiveCell.FormulaR1C1 = "=ABS(RC[1])"
    Range("J13").Select                         ' total retainage to date
    Selection.ClearContents
    ActiveCell.FormulaR1C1 = "0"                ' add retainage back to totals
    
    Range("H10:J28").Select
    Selection.NumberFormat = "$#,##0.00"
    Selection.ShrinkToFit = True
    
    Range("J32").Select
    ActiveCell.FormulaR1C1 = "1"
    
    ActiveSheet.Select
    ActiveSheet.Protect Password:="roadway123"
    Application.ScreenUpdating = True
    Range("G7").Select
    End If
    

       
End Sub
    


