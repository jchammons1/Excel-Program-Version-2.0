Attribute VB_Name = "Module_Sort"
'                                                                                                                         '
'                                                                                                                         '
'************************************* Sorting the pay items for the estimates ############################################
'                                                                                                                         '
'                                                                                                                         '

 Sub SortPayItems()
    Database.Select
    Cells.Select
    ActiveWorkbook.Worksheets("Database").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Database").Sort.SortFields.Add2 Key:=Range( _
        "AV2:AV9999"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Database").Sort
        .SetRange Range("A1:CS9999")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

'
'
 Sub SortPayItems_900_WebBidList()

End Sub

