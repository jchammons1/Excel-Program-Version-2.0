Attribute VB_Name = "Module_ProgramUpdate"

''
'' This function will get the files from GitHub
''
Function DownloadFile(url As String, savePath As String)

    Dim http As Object
    Dim stream As Object

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.Send

    If http.Status = 200 Then
        Set stream = CreateObject("ADODB.Stream")
        stream.Open
        stream.Type = 1
        stream.Write http.responseBody
        stream.SaveToFile savePath, 2
        stream.Close
    Else
        MsgBox "Download failed: " & url
    End If

End Function
''''__________________________________________________________________________________________________!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Sub UpdateAllModules()

    If MsgBox("Update code from GitHub?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
        Call Create_TempFolder
        Call Determine_Files_Needed
        Call Download_Files_Needed
        Call Replace_Files
    MsgBox "Update complete"
    
End Sub


Sub Determine_Files_Needed()
    Dim baseURL As String
    baseURL = "https://github.com/jchammons1/Excel-Program-Version-2.0/tree/main"
    
    Dim files As Variant
    
    files = Array( _
        "Module_900_WebBidList.bas", "Module_902.bas", "Module_AverageUnitCost.bas", _
        "Module_BIDTAB.bas", "Module_CAD.bas", "Module_CAD_Calculate.bas", _
        "Module_CAD_ChangeOrder.bas", "Module_CAD_Final.bas", "Module_CAD001_Signature.bas", _
        "Module_DatabaseUpdate.bas", "Module_DatePicker.bas", "Module_DELETE.bas", _
        "Module_EngineerPayment.bas", "Module_Estimates.bas", "Module_FlagItems.bas", _
        "Module_Format.bas", "Module_FuelCalculations.bas", "Module_GreenCover.bas", "Module_ImageHeader.bas", _
        "Module_LettingResults.bas", "Module_ListBox.bas", "Module_MaterialReport.bas", _
        "Module_MoblizationTraffic.bas", "Module_Music.bas", "Module_Preparation.bas", "Module_PrintArea.bas", _
        "Module_ProcessPayItem.bas", "Module_ProgramUpdate.bas", "Module_Sort.bas", _
        "UserForm_BaseFuelData.frm", "UserForm_BidDatePicker.frm", "UserForm_CAD_Signature.frm", _
        "UserForm_DASHBOARD.frm", "UserForm_DocumentHeading.frm", "UserForm_EditPayItem.frm", _
        "UserForm_EditSuppListBox.frm", "UserForm_EditSuppTypes.frm", "UserForm_EditTypeListBox.frm", _
        "UserForm_EditTypes.frm", "UserForm_EstimateDate.frm", "UserForm_LettingResults.frm", "UserForm_MonthlyFuelData.frm", _
        "UserForm_BaseFuelData.frx", "UserForm_BidDatePicker.frx", "UserForm_CAD_Signature.frx", _
        "UserForm_DASHBOARD.frx", "UserForm_DocumentHeading.frx", "UserForm_EditPayItem.frx", _
        "UserForm_EditSuppListBox.frx", "UserForm_EditSuppTypes.frx", "UserForm_EditTypeListBox.frx", _
        "UserForm_EditTypes.frx", "UserForm_EstimateDate.frx", "UserForm_LettingResults.frx", "UserForm_MonthlyFuelData.frx" _
    )
End Sub



''
'' Download all the files from GitHub
''
Sub Download_Files_Needed()
    Dim i As Long
    Dim fullURL As String
    Dim localPath As String
    
    For i = LBound(files) To UBound(files)
    
        fullURL = baseURL & files(i)
        localPath = tempfolder & files(i)
    
        Call DownloadFile(fullURL, localPath)
    
    Next i
End Sub




''
'' Replace modules and user forms
''
Sub Replace_Files()
    Dim fileName As String
    Dim moduleName As String
    
    For i = LBound(files) To UBound(files)
    
        fileName = files(i)
    
        ' Only import .bas or .frm files
        If Right(fileName, 4) = ".bas" Or Right(fileName, 4) = ".frm" Or Right(fileName, 4) = ".frx" Then
    
            moduleName = Left(fileName, Len(fileName) - 4)
    
            ' Remove existing
            On Error Resume Next
            ActiveWorkbook.VBProject.VBComponents.Remove _
                ActiveWorkbook.VBProject.VBComponents(moduleName)
            On Error GoTo 0
    
            ' Import new
            ActiveWorkbook.VBProject.VBComponents.Import tempfolder & fileName
    
        End If
    
    Next i
End Sub



''
'' Creates temp folder on the users computer.  This is needed in order to have a temporary location to store the files before replacing the existing modules and user forms
''
Sub Create_TempFolder()
    Dim tempfolder As String
    Dim file As String
    
    tempfolder = Environ("TEMP") & "\OSARC_Update\"
    
    If Dir(tempfolder, vbDirectory) = "" Then
        MkDir tempfolder
    Else
        ' Delete all files inside folder
        file = Dir(tempfolder & "*.*")
    
        Do While file <> ""
            Kill tempfolder & file
            file = Dir
        Loop
    End If
End Sub



''
'' Remove VBA modules and user forms
''
Sub Replace_Module(modName As String, filePath As String)

    If MsgBox("This will update VBA modules. Continue?", vbYesNo + vbWarning) = vbNo Then Exit Sub
    
        ' Backup first
        Call Backup_Module(modName)
    
        ' Remove existing module
        On Error Resume Next
        ActiveWorkbook.VBProject.VBComponents.Remove _
            ActiveWorkbook.VBProject.VBComponents(modName)
        On Error GoTo 0
    
        ' Import new module
        ActiveWorkbook.VBProject.VBComponents.Import filePath
        
End Sub


''
'' Backup modules in case something breaks so we can instantly roll back the changes
''
Sub Backup_Module(modName As String)

    Dim backupPath As String
    
    backupPath = "C:\OSARC\VBA_Backup\" & modName & ".bas"
    
    On Error Resume Next
    ActiveWorkbook.VBProject.VBComponents(modName).Export backupPath
    On Error GoTo 0

End Sub
