Attribute VB_Name = "Module_ProgramUpdate"
Attribute VB_Name = "Module_ProgramUpdate"
''**************************************************************
''Variable for this module  ************************************
''**************************************************************
    
' Determine files needed
Dim baseURL As String
Dim files As Variant

' Create download files needed
Dim fullURL As String
Dim localPath As String

' Create temp folder
Dim tempfolder As String
Dim file As String

' Replace files
Dim fileName As String
Dim moduleName As String



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
        stream.Type = 1
        stream.Open
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
    
    baseURL = "https://raw.githubusercontent.com/jchammons1/Excel-Program-Version-2.0/main/"
    ''  "Module_ProgramUpdate.bas" cannot be included because this is the code that is running to replace the files
    ''  It will produce Module_ProgramUpdate1.bas
    files = Array( _
        "Module_900_WebBidList.bas", "Module_902.bas", "Module_AverageUnitCost.bas", _
        "Module_BIDTAB.bas", "Module_CAD.bas", "Module_CAD_Calculate.bas", _
        "Module_CAD_ChangeOrder.bas", "Module_CAD_Final.bas", "Module_CAD001_Signature.bas", _
        "Module_DatabaseUpdate.bas", "Module_DatePicker.bas", "Module_DELETE.bas", _
        "Module_EngineerPayment.bas", "Module_Estimates.bas", "Module_FlagItems.bas", _
        "Module_Format.bas", "Module_FuelCalculations.bas", "Module_GreenCover.bas", "Module_ImageHeader.bas", _
        "Module_LettingResults.bas", "Module_ListBox.bas", "Module_MaterialReport.bas", _
        "Module_MoblizationTraffic.bas", "Module_Music.bas", "Module_Preparation.bas", "Module_PrintArea.bas", _
        "Module_ProcessPayItem.bas", "Module_Sort.bas", _
        "UserForm_BaseFuelData.zip", "UserForm_BidDatePicker.zip", "UserForm_CAD_Signature.zip", _
        "UserForm_DASHBOARD.zip", "UserForm_DocumentHeading.zip", "UserForm_EditPayItem.zip", _
        "UserForm_EditSuppListBox.zip", "UserForm_EditSuppTypes.zip", "UserForm_EditTypeListBox.zip", _
        "UserForm_EditTypes.zip", "UserForm_EstimateDate.zip", "UserForm_LettingResults.zip", "UserForm_MonthlyFuelData.zip" _
    )
End Sub


''
'' Download all the files from GitHub
''
Sub Download_Files_Needed()
    Dim i As Long
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
    Dim i As Long
    Dim filepath As String
    
    For i = LBound(files) To UBound(files)
    
        fileName = files(i)
    
        ' Only import .bas or .frm files
    If Right(fileName, 4) = ".bas" Then
    
        ' --- existing logic for modules ---
        filepath = tempfolder & fileName
    
        On Error Resume Next
        ActiveWorkbook.VBProject.VBComponents.Remove _
            ActiveWorkbook.VBProject.VBComponents(Left(fileName, Len(fileName) - 4))
        On Error GoTo 0
    
        ActiveWorkbook.VBProject.VBComponents.Import filepath
    
    ElseIf Right(fileName, 4) = ".zip" Then

    ' --- NEW LOGIC FOR USERFORM ZIP ---
    
    Dim zipPath As String
    Dim formName As String
    
    zipPath = tempfolder & fileName
    formName = Left(fileName, Len(fileName) - 4)

    Debug.Print "ZIP PATH:", zipPath
    Debug.Print "EXISTS?", Dir(zipPath)
    
    DoEvents
    Application.Wait Now + TimeValue("0:00:01")


    ' 1. Unzip
    Call UnzipFile(zipPath, tempfolder)

    ' 2. Remove existing form
    On Error Resume Next
    ActiveWorkbook.VBProject.VBComponents.Remove _
        ActiveWorkbook.VBProject.VBComponents(formName)
    On Error GoTo 0

    ' 3. Import the .frm (FRX will load automatically)
    ActiveWorkbook.VBProject.VBComponents.Import _
        tempfolder & formName & ".frm"

End If
    
    Next i
End Sub



''
'' Creates temp folder on the users computer.  This is needed in order to have a temporary location to store the files before replacing the existing modules and user forms
''
Sub Create_TempFolder()
    
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

Sub UnzipFile(zipPath As String, extractTo As String)

    Dim command As String
    Dim wsh As Object

    ' Ensure destination folder exists
    If Dir(extractTo, vbDirectory) = "" Then MkDir extractTo

    ' PowerShell unzip command
    command = "powershell -command ""Expand-Archive -Path '" & zipPath & "' -DestinationPath '" & extractTo & "' -Force"""

    ' ? THIS IS THE FIX (wait for completion)
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run command, 0, True   ' True = wait until finished

End Sub

