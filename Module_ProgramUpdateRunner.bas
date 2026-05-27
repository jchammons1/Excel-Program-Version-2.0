Attribute VB_Name = "Module_ProgramUpdateRunner"
''
'' This was created to front run the Module_ProgramUpdate.bas  In order to update this module without creating Module_ProgramUpdate1.bas
'' you have to update it separately.  The module code cannot run and remove itself and replace itself at the same time
''

Private Const BASE_URL As String = "https://raw.githubusercontent.com/jchammons1/Excel-Program-Version-2.0/main/"
Private Const UPDATER_BAS As String = "Module_ProgramUpdate.bas"

Public Sub RunUpdate()

    ' 1) Create temp folder
    Dim tempfolder As String
    tempfolder = Environ$("TEMP") & "\OSARC_Update\"
    If Dir$(tempfolder, vbDirectory) = "" Then MkDir tempfolder

    ' 2) Download the latest Module_ProgramUpdate.bas FIRST
    Dim url As String, savePath As String
    url = BASE_URL & UPDATER_BAS
    savePath = tempfolder & UPDATER_BAS
    DownloadFileBinary url, savePath

    ' 3) Remove the existing Module_ProgramUpdate (so import won’t become ...1)
    RemoveVBComponentIfExists ActiveWorkbook, "Module_ProgramUpdate"

    ' 4) Import the new one
    ActiveWorkbook.VBProject.VBComponents.Import savePath

    ' 5) Now run the updated updater
    ' Using Application.Run ensures it binds to the latest loaded code
    Application.Run "UpdateAllModules"

End Sub

Private Sub DownloadFileBinary(ByVal url As String, ByVal savePath As String)
    Dim http As Object, stream As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.Send

    If http.Status <> 200 Then
        MsgBox "Download failed: " & url & vbCrLf & "HTTP Status: " & http.Status, vbExclamation
        Exit Sub
    End If

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 'binary
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile savePath, 2
    stream.Close
End Sub

Private Sub RemoveVBComponentIfExists(ByVal wb As Workbook, ByVal compName As String)
    Dim comp As Object
    On Error Resume Next
    Set comp = wb.VBProject.VBComponents(compName)
    On Error GoTo 0

    If Not comp Is Nothing Then
        ' Don't remove document modules
        If comp.Type <> 100 Then
            wb.VBProject.VBComponents.Remove comp
        End If
    End If
End Sub
