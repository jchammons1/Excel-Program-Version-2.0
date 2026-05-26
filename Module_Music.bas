Attribute VB_Name = "Module_Music"
Option Explicit
Dim localFilePath As String

Declare PtrSafe Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long


Sub DownloadWAVFile()
    If DATA_VALIDATION.Range("AA2") = "On" Then
        Dim http As Object
        Dim fileUrl As String
        
        Dim fileStream As Object
    
        ' Initialize variables
        fileUrl = "https://www.osarc.ms.gov/SuperMarioBros30sec.wav"
        localFilePath = Environ("TEMP") & "\downloaded_file.wav"
    
        ' Create XMLHTTP object
        Set http = CreateObject("MSXML2.XMLHTTP")
        http.Open "GET", fileUrl, False
        http.Send
    
        ' Create ADODB Stream object
        Set fileStream = CreateObject("ADODB.Stream")
        fileStream.Type = 1 ' Binary
        fileStream.Open
        fileStream.Write http.responseBody
        fileStream.SaveToFile localFilePath, 2 ' Overwrite if file exists
        fileStream.Close
    
        ' Clean up
        Set http = Nothing
        Set fileStream = Nothing
    
        Call PlaySoundFile
    End If
End Sub


Sub PlaySoundFile()
    Dim soundFile As String
    soundFile = localFilePath ' Change this to the path of your sound file
    PlaySound soundFile, 0, &H1 ' &H1 is the flag for synchronous play
End Sub

