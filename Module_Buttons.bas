Attribute VB_Name = "Module_Buttons"
Sub Show_DASHBOARD()
    UserForm_DASHBOARD.Show vbModeless  ' Add modeless so the count down can show up.  if this is vbmodal by default, it will not let you leave the form open
End Sub

Sub Show_CADEditDataForm()
    ActiveSheet.Unprotect
    UserForm_EditCADListBox.Show
End Sub

Sub Show_LettingResultsForm()
    UserForm_LettingResults.Show vbModeless
End Sub

Sub Open_PayItemUnitCost()
    Dim url As String
    url = "https://forms.office.com/Pages/ResponsePage.aspx?id=HsixX3TZMkKZmHKzsCXWhXeQlq1UYTZMqfDsB6bL2M9UMjdPNUszMjNWM1A4WkJZQzdUQUY3QkIxWS4u"
    ActiveWorkbook.FollowHyperlink url
End Sub

Sub Open_ProgramManual()
    Dim url As String
    url = "https://drive.google.com/file/d/1kEHmPJahaOmEQF-06cuhkZJn8NB2bxIp/view"
    ActiveWorkbook.FollowHyperlink url
End Sub

