VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_EditSuppListBox 
   Caption         =   "Supplemental Type or Correction Description Form"
   ClientHeight    =   11595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17310
   OleObjectBlob   =   "UserForm_EditSuppListBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_EditSuppListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Button_EditSuppItemType_Click()
'Modify a type
    PAYITEMTYPE_SUPPLEMENTAL.Visible = True
    PAYITEMTYPE_SUPPLEMENTAL.Unprotect
    Call EditRow_SuppType
End Sub


Sub EditRow_SuppType()
'Shows the edit form for the pay item selected in the list box
    Dim frm_supptype As New UserForm_EditSuppTypes
    frm_supptype.currentRow = ListBox_SuppType.ListIndex
    frm_supptype.Show vbModal
End Sub



 Sub UserForm_Activate()
    PAYITEMTYPE_SUPPLEMENTAL.Visible = True
    PAYITEMTYPE_SUPPLEMENTAL.Unprotect
    Call AddTypeToListBox_Supp
End Sub

Sub AddTypeToListBox_Supp()
    'Get the type range on the Data Validation sheet
    Dim SuppType_Range As Range
    Set SuppType_Range = GetSuppType_Range
    'Link the data to the list box.  Setup list box headings and column widths
        With ListBox_SuppType
            .RowSource = SuppType_Range.Address(external:=True)
            .ColumnCount = SuppType_Range.Columns.Count
            .ColumnWidths = "650"
            .ColumnHeads = True
            .ListIndex = 0
        End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        ' Hide the sheet.  This is needed because when you open the form and immediately X out to close.  The sheet does not hide.
        PAYITEMTYPE_SUPPLEMENTAL.Visible = xlSheetHidden
    End If
End Sub
