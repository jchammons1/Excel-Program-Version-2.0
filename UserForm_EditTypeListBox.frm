VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_EditTypeListBox 
   Caption         =   "User Defined Type Form"
   ClientHeight    =   11655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17205
   OleObjectBlob   =   "UserForm_EditTypeListBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_EditTypeListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


 Sub Button_EditPayItemType_Click()
'Modify a pay item and open Edit Pay Item User Form
    PAYITEMTYPE.Visible = True
    PAYITEMTYPE.Unprotect
    Call EditRow_Type
End Sub

Sub EditRow_Type()
'Shows the edit form for the pay item selected in the list box
    Dim frm_type As New UserForm_EditTypes
    frm_type.currentRow = ListBox_UserDefinedType.ListIndex
    frm_type.Show vbModal
End Sub



 Sub UserForm_Activate()
    PAYITEMTYPE.Visible = True
    PAYITEMTYPE.Unprotect
    Call AddTypeToListBox_UserDefined
End Sub

Sub AddTypeToListBox_UserDefined()
    'Get the type range on the Data Validation sheet
    Dim Type_Range As Range
    Set Type_Range = GetType_Range
    'Link the data to the list box.  Setup list box headings and column widths
        With ListBox_UserDefinedType
            .RowSource = Type_Range.Address(external:=True)
            .ColumnCount = Type_Range.Columns.Count
            .ColumnWidths = "750"
            .ColumnHeads = True
            .ListIndex = 0
        End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        ' Hide the sheet
        PAYITEMTYPE.Visible = xlSheetHidden
    End If
End Sub







