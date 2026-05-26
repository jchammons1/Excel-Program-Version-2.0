VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_EditTypes 
   Caption         =   "User Defined Edit Form"
   ClientHeight    =   3480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17505
   OleObjectBlob   =   "UserForm_EditTypes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_EditTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private m_currentRow_Type As Long

'Add a private memember. Patches the current row in the list box
Public Property Let currentRow(ByVal newCurrentRow_Type As Long)
    m_currentRow_Type = newCurrentRow_Type
End Property


 Function ModifyEditPayitem_Type()
'Save data for selected pay item from the List Box
    If MsgBox("Do you want to overwrite this record?", _
        vbYesNo, "Save changes") = vbYes Then
       ' Else
            With PAYITEMTYPE.Range("A2").Offset(m_currentRow_Type)
        ' Current Period Quantity Tab
            .Cells(1, 1).Value = txt_PayItemTypeEdit.Value

            End With
        End If

Unload Me
End Function



 Sub Button_SaveUserDefinedType_Click()
'Saved modifications to the pay item type
    PAYITEMTYPE.Unprotect
    PAYITEMTYPE.Visible = True
        Call ModifyEditPayitem_Type
    PAYITEMTYPE.Protect
    PAYITEMTYPE.Visible = False
End Sub



 Sub UserForm_Activate()
'Activate the Edit Form with values of the selected row from the List Box
    Call LoadEditData_Type
End Sub


'Load data from the selected row in the list box
 Sub LoadEditData_Type()
    With PAYITEMTYPE.Range("A2").Offset(m_currentRow_Type)
    'General tab
        txt_PayItemTypeEdit.Value = .Cells(1, 1).Value
    End With
End Sub

       


