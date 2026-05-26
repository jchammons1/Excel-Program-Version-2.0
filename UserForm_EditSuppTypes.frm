VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_EditSuppTypes 
   Caption         =   "Supplemental or Correction Types Form"
   ClientHeight    =   3390
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17610
   OleObjectBlob   =   "UserForm_EditSuppTypes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_EditSuppTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_currentRow_SuppType As Long

'Add a private memember. Patches the current row in the list box
Public Property Let currentRow(ByVal newCurrentRow_SuppType As Long)
    m_currentRow_SuppType = newCurrentRow_SuppType
End Property


 Function ModifyEditPayitem_SuppType()
'Save data for selected pay item from the List Box
    If MsgBox("Do you want to overwrite this record?", _
        vbYesNo, "Save changes") = vbYes Then
       ' Else
            With PAYITEMTYPE_SUPPLEMENTAL.Range("A2").Offset(m_currentRow_SuppType)
        ' Current Period Quantity Tab
            .Cells(1, 1).Value = txt_PayItemSuppTypeEdit.Value

            End With
        End If

Unload Me
End Function



 Sub Button_SaveSuppType_Click()
'Saved modifications to the pay item type
    PAYITEMTYPE_SUPPLEMENTAL.Unprotect
    PAYITEMTYPE_SUPPLEMENTAL.Visible = True
        Call ModifyEditPayitem_SuppType
    PAYITEMTYPE_SUPPLEMENTAL.Protect
    PAYITEMTYPE_SUPPLEMENTAL.Visible = False
End Sub



 Sub UserForm_Activate()
'Activate the Edit Form with values of the selected row from the List Box
    Call LoadEditData_SuppType
End Sub


'Load data from the selected row in the list box
 Sub LoadEditData_SuppType()
    With PAYITEMTYPE_SUPPLEMENTAL.Range("A2").Offset(m_currentRow_SuppType)
    'General tab
        txt_PayItemSuppTypeEdit.Value = .Cells(1, 1).Value
    End With
End Sub

       

