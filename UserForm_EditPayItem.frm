VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_EditPayItem 
   Caption         =   "Edit Pay Item Form"
   ClientHeight    =   5775
   ClientLeft      =   75
   ClientTop       =   300
   ClientWidth     =   21450
   OleObjectBlob   =   "UserForm_EditPayItem.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_EditPayItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_currentRow_Database As Long

'Add a private memember. Patches the current row in the list box
Public Property Let currentRow(ByVal newCurrentRow_Database As Long)
    m_currentRow_Database = newCurrentRow_Database
End Property

 Sub Button_ModifyPayItemEdit_Click()
'Saved modifications to the pay item
    Database.Unprotect
    Database.Visible = True
     Call ModifyEditPayitem_Database
    Database.Protect
    Database.Visible = False
End Sub

 Function ModifyEditPayitem_Database()
'Save data for selected pay item from the List Box
    If MsgBox("Do you want to overwrite this record? Ensure all fields are populated with 'NA' or the correct data. Then, save the pay item edit", _
        vbYesNo, "Save changes") = vbYes Then
        ' If a value is blank, then message the user and tell them to populate the data
            If Me.txt_QuantityEdit.Value = "" Or Me.txt_PayItemDescriptionEdit.Value = "" Or Me.txt_UnitEdit.Value = "" _
            Or Me.txt_UnitPriceEdit.Value = "" Or Me.cmb_FuelCodeEdit.Value = "" _
            Or Me.txt_BinderPercentEdit.Value = "" Or Me.cmb_BinderTypeEdit.Value = "" _
            Or Me.cmb_TypeEdit.Value = "" Or Me.cmb_ParticipatingEdit.Value = "" Or Me.txt_SortOrderEdit.Value = "" Then
                MsgBox ("Complete the required data fields.  Data not save!!")
        Else
            With Database.Range("A2").Offset(m_currentRow_Database)
    ' General Tab
            .Cells(1, 1).Value = txt_PayItemNoEdit.Value                            ' locked
            .Cells(1, 2).Value = txt_PayItemDescriptionEdit.Value
            .Cells(1, 3).Value = txt_QuantityEdit.Value
            .Cells(1, 4).Value = txt_UnitEdit.Value
            .Cells(1, 5).Value = txt_UnitPriceEdit.Value
            .Cells(1, 6).Value = txt_QuantityEdit.Value * txt_UnitPriceEdit.Value               ' Calculate Subtotal
            .Cells(1, 46).Value = cmb_TypeEdit.Value
            .Cells(1, 8).Value = cmb_FuelCodeEdit.Value
            .Cells(1, 48).Value = txt_SortOrderEdit.Value
            .Cells(1, 70).Value = txt_SpecYearEdit.Value                            ' locked
        ' Mix Design Tab
            .Cells(1, 18).Value = txt_BinderPercentEdit.Value                                   ' Mix Design binder data
            .Cells(1, 19).Value = cmb_BinderTypeEdit.Value
        ' Change Order Tab
            .Cells(1, 47).Value = cmb_ParticipatingEdit.Value
            .Cells(1, 51).Value = cmb_ChangeOrderEdit.Value
        ' Monthly Fuel Adjustment Tab
            '.Cells(1, 9).Value = txt_PlacementDateEdit.Value                                   ' Placement Date on the Fuel Adjustment tab.  Deleting this 12-18-24
            .Cells(1, 10).Value = cmb_BaseFuelMonthEdit.Value                                   ' Base Fuel adjustment Values
            .Cells(1, 11).Value = cmb_BaseFuelYearEdit.Value
            .Cells(1, 14).Value = txt_BaseGas_Edit.Value
            .Cells(1, 15).Value = txt_BaseDiesel_Edit.Value
            .Cells(1, 20).Value = txt_BasePG6422_Edit.Value
            .Cells(1, 21).Value = txt_BasePG6722_Edit.Value
            .Cells(1, 22).Value = txt_BasePG7622_Edit.Value
            .Cells(1, 23).Value = txt_BasePG8222_Edit.Value
            .Cells(1, 24).Value = txt_BaseSS1_Edit.Value
            .Cells(1, 25).Value = txt_BaseCRS2_Edit.Value
            .Cells(1, 26).Value = txt_BaseCRS2P_Edit.Value
            .Cells(1, 27).Value = txt_BaseEA1_Edit.Value
            .Cells(1, 28).Value = txt_BaseCSS1Undiluted_Edit.Value
            .Cells(1, 29).Value = txt_BaseCSS1_Edit.Value
            .Cells(1, 12).Value = cmb_CurrentFuelMonthEdit.Value                                        ' Current Fuel adjustment Values
            .Cells(1, 13).Value = cmb_CurrentFuelYearEdit.Value
            .Cells(1, 16).Value = txt_CurrentGas_Edit.Value
            .Cells(1, 17).Value = txt_CurrentDiesel_Edit.Value
            .Cells(1, 30).Value = txt_CurrentPG6422_Edit.Value
            .Cells(1, 31).Value = txt_CurrentPG6722_Edit.Value
            .Cells(1, 32).Value = txt_CurrentPG7622_Edit.Value
            .Cells(1, 33).Value = txt_CurrentPG8222_Edit.Value
            .Cells(1, 34).Value = txt_CurrentSS1_Edit.Value
            .Cells(1, 35).Value = txt_CurrentCRS2_Edit.Value
            .Cells(1, 36).Value = txt_CurrentCRS2P_Edit.Value
            .Cells(1, 37).Value = txt_CurrentEA1_Edit.Value
            .Cells(1, 38).Value = txt_CurrentCSS1Undiluted_Edit.Value
            .Cells(1, 39).Value = txt_CurrentCSS1_Edit.Value
        ' Hidden modified time and by tracking
            .Cells(1, 66).Value = Now
            .Cells(1, 67).Value = ActiveWorkbook.BuiltinDocumentProperties("Last Author")
            .Cells(1, 70).Value = txt_SpecYearEdit.Value
            
            End With
            End If
    End If
Unload Me
End Function



'Activate the Edit Form with values of the selected row from the List Box
 Sub UserForm_Activate()
    Call LoadEditData_Database
End Sub

'Load data from the selected row in the list box
 Sub LoadEditData_Database()
    With Database.Range("A2").Offset(m_currentRow_Database)
    'General tab
        txt_PayItemNoEdit.Value = .Cells(1, 1).Value
        txt_PayItemDescriptionEdit.Value = .Cells(1, 2).Value
        txt_QuantityEdit.Value = .Cells(1, 3).Value
        txt_UnitEdit.Value = .Cells(1, 4).Value
        txt_UnitPriceEdit.Value = .Cells(1, 5).Value
        txt_SubtotalEdit.Value = .Cells(1, 6).Value
        cmb_TypeEdit.Value = .Cells(1, 46).Value
        cmb_FuelCodeEdit.Value = .Cells(1, 8).Value
        txt_SortOrderEdit.Value = .Cells(1, 48).Value
        txt_SpecYearEdit.Value = .Cells(1, 70).Value    ' Spec Year
    'Mix Design Tab
         txt_BinderPercentEdit.Value = .Cells(1, 18).Value                                  ' Mix Design binder data
         cmb_BinderTypeEdit.Value = .Cells(1, 19).Value
    'Change Order Tab
        cmb_ParticipatingEdit.Value = .Cells(1, 47).Value
        cmb_ChangeOrderEdit = .Cells(1, 51).Value
    ' Monthly Fuel Adjustment Tab
        cmb_BaseFuelMonthEdit.Value = .Cells(1, 10).Value                                   ' Base Fuel adjustment Values
        cmb_BaseFuelYearEdit.Value = .Cells(1, 11).Value
        txt_BaseGas_Edit.Value = .Cells(1, 14).Value
        txt_BaseDiesel_Edit.Value = .Cells(1, 15).Value
        txt_BasePG6422_Edit.Value = .Cells(1, 20).Value
        txt_BasePG6722_Edit.Value = .Cells(1, 21).Value
        txt_BasePG7622_Edit.Value = .Cells(1, 22).Value
        txt_BasePG8222_Edit.Value = .Cells(1, 23).Value
        txt_BaseSS1_Edit.Value = .Cells(1, 24).Value
        txt_BaseCRS2_Edit.Value = .Cells(1, 25).Value
        txt_BaseCRS2P_Edit.Value = .Cells(1, 26).Value
        txt_BaseEA1_Edit.Value = .Cells(1, 27).Value
        txt_BaseCSS1Undiluted_Edit.Value = .Cells(1, 28).Value
        txt_BaseCSS1_Edit.Value = .Cells(1, 29).Value
        cmb_CurrentFuelMonthEdit.Value = .Cells(1, 12).Value                                        ' Current Fuel adjustment Values
        cmb_CurrentFuelYearEdit.Value = .Cells(1, 13).Value
        txt_CurrentGas_Edit.Value = .Cells(1, 16).Value
        txt_CurrentDiesel_Edit.Value = .Cells(1, 17).Value
        txt_CurrentPG6422_Edit.Value = .Cells(1, 30).Value
        txt_CurrentPG6722_Edit.Value = .Cells(1, 31).Value
        txt_CurrentPG7622_Edit.Value = .Cells(1, 32).Value
        txt_CurrentPG8222_Edit.Value = .Cells(1, 33).Value
        txt_CurrentSS1_Edit.Value = .Cells(1, 34).Value
        txt_CurrentCRS2_Edit.Value = .Cells(1, 35).Value
        txt_CurrentCRS2P_Edit.Value = .Cells(1, 36).Value
        txt_CurrentEA1_Edit.Value = .Cells(1, 37).Value
        txt_CurrentCSS1Undiluted_Edit.Value = .Cells(1, 38).Value
        txt_CurrentCSS1_Edit.Value = .Cells(1, 39).Value

    End With
End Sub



Sub UpdateTypeComboBoxEdit()
        Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DATA_VALIDATION") ' Change to your sheet name
    
    Select Case ws.Range("Z2").Value ' Change A1 to your cell
        Case "No"
            Me.cmb_TypeEdit.RowSource = "PAYITEMTYPE!A2:A55" ' Change to your range name
        Case "Yes"
            Me.cmb_TypeEdit.RowSource = "PAYITEMTYPE_SUPPLEMENTAL!A2:A51"
        ' Add more cases as needed
    End Select
End Sub


Private Sub UserForm_Initialize()
    Call UpdateTypeComboBoxEdit
End Sub
