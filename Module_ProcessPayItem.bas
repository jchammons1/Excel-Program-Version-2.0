Attribute VB_Name = "Module_ProcessPayItem"
Sub Process_PayItems()


' Retrieve values from Pay Item List sheet and inputs the values on the form
Dim PayItemNo As Variant
Dim PayItemDescription As Variant
Dim Unit As Variant
Dim FuelCode As Variant
Dim Participating As Variant
'Dim Type as Variant
'Dim PAYITEMTYPE As Variant
Dim SortOrder As Variant

PayItemNo_Lookup = Application.VLookup(cmb_PayItemSearch.Value, PayItemList.Range("A1:I9999"), 2, False)
PayItemDescription_Lookup = Application.VLookup(cmb_PayItemSearch.Value, PayItemList.Range("A1:I9999"), 3, False)
Unit_Lookup = Application.VLookup(cmb_PayItemSearch.Value, PayItemList.Range("A1:I9999"), 4, False)
FuelCode_Lookup = Application.VLookup(cmb_PayItemSearch.Value, PayItemList.Range("A1:I9999"), 5, False)
Participating_Lookup = Application.VLookup(cmb_PayItemSearch.Value, PayItemList.Range("A1:I9999"), 6, False)
Type_Lookup = Application.VLookup(cmb_PayItemSearch.Value, PayItemList.Range("A1:I9999"), 7, False)
SpecYear_Lookup = Application.VLookup(cmb_PayItemSearch.Value, PayItemList.Range("A1:I9999"), 8, False)
SortOrder_Lookup = Application.VLookup(cmb_PayItemSearch.Value, PayItemList.Range("A1:I9999"), 9, False)


    If IsError(PayItemNo_Lookup) Then
            txt_PayItemNo.Value = cmb_PayItemSearch.Value
            Else
            txt_PayItemNo.Value = Application.VLookup(cmb_PayItemSearch.Value, PayItemList.Range("A1:I9999"), 2, False)
    End If
    If IsError(PayItemDescription_Lookup) Then
            txt_PayItemDescription.Value = "Manually complete all the fields to add this new pay item.  The pay item is not a selection"
            Else
            txt_PayItemDescription.Value = Application.VLookup(cmb_PayItemSearch.Value, PayItemList.Range("A1:I9999"), 3, False)
    End If
    If IsError(Unit_Lookup) Then
            txt_Unit.Value = ""
            Else
            txt_Unit.Value = Application.VLookup(cmb_PayItemSearch.Value, PayItemList.Range("A1:I9999"), 4, False)
    End If
    If IsError(FuelCode_Lookup) Then
            cmb_FuelCode.Value = "NA"
            Else
            cmb_FuelCode.Value = Application.VLookup(cmb_PayItemSearch.Value, PayItemList.Range("A1:I9999"), 5, False)
    End If
    If IsError(Participating_Lookup) Then
            cmb_Participating.Value = "Yes"
            Else
            cmb_Participating.Value = Application.VLookup(cmb_PayItemSearch.Value, PayItemList.Range("A1:I9999"), 6, False)
    End If
    If IsError(Type_Lookup) Then
            cmb_Type.Value = ""
            Else
            cmb_Type.Value = Application.VLookup(cmb_PayItemSearch.Value, PayItemList.Range("A1:I9999"), 7, False)
    End If
        If IsError(SpecYear_Lookup) Then
            txt_SpecYear.Value = "1900"
            Else
            txt_SpecYear.Value = Application.VLookup(cmb_PayItemSearch.Value, PayItemList.Range("A1:I9999"), 8, False)
    End If
    If IsError(SortOrder_Lookup) Then
            txt_SortOrder.Value = "9999"
            Else
            txt_SortOrder.Value = Application.VLookup(cmb_PayItemSearch.Value, PayItemList.Range("A1:I9999"), 9, False)
    End If

    With Me
        txt_Subtotal.Value = txt_Quantity.Value * txt_UnitPrice.Value
    End With
End Sub
