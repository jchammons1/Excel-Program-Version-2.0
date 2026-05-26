Attribute VB_Name = "Module_ListBox"
Option Explicit
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'*********************************** Functions for the Database List Box on the Pay Item form*******************************************>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################

Public Function GetDatabase_Range() As Range
' Get the data range from the Database sheet
    Set GetDatabase_Range = Database.Range("A2").CurrentRegion
'Remove the header from the available data seletion.  Move the range down one (1) row and remove the last row
    Set GetDatabase_Range = GetDatabase_Range.Offset(1).Resize(GetDatabase_Range.Rows.Count - 1)
End Function


Public Sub DeleteRow_Database(ByVal row As Long)
'Delete entire row from the Database sheet
    Database.Range("A2").Offset(row).EntireRow.Delete
End Sub



'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'*********************************** Functions for the Type List Box  *******************************************************************>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################

Public Function GetType_Range() As Range
' Get the data range from the Pay Item Type Sheet
    Set GetType_Range = PAYITEMTYPE.Range("A2").CurrentRegion
'Remove the header from the available data seletion.  Move the range down one (1) row and remove the last row
    Set GetType_Range = GetType_Range.Offset(1).Resize(GetType_Range.Rows.Count - 1)
End Function

Public Function GetSuppType_Range() As Range
' Get the data range from the Pay Item Type Sheet
    Set GetSuppType_Range = PAYITEMTYPE_SUPPLEMENTAL.Range("A2").CurrentRegion
'Remove the header from the available data seletion.  Move the range down one (1) row and remove the last row
    Set GetSuppType_Range = GetSuppType_Range.Offset(1).Resize(GetSuppType_Range.Rows.Count - 1)
End Function


'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'*********************************** Functions for the CAD Database List Box*************************************************************>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################

Function GetCAD_Range() As Range
' Get the data range from the CAD Database sheet embedded in the Active CAD starting at Column OA
    Set GetCAD_Range = ActiveSheet.Range("OA2").CurrentRegion
'Remove the header from the available data seletion.  Move the range down one (1) row and remove the last row
    Set GetCAD_Range = GetCAD_Range.Offset(1).Resize(GetCAD_Range.Rows.Count - 1)
End Function

