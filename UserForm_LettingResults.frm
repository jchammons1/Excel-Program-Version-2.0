VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_LettingResults 
   Caption         =   "Bid Letting Results Form"
   ClientHeight    =   13170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18945
   OleObjectBlob   =   "UserForm_LettingResults.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_LettingResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Button_LettingResults2_Click()

End Sub

 Sub Button_SaveBidLettingResults_Click()
'
'
'
'******************************************************************
'
'
'
    BIDTAB.Select
    BIDTAB.Unprotect
    Range("DB3:DD15").Select
    Selection.Locked = False
    Range("DF12:DF13").Select
    Selection.Locked = False
    
    With Me
            'Contractor
            Range("DB3").Value = Me.txt_Bidder1.Value
            Range("DB4").Value = Me.txt_Bidder2.Value
            Range("DB5").Value = Me.txt_Bidder3.Value
            Range("DB6").Value = Me.txt_Bidder4.Value
            Range("DB7").Value = Me.txt_Bidder5.Value
            Range("DB8").Value = Me.txt_Bidder6.Value
            Range("DB9").Value = Me.txt_Bidder7.Value
            Range("DB10").Value = Me.txt_Bidder8.Value
            Range("DB11").Value = Me.txt_Bidder9.Value
            Range("DB12").Value = Me.txt_Bidder10.Value
            Range("DB13").Value = Me.txt_Bidder11.Value
            Range("DB14").Value = Me.txt_Bidder12.Value
            Range("DB15").Value = Me.txt_Bidder13.Value
            'Bid Amount
            Range("DD3").Value = Me.cmb_BidderAmount1.Value
            Range("DD4").Value = Me.cmb_BidderAmount2.Value
            Range("DD5").Value = Me.cmb_BidderAmount3.Value
            Range("DD6").Value = Me.cmb_BidderAmount4.Value
            Range("DD7").Value = Me.cmb_BidderAmount5.Value
            Range("DD8").Value = Me.cmb_BidderAmount6.Value
            Range("DD9").Value = Me.cmb_BidderAmount7.Value
            Range("DD10").Value = Me.cmb_BidderAmount8.Value
            Range("DD11").Value = Me.cmb_BidderAmount9.Value
            Range("DD12").Value = Me.cmb_BidderAmount10.Value
            Range("DD13").Value = Me.cmb_BidderAmount11.Value
            Range("DD14").Value = Me.cmb_BidderAmount12.Value
            Range("DD15").Value = Me.cmb_BidderAmount13.Value
            ' Award Status
            Range("DF12").Value = Me.cmb_AwardStatus.Value
            Range("DF13").Value = Me.txt_AwardDate.Value
            
    End With
        
    
    Range("DB3:DD15").Select
    Selection.Locked = True
    Range("DF12:DF13").Select
    Selection.Locked = True
    BIDTAB.Protect
    Range("DB3").Select
    
End Sub



Private Sub UserForm_Initialize()
    With Me
            'Contractor
            Me.txt_Bidder1.Value = Range("DB3").Value
            Me.txt_Bidder2.Value = Range("DB4").Value
            Me.txt_Bidder3.Value = Range("DB5").Value
            Me.txt_Bidder4.Value = Range("DB6").Value
            Me.txt_Bidder5.Value = Range("DB7").Value
            Me.txt_Bidder6.Value = Range("DB8").Value
            Me.txt_Bidder7.Value = Range("DB9").Value
            Me.txt_Bidder8.Value = Range("DB10").Value
            Me.txt_Bidder9.Value = Range("DB11").Value
            Me.txt_Bidder10.Value = Range("DB12").Value
            Me.txt_Bidder11.Value = Range("DB13").Value
            Me.txt_Bidder12.Value = Range("DB14").Value
            Me.txt_Bidder13.Value = Range("DB15").Value
            'Bid Amount
            Me.cmb_BidderAmount1.Value = Range("DD3").Value
            Me.cmb_BidderAmount2.Value = Range("DD4").Value
            Me.cmb_BidderAmount3.Value = Range("DD5").Value
            Me.cmb_BidderAmount4.Value = Range("DD6").Value
            Me.cmb_BidderAmount5.Value = Range("DD7").Value
            Me.cmb_BidderAmount6.Value = Range("DD8").Value
            Me.cmb_BidderAmount7.Value = Range("DD9").Value
            Me.cmb_BidderAmount8.Value = Range("DD10").Value
            Me.cmb_BidderAmount9.Value = Range("DD11").Value
            Me.cmb_BidderAmount10.Value = Range("DD12").Value
            Me.cmb_BidderAmount11.Value = Range("DD13").Value
            Me.cmb_BidderAmount12.Value = Range("DD14").Value
            Me.cmb_BidderAmount13.Value = Range("DD15").Value
            ' Award Status
            Me.cmb_AwardStatus.Value = Range("DF12").Value
            Me.txt_AwardDate.Value = Range("DF13").Value
    End With
    'Call FormatComboBox
End Sub
