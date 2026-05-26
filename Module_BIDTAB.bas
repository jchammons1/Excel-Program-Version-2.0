Attribute VB_Name = "Module_BIDTAB"
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################
'*********************************** This section BidTab preparation ********************************************************************>>>>>>>>>>>>>>>>>>
'*******************                                                                                                                    ###################
'*******************                                                                                                                    ###################

 Sub BidTab_Step1()
'Do Loop to unlock the cells and change the color to light yellow for visual data entry for the user.
' Step 1 is unlocking the cells the District Engineer will enter the data from the bid letting
' Set the message.  After changin the below Do while code to Select Case for the fuel code, the processing message is not needed because it is a lot faster
    UserForm_DASHBOARD.lblmessage3.Caption = "Processing, Step #1"
    DoEvents

    BIDTAB.Select
    BIDTAB.Range("DA21:DZ999").Clear
    Application.ScreenUpdating = False
    ActiveSheet.Unprotect
            row = 21
            column = 1
            endrow = CountPayItems

Do While row < endrow
        If Cells(row, column + 46).Value = "Yes" Or Cells(row, column + 46).Value = "No" Then
        
            Cells(row, column + 104).Select               'Bidder #1
                Call Format_BidTabColumnA
            Cells(row, column + 105).Select
                Cells(row, column + 105) = "=RC[-1]*RC[-103]"
                Call Format_BidTabColumnB

            Cells(row, column + 106).Select              'Bidder #2
                Call Format_BidTabColumnA
            Cells(row, column + 107).Select
                Cells(row, column + 107) = "=RC[-1]*RC[-105]"
                Call Format_BidTabColumnB
                
            Cells(row, column + 108).Select             'Bidder #3
                Call Format_BidTabColumnA
            Cells(row, column + 109).Select
                Cells(row, column + 109) = "=RC[-1]*RC[-107]"
                Call Format_BidTabColumnB

            Cells(row, column + 110).Select              'Bidder #4
                Call Format_BidTabColumnA
            Cells(row, column + 111).Select
                Cells(row, column + 111) = "=RC[-1]*RC[-109]"
                Call Format_BidTabColumnB
                           
            Cells(row, column + 112).Select              'Bidder #5
                Call Format_BidTabColumnA
            Cells(row, column + 113).Select
                Cells(row, column + 113) = "=RC[-1]*RC[-111]"
                Call Format_BidTabColumnB
                
            Cells(row, column + 114).Select              'Bidder #6
                Call Format_BidTabColumnA
            Cells(row, column + 115).Select
                Cells(row, column + 115) = "=RC[-1]*RC[-113]"
                Call Format_BidTabColumnB
                            
            Cells(row, column + 116).Select             'Bidder #7
                Call Format_BidTabColumnA
            Cells(row, column + 117).Select
                Cells(row, column + 117) = "=RC[-1]*RC[-115]"
                Call Format_BidTabColumnB
                           
            Cells(row, column + 118).Select             'Bidder #8
                Call Format_BidTabColumnA
            Cells(row, column + 119).Select
                Cells(row, column + 119) = "=RC[-1]*RC[-117]"
                Call Format_BidTabColumnB
                            
            Cells(row, column + 120).Select              'Bidder #9
                Call Format_BidTabColumnA
            Cells(row, column + 121).Select
                Cells(row, column + 121) = "=RC[-1]*RC[-119]"
                Call Format_BidTabColumnB
                
            Cells(row, column + 122).Select             'Bidder #10
                Call Format_BidTabColumnA
            Cells(row, column + 123).Select
                Cells(row, column + 123) = "=RC[-1]*RC[-121]"
                Call Format_BidTabColumnB
                            
            Cells(row, column + 124).Select             'Bidder #11
                Call Format_BidTabColumnA
            Cells(row, column + 125).Select
                Cells(row, column + 125) = "=RC[-1]*RC[-123]"
                Call Format_BidTabColumnB
                          
            Cells(row, column + 126).Select             'Bidder #12
                Call Format_BidTabColumnA
            Cells(row, column + 127).Select
                Cells(row, column + 127) = "=RC[-1]*RC[-125]"
                Call Format_BidTabColumnB
                           
            Cells(row, column + 128).Select             'Bidder #13
                Call Format_BidTabColumnA
            Cells(row, column + 129).Select
                Cells(row, column + 129) = "=RC[-1]*RC[-127]"
                Call Format_BidTabColumnB
        End If
        
        
        If Cells(row, column + 1).Value = "PROJECT PARTICIPATING TOTAL" Then
        
            Cells(row, column + 105).Select             'Bidder #1
            Cells(row, column + 105) = "=IFERROR(SUMIFS(C,C47,""Yes""),"""")"
            Call Format_BidTabColumnB
            
            Cells(row, column + 107).Select              'Bidder #2
                Cells(row, column + 107) = "=IFERROR(SUMIFS(C,C47,""Yes""),"""")"
                Call Format_BidTabColumnB
                
            Cells(row, column + 109).Select             'Bidder #3
                Cells(row, column + 109) = "=IFERROR(SUMIFS(C,C47,""Yes""),"""")"
                Call Format_BidTabColumnB

            Cells(row, column + 111).Select              'Bidder #4
                Cells(row, column + 111) = "=IFERROR(SUMIFS(C,C47,""Yes""),"""")"
                Call Format_BidTabColumnB
                           
            Cells(row, column + 113).Select              'Bidder #5
                Cells(row, column + 113) = "=IFERROR(SUMIFS(C,C47,""Yes""),"""")"
                Call Format_BidTabColumnB
                
            Cells(row, column + 115).Select              'Bidder #6
                Cells(row, column + 115) = "=IFERROR(SUMIFS(C,C47,""Yes""),"""")"
                Call Format_BidTabColumnB
                            
            Cells(row, column + 117).Select             'Bidder #7
                Cells(row, column + 117) = "=IFERROR(SUMIFS(C,C47,""Yes""),"""")"
                Call Format_BidTabColumnB
                           
            Cells(row, column + 119).Select             'Bidder #8
                Cells(row, column + 119) = "=IFERROR(SUMIFS(C,C47,""Yes""),"""")"
                Call Format_BidTabColumnB
                            
            Cells(row, column + 121).Select              'Bidder #9
                Cells(row, column + 121) = "=IFERROR(SUMIFS(C,C47,""Yes""),"""")"
                Call Format_BidTabColumnB

            Cells(row, column + 123).Select             'Bidder #10
                Cells(row, column + 123) = "=IFERROR(SUMIFS(C,C47,""Yes""),"""")"
                Call Format_BidTabColumnB
                            
            Cells(row, column + 125).Select             'Bidder #11
                Cells(row, column + 125) = "=IFERROR(SUMIFS(C,C47,""Yes""),"""")"
                Call Format_BidTabColumnB
                          
            Cells(row, column + 127).Select             'Bidder #12
                Cells(row, column + 127) = "=IFERROR(SUMIFS(C,C47,""Yes""),"""")"
                Call Format_BidTabColumnB
                           
            Cells(row, column + 129).Select             'Bidder #13
                Cells(row, column + 129) = "=IFERROR(SUMIFS(C,C47,""Yes""),"""")"
                Call Format_BidTabColumnB
        End If
        
        If Cells(row, column + 1).Value = "PROJECT NON-PARTICIPATING TOTAL" Then
        
            Cells(row, column + 105).Select             'Bidder #1
            Cells(row, column + 105) = "=IFERROR(SUMIFS(C,C47,""No""),"""")"
            Call Format_BidTabColumnB
            
            Cells(row, column + 107).Select              'Bidder #2
                Cells(row, column + 107) = "=IFERROR(SUMIFS(C,C47,""No""),"""")"
                Call Format_BidTabColumnB
                
            Cells(row, column + 109).Select             'Bidder #3
                Cells(row, column + 109) = "=IFERROR(SUMIFS(C,C47,""No""),"""")"
                Call Format_BidTabColumnB

            Cells(row, column + 111).Select              'Bidder #4
                Cells(row, column + 111) = "=IFERROR(SUMIFS(C,C47,""No""),"""")"
                Call Format_BidTabColumnB
                           
            Cells(row, column + 113).Select              'Bidder #5
                Cells(row, column + 113) = "=IFERROR(SUMIFS(C,C47,""No""),"""")"
                Call Format_BidTabColumnB
                
            Cells(row, column + 115).Select              'Bidder #6
                Cells(row, column + 115) = "=IFERROR(SUMIFS(C,C47,""No""),"""")"
                Call Format_BidTabColumnB
                            
            Cells(row, column + 117).Select             'Bidder #7
                Cells(row, column + 117) = "=IFERROR(SUMIFS(C,C47,""No""),"""")"
                Call Format_BidTabColumnB
                           
            Cells(row, column + 119).Select             'Bidder #8
                Cells(row, column + 119) = "=IFERROR(SUMIFS(C,C47,""No""),"""")"
                Call Format_BidTabColumnB
                            
            Cells(row, column + 121).Select              'Bidder #9
                Cells(row, column + 121) = "=IFERROR(SUMIFS(C,C47,""No""),"""")"
                Call Format_BidTabColumnB

            Cells(row, column + 123).Select             'Bidder #10
                Cells(row, column + 123) = "=IFERROR(SUMIFS(C,C47,""No""),"""")"
                Call Format_BidTabColumnB
                            
            Cells(row, column + 125).Select             'Bidder #11
                Cells(row, column + 125) = "=IFERROR(SUMIFS(C,C47,""No""),"""")"
                Call Format_BidTabColumnB
                          
            Cells(row, column + 127).Select             'Bidder #12
                Cells(row, column + 127) = "=IFERROR(SUMIFS(C,C47,""No""),"""")"
                Call Format_BidTabColumnB
                           
            Cells(row, column + 129).Select             'Bidder #13
                Cells(row, column + 129) = "=IFERROR(SUMIFS(C,C47,""No""),"""")"
                Call Format_BidTabColumnB
        End If
        
        If Cells(row, column + 1).Value = "PROJECT PARTICIPATING AND NON-PARTICIPATING TOTAL" Then
        
            Cells(row, column + 105).Select             'Bidder #1
            Cells(row, column + 105) = "=SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No"")"
            Call Format_BidTabColumnB
            
            Cells(row, column + 107).Select              'Bidder #2
                Cells(row, column + 107) = "=SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No"")"
                Call Format_BidTabColumnB
                
            Cells(row, column + 109).Select             'Bidder #3
                Cells(row, column + 109) = "=SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No"")"
                Call Format_BidTabColumnB

            Cells(row, column + 111).Select              'Bidder #4
                Cells(row, column + 111) = "=SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No"")"
                Call Format_BidTabColumnB
                           
            Cells(row, column + 113).Select              'Bidder #5
                Cells(row, column + 113) = "=SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No"")"
                Call Format_BidTabColumnB
                
            Cells(row, column + 115).Select              'Bidder #6
                Cells(row, column + 115) = "=SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No"")"
                Call Format_BidTabColumnB
                            
            Cells(row, column + 117).Select             'Bidder #7
                Cells(row, column + 117) = "=SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No"")"
                Call Format_BidTabColumnB
                           
            Cells(row, column + 119).Select             'Bidder #8
                Cells(row, column + 119) = "=SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No"")"
                Call Format_BidTabColumnB
                            
            Cells(row, column + 121).Select              'Bidder #9
                Cells(row, column + 121) = "=SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No"")"
                Call Format_BidTabColumnB

            Cells(row, column + 123).Select             'Bidder #10
                Cells(row, column + 123) = "=SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No"")"
                Call Format_BidTabColumnB
                            
            Cells(row, column + 125).Select             'Bidder #11
                Cells(row, column + 125) = "=SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No"")"
                Call Format_BidTabColumnB
                          
            Cells(row, column + 127).Select             'Bidder #12
                Cells(row, column + 127) = "=SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No"")"
                Call Format_BidTabColumnB
                           
            Cells(row, column + 129).Select             'Bidder #13
                Cells(row, column + 129) = "=SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No"")"
                Call Format_BidTabColumnB
        End If
        
        
        If Cells(row, column + 1).Value = "PERCENT OVER / UNDER PARTICIPATING OSARC ESTIMATE" Then
        
            Cells(row, column + 105).Select             'Bidder #1
                Cells(row, column + 105) = "=IFERROR((SUMIFS(C,C47,""Yes"")-SUMIFS(C6,C47,""Yes""))/SUMIFS(C6,C47,""Yes""),"""")"
                Call Format_BidTabColumnPercent
            
            Cells(row, column + 107).Select              'Bidder #2
                Cells(row, column + 107) = "=IFERROR((SUMIFS(C,C47,""Yes"")-SUMIFS(C6,C47,""Yes""))/SUMIFS(C6,C47,""Yes""),"""")"
                Call Format_BidTabColumnPercent
                
            Cells(row, column + 109).Select             'Bidder #3
                Cells(row, column + 109) = "=IFERROR((SUMIFS(C,C47,""Yes"")-SUMIFS(C6,C47,""Yes""))/SUMIFS(C6,C47,""Yes""),"""")"
                Call Format_BidTabColumnPercent

            Cells(row, column + 111).Select              'Bidder #4
                Cells(row, column + 111) = "=IFERROR((SUMIFS(C,C47,""Yes"")-SUMIFS(C6,C47,""Yes""))/SUMIFS(C6,C47,""Yes""),"""")"
                Call Format_BidTabColumnPercent
                           
            Cells(row, column + 113).Select              'Bidder #5
                Cells(row, column + 113) = "=IFERROR((SUMIFS(C,C47,""Yes"")-SUMIFS(C6,C47,""Yes""))/SUMIFS(C6,C47,""Yes""),"""")"
                Call Format_BidTabColumnPercent
                
            Cells(row, column + 115).Select              'Bidder #6
                Cells(row, column + 115) = "=IFERROR((SUMIFS(C,C47,""Yes"")-SUMIFS(C6,C47,""Yes""))/SUMIFS(C6,C47,""Yes""),"""")"
                Call Format_BidTabColumnPercent
                            
            Cells(row, column + 117).Select             'Bidder #7
                Cells(row, column + 117) = "=IFERROR((SUMIFS(C,C47,""Yes"")-SUMIFS(C6,C47,""Yes""))/SUMIFS(C6,C47,""Yes""),"""")"
                Call Format_BidTabColumnPercent
                           
            Cells(row, column + 119).Select             'Bidder #8
                Cells(row, column + 119) = "=IFERROR((SUMIFS(C,C47,""Yes"")-SUMIFS(C6,C47,""Yes""))/SUMIFS(C6,C47,""Yes""),"""")"
                Call Format_BidTabColumnPercent
                            
            Cells(row, column + 121).Select              'Bidder #9
                Cells(row, column + 121) = "=IFERROR((SUMIFS(C,C47,""Yes"")-SUMIFS(C6,C47,""Yes""))/SUMIFS(C6,C47,""Yes""),"""")"
                Call Format_BidTabColumnPercent

            Cells(row, column + 123).Select             'Bidder #10
                Cells(row, column + 123) = "=IFERROR((SUMIFS(C,C47,""Yes"")-SUMIFS(C6,C47,""Yes""))/SUMIFS(C6,C47,""Yes""),"""")"
                Call Format_BidTabColumnPercent
                            
            Cells(row, column + 125).Select             'Bidder #11
                Cells(row, column + 125) = "=IFERROR((SUMIFS(C,C47,""Yes"")-SUMIFS(C6,C47,""Yes""))/SUMIFS(C6,C47,""Yes""),"""")"
                Call Format_BidTabColumnPercent
                          
            Cells(row, column + 127).Select             'Bidder #12
                Cells(row, column + 127) = "=IFERROR((SUMIFS(C,C47,""Yes"")-SUMIFS(C6,C47,""Yes""))/SUMIFS(C6,C47,""Yes""),"""")"
                Call Format_BidTabColumnPercent
                           
            Cells(row, column + 129).Select             'Bidder #13
                Cells(row, column + 129) = "=IFERROR((SUMIFS(C,C47,""Yes"")-SUMIFS(C6,C47,""Yes""))/SUMIFS(C6,C47,""Yes""),"""")"
                Call Format_BidTabColumnPercent
        End If
        
        
        If Cells(row, column + 1).Value = "PERCENT OVER / UNDER NON-PARTICIPATING OSARC ESTIMATE" Then
        
            Cells(row, column + 105).Select             'Bidder #1
                Cells(row, column + 105) = "=IFERROR((SUMIFS(C,C47,""No"")-SUMIFS(C6,C47,""No""))/SUMIFS(C6,C47,""No""),"""")"
                Call Format_BidTabColumnPercent
            
            Cells(row, column + 107).Select              'Bidder #2
                Cells(row, column + 107) = "=IFERROR((SUMIFS(C,C47,""No"")-SUMIFS(C6,C47,""No""))/SUMIFS(C6,C47,""No""),"""")"
                Call Format_BidTabColumnPercent
                
            Cells(row, column + 109).Select             'Bidder #3
                Cells(row, column + 109) = "=IFERROR((SUMIFS(C,C47,""No"")-SUMIFS(C6,C47,""No""))/SUMIFS(C6,C47,""No""),"""")"
                Call Format_BidTabColumnPercent

            Cells(row, column + 111).Select              'Bidder #4
                Cells(row, column + 111) = "=IFERROR((SUMIFS(C,C47,""No"")-SUMIFS(C6,C47,""No""))/SUMIFS(C6,C47,""No""),"""")"
                Call Format_BidTabColumnPercent
                           
            Cells(row, column + 113).Select              'Bidder #5
                Cells(row, column + 113) = "=IFERROR((SUMIFS(C,C47,""No"")-SUMIFS(C6,C47,""No""))/SUMIFS(C6,C47,""No""),"""")"
                Call Format_BidTabColumnPercent
                
            Cells(row, column + 115).Select              'Bidder #6
                Cells(row, column + 115) = "=IFERROR((SUMIFS(C,C47,""No"")-SUMIFS(C6,C47,""No""))/SUMIFS(C6,C47,""No""),"""")"
                Call Format_BidTabColumnPercent
                            
            Cells(row, column + 117).Select             'Bidder #7
                Cells(row, column + 117) = "=IFERROR((SUMIFS(C,C47,""No"")-SUMIFS(C6,C47,""No""))/SUMIFS(C6,C47,""No""),"""")"
                Call Format_BidTabColumnPercent
                           
            Cells(row, column + 119).Select             'Bidder #8
                Cells(row, column + 119) = "=IFERROR((SUMIFS(C,C47,""No"")-SUMIFS(C6,C47,""No""))/SUMIFS(C6,C47,""No""),"""")"
                Call Format_BidTabColumnPercent
                            
            Cells(row, column + 121).Select              'Bidder #9
                Cells(row, column + 121) = "=IFERROR((SUMIFS(C,C47,""No"")-SUMIFS(C6,C47,""No""))/SUMIFS(C6,C47,""No""),"""")"
                Call Format_BidTabColumnPercent

            Cells(row, column + 123).Select             'Bidder #10
                Cells(row, column + 123) = "=IFERROR((SUMIFS(C,C47,""No"")-SUMIFS(C6,C47,""No""))/SUMIFS(C6,C47,""No""),"""")"
                Call Format_BidTabColumnPercent
                            
            Cells(row, column + 125).Select             'Bidder #11
                Cells(row, column + 125) = "=IFERROR((SUMIFS(C,C47,""No"")-SUMIFS(C6,C47,""No""))/SUMIFS(C6,C47,""No""),"""")"
                Call Format_BidTabColumnPercent
                          
            Cells(row, column + 127).Select             'Bidder #12
                Cells(row, column + 127) = "=IFERROR((SUMIFS(C,C47,""No"")-SUMIFS(C6,C47,""No""))/SUMIFS(C6,C47,""No""),"""")"
                Call Format_BidTabColumnPercent
                           
            Cells(row, column + 129).Select             'Bidder #13
                Cells(row, column + 129) = "=IFERROR((SUMIFS(C,C47,""No"")-SUMIFS(C6,C47,""No""))/SUMIFS(C6,C47,""No""),"""")"
                Call Format_BidTabColumnPercent
        End If
        
        
        If Cells(row, column + 1).Value = "PERCENT OVER / UNDER PARTICIPATING AND NON-PARTICIPATING OSARC ESTIMATE" Then
            Cells(row, column + 105).Select             'Bidder #1
                Cells(row, column + 105) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-((SUMIFS(C6,C47,""Yes"")+SUMIFS(C6,C47,""No""))))/((SUMIFS(C6,C47,""No"")+SUMIFS(C6,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent
            
            Cells(row, column + 107).Select              'Bidder #2
                Cells(row, column + 107) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-((SUMIFS(C6,C47,""Yes"")+SUMIFS(C6,C47,""No""))))/((SUMIFS(C6,C47,""No"")+SUMIFS(C6,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent
                
            Cells(row, column + 109).Select             'Bidder #3
                Cells(row, column + 109) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-((SUMIFS(C6,C47,""Yes"")+SUMIFS(C6,C47,""No""))))/((SUMIFS(C6,C47,""No"")+SUMIFS(C6,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent

            Cells(row, column + 111).Select              'Bidder #4
                Cells(row, column + 111) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-((SUMIFS(C6,C47,""Yes"")+SUMIFS(C6,C47,""No""))))/((SUMIFS(C6,C47,""No"")+SUMIFS(C6,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent
                           
            Cells(row, column + 113).Select              'Bidder #5
                Cells(row, column + 113) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-((SUMIFS(C6,C47,""Yes"")+SUMIFS(C6,C47,""No""))))/((SUMIFS(C6,C47,""No"")+SUMIFS(C6,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent
                
            Cells(row, column + 115).Select              'Bidder #6
                Cells(row, column + 115) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-((SUMIFS(C6,C47,""Yes"")+SUMIFS(C6,C47,""No""))))/((SUMIFS(C6,C47,""No"")+SUMIFS(C6,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent
                            
            Cells(row, column + 117).Select             'Bidder #7
                Cells(row, column + 117) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-((SUMIFS(C6,C47,""Yes"")+SUMIFS(C6,C47,""No""))))/((SUMIFS(C6,C47,""No"")+SUMIFS(C6,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent
                           
            Cells(row, column + 119).Select             'Bidder #8
                Cells(row, column + 119) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-((SUMIFS(C6,C47,""Yes"")+SUMIFS(C6,C47,""No""))))/((SUMIFS(C6,C47,""No"")+SUMIFS(C6,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent
                            
            Cells(row, column + 121).Select              'Bidder #9
                Cells(row, column + 121) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-((SUMIFS(C6,C47,""Yes"")+SUMIFS(C6,C47,""No""))))/((SUMIFS(C6,C47,""No"")+SUMIFS(C6,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent

            Cells(row, column + 123).Select             'Bidder #10
                Cells(row, column + 123) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-((SUMIFS(C6,C47,""Yes"")+SUMIFS(C6,C47,""No""))))/((SUMIFS(C6,C47,""No"")+SUMIFS(C6,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent
                            
            Cells(row, column + 125).Select             'Bidder #11
                Cells(row, column + 125) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-((SUMIFS(C6,C47,""Yes"")+SUMIFS(C6,C47,""No""))))/((SUMIFS(C6,C47,""No"")+SUMIFS(C6,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent
                          
            Cells(row, column + 127).Select             'Bidder #12
                Cells(row, column + 127) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-((SUMIFS(C6,C47,""Yes"")+SUMIFS(C6,C47,""No""))))/((SUMIFS(C6,C47,""No"")+SUMIFS(C6,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent
                           
            Cells(row, column + 129).Select             'Bidder #13
                Cells(row, column + 129) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-((SUMIFS(C6,C47,""Yes"")+SUMIFS(C6,C47,""No""))))/((SUMIFS(C6,C47,""No"")+SUMIFS(C6,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent
        End If
        
        If Cells(row, column + 1).Value = "DIFFERENCE PARTICIPATING AND NON-PARTICIPATING BIDDER AMOUNT" Then
            Cells(row, column + 105).Select             'Bidder #1
                Cells(row, column + 105) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C105)),"""")"
                Call Format_BidTabColumnB
            
            Cells(row, column + 107).Select              'Bidder #2
                Cells(row, column + 107) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C107)),"""")"
                Call Format_BidTabColumnB
                
            Cells(row, column + 109).Select             'Bidder #3
                Cells(row, column + 109) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C109)),"""")"
                Call Format_BidTabColumnB

            Cells(row, column + 111).Select              'Bidder #4
                Cells(row, column + 111) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C111)),"""")"
                Call Format_BidTabColumnB
                           
            Cells(row, column + 113).Select              'Bidder #5
                Cells(row, column + 113) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C113)),"""")"
                Call Format_BidTabColumnB
                
            Cells(row, column + 115).Select              'Bidder #6
                Cells(row, column + 115) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C115)),"""")"
                Call Format_BidTabColumnB
                            
            Cells(row, column + 117).Select             'Bidder #7
                Cells(row, column + 117) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C117)),"""")"
                Call Format_BidTabColumnB
                           
            Cells(row, column + 119).Select             'Bidder #8
                Cells(row, column + 119) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C119)),"""")"
                Call Format_BidTabColumnB
                            
            Cells(row, column + 121).Select              'Bidder #9
                Cells(row, column + 121) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C121)),"""")"
                Call Format_BidTabColumnB

            Cells(row, column + 123).Select             'Bidder #10
                Cells(row, column + 123) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C123)),"""")"
                Call Format_BidTabColumnB
                            
            Cells(row, column + 125).Select             'Bidder #11
                Cells(row, column + 125) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C125)),"""")"
                Call Format_BidTabColumnB
                          
            Cells(row, column + 127).Select             'Bidder #12
                Cells(row, column + 127) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C127)),"""")"
                Call Format_BidTabColumnB
                           
            Cells(row, column + 129).Select             'Bidder #13
                Cells(row, column + 129) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C129)),"""")"
                Call Format_BidTabColumnB
        End If
                                           
        If Cells(row, column + 1).Value = "PERCENT OVER / UNDER PARTICIPATING AND NON-PARTICIPATING BIDDER AMOUNT" Then
            Cells(row, column + 105).Select             'Bidder #1
                Cells(row, column + 105) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C105)/((SUMIFS(C,C47,""No"")+SUMIFS(C,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent
            
            Cells(row, column + 107).Select              'Bidder #2
                Cells(row, column + 107) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C107)/((SUMIFS(C,C47,""No"")+SUMIFS(C,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent
                
            Cells(row, column + 109).Select             'Bidder #3
                Cells(row, column + 109) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C109)/((SUMIFS(C,C47,""No"")+SUMIFS(C,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent

            Cells(row, column + 111).Select              'Bidder #4
                Cells(row, column + 111) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C111)/((SUMIFS(C,C47,""No"")+SUMIFS(C,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent
                           
            Cells(row, column + 113).Select              'Bidder #5
                Cells(row, column + 113) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C113)/((SUMIFS(C,C47,""No"")+SUMIFS(C,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent
                
            Cells(row, column + 115).Select              'Bidder #6
                Cells(row, column + 115) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C115)/((SUMIFS(C,C47,""No"")+SUMIFS(C,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent
                            
            Cells(row, column + 117).Select             'Bidder #7
                Cells(row, column + 117) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C117)/((SUMIFS(C,C47,""No"")+SUMIFS(C,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent
                           
            Cells(row, column + 119).Select             'Bidder #8
                Cells(row, column + 119) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C119)/((SUMIFS(C,C47,""No"")+SUMIFS(C,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent
                            
            Cells(row, column + 121).Select              'Bidder #9
                Cells(row, column + 121) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C121)/((SUMIFS(C,C47,""No"")+SUMIFS(C,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent

            Cells(row, column + 123).Select             'Bidder #10
                Cells(row, column + 123) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C123)/((SUMIFS(C,C47,""No"")+SUMIFS(C,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent
                            
            Cells(row, column + 125).Select             'Bidder #11
                Cells(row, column + 125) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C125)/((SUMIFS(C,C47,""No"")+SUMIFS(C,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent
                          
            Cells(row, column + 127).Select             'Bidder #12
                Cells(row, column + 127) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C127)/((SUMIFS(C,C47,""No"")+SUMIFS(C,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent
                           
            Cells(row, column + 129).Select             'Bidder #13
                Cells(row, column + 129) = "=IFERROR((((SUMIFS(C,C47,""Yes"")+SUMIFS(C,C47,""No""))-R17C129)/((SUMIFS(C,C47,""No"")+SUMIFS(C,C47,""Yes"")))),"""")"
                Call Format_BidTabColumnPercent
                

          ' Format these boxes so the DE can jump from top to bottom for easier reading.  Originally it was jumping from the yellow unit cost which kind of
          ' hides values they want to see as they are typing
            Cells(row, column + 104).Select
            Format_BidTabBoxes
            Cells(row, column + 106).Select
            Format_BidTabBoxes
            Cells(row, column + 108).Select
            Format_BidTabBoxes
            Cells(row, column + 110).Select
            Format_BidTabBoxes
            Cells(row, column + 112).Select
            Format_BidTabBoxes
            Cells(row, column + 114).Select
            Format_BidTabBoxes
            Cells(row, column + 116).Select
            Format_BidTabBoxes
            Cells(row, column + 118).Select
            Format_BidTabBoxes
            Cells(row, column + 120).Select
            Format_BidTabBoxes
            Cells(row, column + 122).Select
            Format_BidTabBoxes
            Cells(row, column + 124).Select
            Format_BidTabBoxes
            Cells(row, column + 126).Select
            Format_BidTabBoxes
            Cells(row, column + 128).Select
            Format_BidTabBoxes
                
                
        End If
        
        
            row = row + 1
    Loop

    BIDTAB.Select
    Range("DB3").Select
    
End Sub
'
'
'*************************************************************** BID TAB STEP 1 ENTER RESULTS END ***********************************************************************************************
'



'
'
'
'****************************************************************   Bid Tab Step #2 Start  ########################################################################################################
'
'
 Sub BidTab_Step2()
'Do Loop to find each type and subtotal the bid amounts
    'Clear caption
    UserForm_DASHBOARD.lblmessage3.Caption = "Process Step #2"
    DoEvents

    BIDTAB.Select
    Application.ScreenUpdating = False
    ActiveSheet.Unprotect
    PAYITEMTYPE.Visible = True
    PAYITEMTYPE.Select
    Dim RowType As Range                            'Declare a variable for the Pay Item Type and the range of data from the Data Validation sheet
    For Each RowType In Range("A2:A55")            ' Maximum of 100 types the user can add to the program. 4 default types. 50 User Defined and 50 supplemental types.
                                                    
                                                    ' The original estimates only go to A2 and A55
        If IsEmpty(RowType) = False Then
    

            BIDTAB.Select
            row = 21
            column = 1
            endrow = CountPayItems

            Do While row < endrow
                   
 'Calculating the sums for type of participating Erosion Control pay items'
        If Cells(row, column + 46).Value <> "Yes" And Cells(row, column + 1).Value = "Subtotal Participating " & RowType.Value & " Items" Then
            Cells(row, column + 105).Select                         ' Bidder #1
            Cells(row, column + 105) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""Yes"")"
            Format_BidTabColumnB

            Cells(row, column + 107).Select                         ' Bidder #2
            Cells(row, column + 107) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""Yes"")"
            Format_BidTabColumnB
            
            Cells(row, column + 109).Select                         ' Bidder #3
            Cells(row, column + 109) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""Yes"")"
            Format_BidTabColumnB
            
            Cells(row, column + 111).Select                         ' Bidder #4
            Cells(row, column + 111) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""Yes"")"
            Format_BidTabColumnB
            
            Cells(row, column + 113).Select                         ' Bidder #5
            Cells(row, column + 113) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""Yes"")"
            Format_BidTabColumnB
            
            Cells(row, column + 115).Select                         ' Bidder #6
            Cells(row, column + 115) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""Yes"")"
            Format_BidTabColumnB
            
            Cells(row, column + 117).Select                         ' Bidder #7
            Cells(row, column + 117) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""Yes"")"
            Format_BidTabColumnB
            
            Cells(row, column + 119).Select                         ' Bidder #8
            Cells(row, column + 119) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""Yes"")"
            Format_BidTabColumnB
            
            Cells(row, column + 121).Select                         ' Bidder #9
            Cells(row, column + 121) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""Yes"")"
            Format_BidTabColumnB
            
            Cells(row, column + 123).Select                         ' Bidder #10
            Cells(row, column + 123) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""Yes"")"
            Format_BidTabColumnB
            
            Cells(row, column + 125).Select                         ' Bidder #11
            Cells(row, column + 125) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""Yes"")"
            Format_BidTabColumnB
            
            Cells(row, column + 127).Select                         ' Bidder #12
            Cells(row, column + 127) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""Yes"")"
            Format_BidTabColumnB
            
            Cells(row, column + 129).Select                         ' Bidder #13
            Cells(row, column + 129) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""Yes"")"
            Format_BidTabColumnB
            
            End If
            
        If Cells(row, column + 46).Value <> "No" And Cells(row, column + 1).Value = "Subtotal NON-Participating " & RowType.Value & " Items" Then
            Cells(row, column + 105).Select                         ' Bidder #1
            Cells(row, column + 105) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""No"")"
            Format_BidTabColumnB

            Cells(row, column + 107).Select                         ' Bidder #2
            Cells(row, column + 107) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""No"")"
            Format_BidTabColumnB
            
            Cells(row, column + 109).Select                         ' Bidder #3
            Cells(row, column + 109) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""No"")"
            Format_BidTabColumnB
            
            Cells(row, column + 111).Select                         ' Bidder #4
            Cells(row, column + 111) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""No"")"
            Format_BidTabColumnB
            
            Cells(row, column + 113).Select                         ' Bidder #5
            Cells(row, column + 113) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""No"")"
            Format_BidTabColumnB
            
            Cells(row, column + 115).Select                         ' Bidder #6
            Cells(row, column + 115) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""No"")"
            Format_BidTabColumnB
            
            Cells(row, column + 117).Select                         ' Bidder #7
            Cells(row, column + 117) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""No"")"
            Format_BidTabColumnB
            
            Cells(row, column + 119).Select                         ' Bidder #8
            Cells(row, column + 119) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""No"")"
            Format_BidTabColumnB
            
            Cells(row, column + 121).Select                         ' Bidder #9
            Cells(row, column + 121) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""No"")"
            Format_BidTabColumnB
            
            Cells(row, column + 123).Select                         ' Bidder #10
            Cells(row, column + 123) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""No"")"
            Format_BidTabColumnB
            
            Cells(row, column + 125).Select                         ' Bidder #11
            Cells(row, column + 125) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""No"")"
            Format_BidTabColumnB
            
            Cells(row, column + 127).Select                         ' Bidder #12
            Cells(row, column + 127) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""No"")"
            Format_BidTabColumnB
            
            Cells(row, column + 129).Select                         ' Bidder #13
            Cells(row, column + 129) = "=SUMIFS(C,C46," & _
                         Chr(34) & _
                         RowType.Value & _
                         Chr(34) & _
                         ",C47,""No"")"
            Format_BidTabColumnB
            
            End If
            
        ' Show which row is processing
        UserForm_DASHBOARD.lblmessage3.Caption = "Processing: " & row & " of " & endrow
        DoEvents
            row = row + 1
        Loop
    End If
        
 Next RowType
End Sub


'
'
'
'Reference Number creation
'
'

 Sub ReferenceNo()
    
    BIDTAB.Select
    BIDTAB.Range("B4").Select
        row = 21
        column = 1
        endrow = 100

            Do While row < endrow
                
                If Cells(row, column + 46).Value = "Yes" Or Cells(row, column + 46).Value = "No" Then
                    Dim i As Integer                        ' Adding Ref.No. to Column A.  Numbering the column to check against the Form 902
                    i = i + 1
                    Cells(row, column + 103) = i
                    Cells(row, column + 103).Select
                    Selection.ShrinkToFit = True
                    Selection.Font.Size = 14
                    Selection.HorizontalAlignment = xlCenter
                    Selection.VerticalAlignment = xlCenter
                    Selection.ShrinkToFit = True
                    ' Font Color
                    Selection.Font.ThemeColor = xlThemeColorLight1
                    Selection.Font.TintAndShade = 0
                
                    'Cell color
                    Selection.Interior.Pattern = xlSolid
                    Selection.Interior.PatternColorIndex = xlAutomatic
                    Selection.Interior.ThemeColor = xlThemeColorDark1
                    Selection.Interior.TintAndShade = -0.249977111117893
                    Selection.Interior.PatternTintAndShade = 0
            
                    'Cell border
                    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
                    Selection.Borders(xlEdgeLeft).ThemeColor = 1
                    Selection.Borders(xlEdgeLeft).TintAndShade = 0
                    Selection.Borders(xlEdgeLeft).Weight = xlMedium
            
                    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
                    Selection.Borders(xlEdgeTop).ThemeColor = 1
                    Selection.Borders(xlEdgeTop).TintAndShade = 0
                    Selection.Borders(xlEdgeTop).Weight = xlMedium
            
                    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
                    Selection.Borders(xlEdgeBottom).ThemeColor = 1
                    Selection.Borders(xlEdgeBottom).TintAndShade = 0
                    Selection.Borders(xlEdgeBottom).Weight = xlMedium
                    
                    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
                    Selection.Borders(xlEdgeRight).ThemeColor = 1
                    Selection.Borders(xlEdgeRight).TintAndShade = 0
                    Selection.Borders(xlEdgeRight).Weight = xlMedium

                End If
                
        row = row + 1
                
        Loop
    Columns("CZ:CZ").Select
    Columns("CZ:CZ").EntireColumn.AutoFit
End Sub


