Sub Ticker_and_Total_Stock_Volume()

'Headings
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Stock Volume"
    Range("K1").Value = "Yearly change"
    Range("L1").Value = "Percent change"
    
    'Total Stock Volume for each ticker
    Dim Total As Double
    Dim closingprice As Double
    Dim openingprice As Double
    Dim j As Long
    
    j = 2
    
    
    
'FOR LOOP - TICKER
'Find the last non-blank cell in the the ticker column
    rowcount = Cells(Rows.Count, "A").End(xlUp).Row
        
'Loop will start
    For i = 2 To rowcount
    
        openingprice = Cells(j, 3).Value
        
'If the Ticker changes do the following
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
'Print New Ticker Letter in column I row 2
            Range("I" & 2 + j).Value = Cells(i, 1).Value                                 '(Cells(i,1).Value = Ticker column)
                
'Print total Stock Volume in column J row 2
            Range("J" & 2 + j).Value = Total
                    
            
            closingprice = Cells(i, 6).Value
            
            Range("K" & 2 + j).Value = closingprice - openingprice
           
            'Print yearly change
            
            If openingprice = 0 Then
           
                Range("L" & 2 + j).Value = 0
             
            Else
             
            'Print percent change
                Range("L" & 2 + j).Value = (closingprice / openingprice)
            
            End If
                    
            'Reset total to 0 for next tickers to work
            Total = 0
                        
            'Move to next row
            j = j + 1
                            
'Else keep adding to the total volume                    '(Cells(i,7).Value = volume column)
        Else
            Total = Total + Cells(i, 7).Value
                                    
        End If
                                    
    Next i
                                          
    
End Sub







'TRIED TO DO IT FOR EACH YEAR BUT DIDNT SEEM TO WORK

'Run on very worksheet
'Dim Ws As Worksheet

'For Each Ws In Worksheets

'Headings
    'Range("I1").Value = "Ticker"
    'Range("J1").Value = "Total Stock Volume"
    
'Each sheet reset
'Total = 0
'j = 0

'FOR LOOP - TICKER
        'Find the last non-blank cell in the the ticker column
        'rowcount = Cells(Rows.Count, "A").End(xlUp).Row
        
            'Loop will start
            'For i = 2 To rowcount
        
            'If the Ticker changes do the following
            'If Ws.Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                'Print New Ticker Letter in column I row 2
                'Ws.Range("I" & 2 + j).Value = Ws.Cells(i, 1).Value                                 '(Cells(i,1).Value = Ticker column)
                
                    'Print total Stock Volume in column J row 2
                    'Ws.Range("J" & 2 + j).Value = Total
                    
                        'Reset total to 0 for next tickers to work
                        'Total = 0
                        
                            'Move to next row
                            'j = j + 1
                            
                                    'Else keep adding to the total volume                    '(Cells(i,7).Value = volume column)
                                    'Else
                                    'Total = Total + Ws.Cells(i, 7).Value
                                    'End If
                                    
                                   ' Next i
                                    
                                    'Next Ws

