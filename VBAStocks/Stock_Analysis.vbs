Attribute VB_Name = "Module2"
Sub VBAChallenge()

        For Each ws In Worksheets

            ' Create a new summary table
            ws.Cells(1, 10).Value = "Ticker"
            ws.Cells(1, 11).Value = "Yearly Change"
            ws.Cells(1, 12).Value = "Percent Change"
            ws.Cells(1, 13).Value = "Total Stock Volume"
            
            ' Challenge - Greatest... Summary Table
            ws.Cells(2, 16).Value = "Greatest % increase"
            ws.Cells(3, 16).Value = "Greatest % decrease"
            ws.Cells(4, 16).Value = "Greatest total volume"
            ws.Cells(1, 17).Value = "Ticker"
            ws.Cells(1, 18).Value = "Value"
            
            ' Challenege variables assignment to 0 before loop starts
            Dim great_inc, great_dec, great_vol As Double
            great_inc = 0
            great_dec = 0
            great_vol = 0
            
            ' Determine the Last Row
            Dim LastRow As Long
            LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
            
            ' Set an initial variable for hold the total of each ticker
            Dim stocktotal As Double
            stocktotal = 0
          
            ' Keep track of the location for each credit card brand in the summary table
            Dim summary As Integer
            summary = 2
            
            Dim firsttickeropen As Double
            firsttickeropen = ws.Range("C2").Value
            
            ' Create a loop to go down ticker column
            For i = 2 To LastRow
               
                ' Define my variables for ticker within the loop
                Dim currentticker As String
                currentticker = ws.Cells(i, 1).Value
                        
                Dim nextticker As String
                nextticker = ws.Cells(i + 1, 1).Value
        
                ' Look ahead if the ticker is still the same
                If currentticker = nextticker Then
        
                    ' Add to the stock sum volume
                    stocktotal = stocktotal + ws.Cells(i, 7).Value
                    
                ' If the cell following the next row has different ticker
                Else
                               
                    ' INSERT THE SUM VOLUME ----------------------------------------
                    ' Add the last value of the same ticker before going to a new ticker
                    stocktotal = stocktotal + ws.Cells(i, 7).Value
                    
                    ' Print the sum of all stock volume to the new summary table
                    ws.Range("M" & summary).Value = stocktotal
                                                
                    ' INSERT THE TICKER --------------------------------------------
                    ' Print the ticker name in the new summary table
                    ws.Range("J" & summary).Value = currentticker
                    
                    ' INSERT THE YEARLY CHANGE----------------------------------------
                    ' Pull out the close stock price of the stock within ticker
                    Dim lasttickerclose As Double
                    lasttickerclose = ws.Cells(i, 6).Value
                    
                    ' Caluculate the yearly change
                    Dim yearlychange As Double
                    yearlychange = lasttickerclose - firsttickeropen
                    
                    ' Calculate percentage change
                    Dim percentchange As Double
                    
                    ' Allow the loop to keeps going even if it comes across opeck stock with 0 value
                    If firsttickeropen = 0 Then
                        
                        ' Get the loop going even with 0 is found, keep the math going!
                        percentchange = yearlychange
                        
                    Else
                        ' Otherwise, do its thing to find the % change
                        percentchange = yearlychange / firsttickeropen
                    
                    End If
                    
                    ' Pull out the open stock price of the ticker in the next row
                    firsttickeropen = ws.Cells(i + 1, 3).Value
                    
                        If yearlychange > 0 Then
                        
                            ' Set color variables
                            Dim ColorRed As Integer
                            ColorRed = 3
                            Dim ColorGreen As Integer
                            ColorGreen = 4
                
                            ' highlight positive change in green
                            ws.Range("K" & summary).Interior.ColorIndex = ColorGreen
                        
                        ElseIf yearlychange < 0 Then
                            
                            ' highlight negative change in red.
                            ws.Range("K" & summary).Interior.ColorIndex = ColorRed
                    
                        End If
                        
                    ' INSERT THE % CHANGE----------------------------------------
                    ' Print the yearly change  in the new summary table
                    ws.Range("K" & summary).Value = yearlychange
                    
                    ' Print the % change  in the new summary table
                    ws.Range("L" & summary).Value = Format(percentchange, "Percent")
                    
                    ' Add one to summary table row to go to the next row
                    summary = summary + 1
                    
                    ' Challenege - ticker stock total comparison before the summary table is being reset
                    
                    ' As the loop goes down, find the greater value and have that be my next great_inc value
                    Dim inc_ticker, dec_ticker, vol_ticker As String
                                        
                    If great_inc < percentchange Then
                    
                        great_inc = percentchange
                        inc_ticker = currentticker
                        
                    End If
                        
                    ' As the loop goes down, find the lower value and have that be my next great_dec value
                    If great_dec > percentchange Then
                    
                        great_dec = percentchange
                        dec_ticker = currentticker
                        
                    End If
                    
                    ' As the loop goes down, find the greater value and have that be my next great_vol value
                    If great_vol < stocktotal Then
                    
                        great_vol = stocktotal
                        vol_ticker = currentticker
                        
                    End If
                    
                    ' Reset the summary table
                    stocktotal = 0
                    
                End If
            
            Next i
        
            ' Challenge - Print all the max values after the loop is complete
            ws.Cells(2, 18).Value = Format(great_inc, "Percent")
            ws.Cells(3, 18).Value = Format(great_dec, "Percent")
            ws.Cells(4, 18).Value = Format(great_vol, "Scientific")
            ws.Cells(2, 17).Value = inc_ticker
            ws.Cells(3, 17).Value = dec_ticker
            ws.Cells(4, 17).Value = vol_ticker
    
            ws.Columns("P:R").AutoFit
         
        Next ws

End Sub
        

