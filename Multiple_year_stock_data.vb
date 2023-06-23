Sub Stock_data()

' Dims
    Dim Percent_Change As Double
    Dim Ticker_Name As String
    Dim Open_Price As Double
    Dim Closing_Price As Double
    
' Loop through all Worksheets
For Each ws In Worksheets
    
    ' Determine the Last Row
    Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Add Headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
     ' Set Open Stock Price
        Open_Price = ws.Cells(2, 3).Value

      ' Set Yearly Change
        Dim Yearly_Change As Double
        Yearly_Change = 0
        
      ' Set Total Stock Volume
        Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0
    
      ' Keep track of location for each Ticker Name in Summary Table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    
     ' Loop through all stock ticker values
        For i = 2 To Last_Row
    
             ' Check if we are within the same stock ticker
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                 ' Create Ticker Name
                    Ticker_Name = ws.Cells(i, 1).Value
                    
                 ' Show Ticker Name in Summary Table
                     ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

                  ' Calculate the Yearly Change
                    Yearly_Change = (ws.Cells(i, 6).Value - Open_Price)

                 ' Show Yearly Change in Summary Table
                     ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                     
                  ' Calculate Percent Change
                     Percent_Change = Yearly_Change / Open_Price
                     
                 ' Show Percent Change in Summary Table
                     ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                     ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                    
                 ' Calculate Total Stock Volume for each Ticker
                     Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                    
                 ' Print Total Stock Volume in Appropriate Column
                     ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                     
                 ' Color Formats
                        If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                        End If
        
                        If ws.Range("J" & Summary_Table_Row).Value < 0 Then
                            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                        End If
                        
                        If ws.Range("J" & Summary_Table_Row).Value = 0 Then
                            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 6
                        End If
                    
                        If ws.Range("K" & Summary_Table_Row).Value > 0 Then
                            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                        End If
        
                        If ws.Range("K" & Summary_Table_Row).Value < 0 Then
                            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                        End If
                        
                        If ws.Range("K" & Summary_Table_Row).Value = 0 Then
                            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 6
                        End If

                 ' Add one to the Summary Table Row
                     Summary_Table_Row = Summary_Table_Row + 1

                ' Reset Open Stock Price and Closing Stock Price
                     Open_Price = ws.Cells(i + 1, 3).Value
                     Total_Stock_Volume = 0

                 Else
                      
                  ' Add to the Total Stock Volume
                     Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

                 End If
        
            Next i
            
            ' Determine the Final Row
                Final_Row = ws.Cells(Rows.Count, 11).End(xlUp).Row
            
            ' Loop through summary table
                For i = 2 To Final_Row
            
                ' Add Headers
                    ws.Range("P1").Value = "Ticker"
                    ws.Range("Q1").Value = "Value"
                    ws.Range("O2").Value = "Greatest % Increase"
                    ws.Range("O3").Value = "Greatest % Decrease"
                    ws.Range("O4").Value = "Greatest Total Volume"
               
                 ' Find Max/Min Values
                    ws.Cells(2, 17).Value = WorksheetFunction.Max(ws.Range("K:K"))
                    ws.Cells(2, 17).NumberFormat = "0.00%"

                    ws.Cells(3, 17).Value = WorksheetFunction.Min(ws.Range("K:K"))
                    ws.Cells(3, 17).NumberFormat = "0.00%"
                    
                    ws.Cells(4, 17).Value = WorksheetFunction.Max(ws.Range("L:L"))
                    
                  ' Lookup Tickers for Max/Min Values
                    ws.Cells(2, 16).Value = WorksheetFunction.XLookup(ws.Cells(2, 17).Value, ws.Range("K:K"), ws.Range("I:I"))
                    ws.Cells(3, 16).Value = WorksheetFunction.XLookup(ws.Cells(3, 17).Value, ws.Range("K:K"), ws.Range("I:I"))
                    ws.Cells(4, 16).Value = WorksheetFunction.XLookup(ws.Cells(4, 17).Value, ws.Range("L:L"), ws.Range("I:I"))

                  Next i

            Next ws
            
End Sub

