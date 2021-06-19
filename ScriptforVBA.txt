Sub StocksSheetsLoop()

'Set initial variable for stock
Dim Ticker As String

    For Each ws In Worksheets
        
        'Describe the data contents of each colum
        ws.Cells(1, "I").Value = "Ticker"
        ws.Cells(1, "J").Value = "Yearly Change"
        ws.Cells(1, "K").Value = "Percent Change"
        ws.Cells(1, "L").Value = "Total Stock Volume"
        
        'Format column widths to fit contents
        ws.Columns("A:L").AutoFit
        
        'Set variables for our data
        Dim LastRow As Long
        Dim Volume As Double
        Volume = 0
        Dim Ticker_Opening_Price As Double
        Dim Ticker_Closing_Price As Double
        Dim Yearly_Change As Double
        Dim Previous_Amount As Long
        Previous_Amount = 2
        Dim Percentage_Change As Double
        
        'Keep track of the location for each stock in the summary table
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
        'Determine Last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Loop through all stocks
        For i = 2 To LastRow
            
            Volume = Volume + ws.Cells(i, 7).Value
                
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                Ticker = ws.Cells(i, 1).Value
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                
                'Put Volume total in the summary table
                ws.Range("L" & Summary_Table_Row).Value = Volume
                
                Volume = 0
                
                Ticker_Opening_Price = ws.Range("C" & Previous_Amount)
                Ticker_Closing_Price = ws.Range("F" & i)
                Yearly_Change = Ticker_Closing_Price - Ticker_Opening_Price
                
                'Yearly_Change in the summary table
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                
                'Format to show green and red
                If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                    
                    'Calculate the change percentage
                    If Ticker_Opening_Price = 0 Then
                        Percentage_Change = 0
                    Else
                        Ticker_Opening_Price = ws.Range("C" & Previous_Amount)
                        Percentage_Change = Yearly_Change / Ticker_Opening_Price
                    End If
                    
                    ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                    ws.Range("k" & Summary_Table_Row) = Percentage_Change
                    
                    
                    
                'Summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                Previous_Amount = i + 1
                End If
                
            Next i
    
    Next ws
            

End Sub