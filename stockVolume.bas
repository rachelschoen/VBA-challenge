Attribute VB_Name = "Module1"
Sub stockVolume()
      
      'First, we'll need to loop through all the worksheets.
      For Each ws In Worksheets
 
     'Because we need to reset each variable for each worksheet, we'll reset the initialization each loop.
     
     'We need to set an initial variable for holding the ticker name and total stock volume.
      Dim ticker As String
      
      
      'Set the initial variables for holding the open price and close price of a ticker.
      Dim openPrice As Double
      Dim closePrice As Double
      
      'Set the initial variables for holding values of the yearly change and the percent change of a ticker.
      Dim yearlyChange As Double
      Dim percentChange As Double
      
      'Set the initial variables for holding the greatest % increase, the greatest % decrease, the greatest total volume and their ticker names.
      Dim greatestIncrease As Double
      Dim greatestDecrease As Double
      Dim greatestTotal As Double
      Dim greatestIncreaseTickerName As String
      Dim greatestDecreaseTickerName As String
      Dim greatestTotalTickerName As String
      
      'Set the inital variable for keeping track of the location of different date's open price.
      Dim priceRow As Long
      priceRow = 2
      
      'Initialize the total stock volume to 0.
      Total = 0
      
      'Keep track of the location for the different names of stocks.
      Dim summaryTableRow As Integer
      summaryTableRow = 2
      
      'Set the header names.
      ws.Range("I1").Value = "Ticker"
      ws.Range("J1").Value = "Yearly Change"
      ws.Range("K1").Value = "Percent Change"
      ws.Range("L1").Value = "Total Stock Volume"
      ws.Range("P1").Value = "Ticker"
      ws.Range("Q1").Value = "Value"
      ws.Range("O2").Value = "Greatest % Increase"
      ws.Range("O3").Value = "Greatest % Decrease"
      ws.Range("O4").Value = "Greatest Total Volume"
      
      'Determine the last row of the worksheet.
      lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
      
      'Loop through column A to find different stock names and add stock volume together.
      For i = 2 To lastRow:
            
            '1) Determine if the cells are different.
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
               'Set the ticker name.
               ticker = ws.Cells(i, 1).Value
               
               'Add to the stock volume total.
               Total = Total + ws.Range("G" & i).Value
               
               'Print the ticker name.
               ws.Range("I" & summaryTableRow).Value = ticker
               
               'Print the total stock volume.
               ws.Range("L" & summaryTableRow).Value = Total
               
               'Calculate the yearly change and percent change.
               
               openPrice = ws.Range("C" & priceRow).Value
               closePrice = ws.Range("F" & i).Value
               yearlyChange = closePrice - openPrice
               
               'If the open price is 0, then we can say the percent change is 0 to avoid a divide by zero error.
               
                  If openPrice = 0 Then
                      percentChange = 0
                
                'Otherwise, we calculate the percent change by dividing the yearly change by the open price.
                
                  Else
                      percentChange = yearlyChange / openPrice
                  
                  End If
                 
                 'Print the values of the yearly change and percent change.
                 
                  ws.Range("J" & summaryTableRow).Value = yearlyChange
                  ws.Range("J" & summaryTableRow).NumberFormat = "0.00"
                  ws.Range("K" & summaryTableRow).Value = percentChange
                  ws.Range("K" & summaryTableRow).NumberFormat = "0.00%"
                  
                        'Add the required conditional formatting. Positive change should be green, while negative change should be red.
                        
                        'Check for positive change.
                        If ws.Range("J" & summaryTableRow).Value > 0 Then
                            ws.Range("J" & summaryTableRow).Interior.ColorIndex = 4
                        
                        'Check for negative change.
                        Else
                            ws.Range("J" & summaryTableRow).Interior.ColorIndex = 3
                        End If
                  
                  'Move to the next row in the summary table.
                  summaryTableRow = summaryTableRow + 1
                  priceRow = i + 1
               
                  'Reset the total stock volume.
                  Total = 0
            
            '2) If the cells are the same, then add the volume to the Total.
            
            Else
              Total = Total + ws.Range("G" & i).Value
                 
            End If
        
        'Move to the next worksheet row in column A.
        Next i
        
        'After finishing parsing all the rows,
        'Set the first ticker's percent change and total stock volume as the greatest ones.
        
        greatestIncrease = ws.Range("K2").Value
        greatestDecrease = ws.Range("K2").Value
        greatestTotal = ws.Range("L2").Value
        
        'Define the last row of the Ticker column.
        lastRowTicker = ws.Cells(Rows.Count, "I").End(xlUp).Row
        
        'Loop through each row of Ticker column to find the greatest results.
         For Row = 2 To lastRowTicker:
               If ws.Range("K" & Row + 1).Value > greatestIncrease Then
                  greatestIncrease = ws.Range("K" & Row + 1).Value
                  greatestIncreaseTickerName = ws.Range("I" & Row + 1).Value
               ElseIf ws.Range("K" & Row + 1).Value < greatestDecrease Then
                  greatestDecrease = ws.Range("K" & Row + 1).Value
                  greatestDecreaseTickerName = ws.Range("I" & Row + 1).Value
                ElseIf ws.Range("L" & Row + 1).Value > greatestTotal Then
                  greatestTotal = ws.Range("L" & Row + 1).Value
                  greatestTotalTickerName = ws.Range("I" & Row + 1).Value
                End If
            Next Row
            
            'Print the greatest % increase, the greatest % decrease, the greatest total volume, and their ticker names.
            ws.Range("P2").Value = greatestIncreaseTickerName
            ws.Range("P3").Value = greatestDecreaseTickerName
            ws.Range("P4").Value = greatestTotalTickerName
            ws.Range("Q2").Value = greatestIncrease
            ws.Range("Q3").Value = greatestDecrease
            ws.Range("Q4").Value = greatestTotal
            ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    'Move onto the next worksheet.
    Next ws

End Sub
