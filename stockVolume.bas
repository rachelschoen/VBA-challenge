Attribute VB_Name = "Module1"
Sub stockVolume()
      
      'First, we'll need to loop through all the worksheets.
      For Each Worksheet In Worksheets
 
     'Because we need to reset each variable for each worksheet, we'll reset the initialization for each year worksheet.
     
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
      
      'We'll initialize the total stock volume to 0.
      Dim Total As Double
      Total = 0
      
      'Keep track of the location for the different names of stocks.
      Dim summaryTableRow As Integer
      summaryTableRow = 2
      
      'Set the header names.
      Worksheet.Range("I1").Value = "Ticker"
      Worksheet.Range("J1").Value = "Yearly Change"
      Worksheet.Range("K1").Value = "Percent Change"
      Worksheet.Range("L1").Value = "Total Stock Volume"
      Worksheet.Range("P1").Value = "Ticker"
      Worksheet.Range("Q1").Value = "Value"
      Worksheet.Range("O2").Value = "Greatest % Increase"
      Worksheet.Range("O3").Value = "Greatest % Decrease"
      Worksheet.Range("O4").Value = "Greatest Total Volume"
      
      'Determine the last row of the worksheet.
      lastRow = Worksheet.Cells(Rows.Count, 1).End(xlUp).Row
      
      'Loop through column A to find different stock names and add stock volume together.
      For i = 2 To lastRow:
            
            '1) Determine if the cells are different.
            If Worksheet.Cells(i + 1, 1).Value <> Worksheet.Cells(i, 1).Value Then
            
               'Set the ticker name.
               ticker = Worksheet.Cells(i, 1).Value
               
               'Add to the stock volume total.
               Total = Total + Worksheet.Range("G" & i).Value
               
               'Print the ticker name.
               Worksheet.Range("I" & summaryTableRow).Value = ticker
               
               'Print the total stock volume.
               Worksheet.Range("L" & summaryTableRow).Value = Total
               
               'Calculate the yearly change and percent change.
               
               openPrice = Worksheet.Range("C" & priceRow).Value
               closePrice = Worksheet.Range("F" & i).Value
               yearlyChange = closePrice - openPrice
               
               'If the open price is 0, then we can say the percent change is 0 to avoid a divide by zero error.
               
                  If openPrice = 0 Then
                      percentChange = 0
                
                'Otherwise, we calculate the percent change by dividing the yearly change by the open price.
                
                  Else
                      percentChange = yearlyChange / openPrice
                  
                  End If
                 
                 'Print the values of the yearly change and percent change.
                 
                  Worksheet.Range("J" & summaryTableRow).Value = yearlyChange
                  Worksheet.Range("J" & summaryTableRow).NumberFormat = "0.00"
                  Worksheet.Range("K" & summaryTableRow).Value = percentChange
                  Worksheet.Range("K" & summaryTableRow).NumberFormat = "0.00%"
                  
                        'Add the required conditional formatting. Positive change should be green, while negative change should be red.
                        
                        'I'll format both yearly change and percent change at the same time, as they have the same relationship.
                        
                        'Check for positive change.
                        If Worksheet.Range("J" & summaryTableRow).Value > 0 Then
                            Worksheet.Range("J" & summaryTableRow).Interior.ColorIndex = 4
                            Worksheet.Range("K" & summaryTableRow).Interior.ColorIndex = 4
                        
                        'Check for negative change.
                        Else
                            Worksheet.Range("J" & summaryTableRow).Interior.ColorIndex = 3
                            Worksheet.Range("K" & summaryTableRow).Interior.ColorIndex = 3
                        End If
                  
                  'Move to the next row in the summary table.
                  summaryTableRow = summaryTableRow + 1
                  priceRow = i + 1
               
                  'Reset the total stock volume.
                  Total = 0
            
            '2) If the cells are the same, add the volume to the Total.
            
            Else
              Total = Total + Worksheet.Range("G" & i).Value
                 
            End If
        
        'Move to the next worksheet row in column A.
        Next i
        
        'After finishing parsing all the Rows,
        'Set the first ticker's percent change and total stock volume as the greatest ones.
        
        greatestIncrease = Worksheet.Range("K2").Value
        greatestDecrease = Worksheet.Range("K2").Value
        greatestTotal = Worksheet.Range("L2").Value
        
        'Define the last row of the Ticker column.
        lastRowTicker = Worksheet.Cells(Rows.Count, "I").End(xlUp).Row
        
        'Loop through each row of Ticker column to find the greatest results.
         For Row = 2 To lastRowTicker:
               If Worksheet.Range("K" & Row + 1).Value > greatestIncrease Then
                  greatestIncrease = Worksheet.Range("K" & Row + 1).Value
                  greatestIncreaseTickerName = Worksheet.Range("I" & Row + 1).Value
               ElseIf Worksheet.Range("K" & Row + 1).Value < greatestDecrease Then
                  greatestDecrease = Worksheet.Range("K" & Row + 1).Value
                  greatestDecreaseTickerName = Worksheet.Range("I" & Row + 1).Value
                ElseIf Worksheet.Range("L" & Row + 1).Value > greatestTotal Then
                  greatestTotal = Worksheet.Range("L" & Row + 1).Value
                  greatestTotalTickerName = Worksheet.Range("I" & Row + 1).Value
                End If
            Next Row
            
            'Print the greatest % increase, the greatest % decrease, the greatest total volume, and their ticker names.
            Worksheet.Range("P2").Value = greatestIncreaseTickerName
            Worksheet.Range("P3").Value = greatestDecreaseTickerName
            Worksheet.Range("P4").Value = greatestTotalTickerName
            Worksheet.Range("Q2").Value = greatestIncrease
            Worksheet.Range("Q3").Value = greatestDecrease
            Worksheet.Range("Q4").Value = greatestTotal
            Worksheet.Range("Q2:Q3").NumberFormat = "0.00%"
    
    'Move onto the next worksheet.
    Next Worksheet

End Sub
