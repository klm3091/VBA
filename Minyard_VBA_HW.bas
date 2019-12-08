Attribute VB_Name = "Module1"
Sub StockData()

'Loop throught the Worksheets
For Each ws In ThisWorkbook.Worksheets

    'Set initial variable to track ticker
    Dim Ticker As String
  
    'Set an initial variable to hold the volume total
    Dim Volume As Double
    Volume = 0
  
    'Set year open and year close as variables
    Dim year_open As Double
    Dim year_close As Double
  
    'Set variable for Year Change and Percent Change
    Dim year_change As Double
    Dim percent_change As Double
  
    'Keep track of the ticker
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
  
  'Headers
  ws.Cells(1, 9).Value = "Ticker"
  ws.Cells(1, 10).Value = "Year Change"
  ws.Cells(1, 11).Value = "Percent Change"
  ws.Cells(1, 12).Value = "Total Volume"
  
'Set last row
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop through all tickers
     For i = 2 To LastRow
     
     
        'Check to see if we are still within the same ticker
         If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
         
         'Set ticker
         Ticker = ws.Cells(i, 1).Value
         
         'Add to year change
         year_open = ws.Cells(i, 3).Value
         
         'Add to volume
         Volume = Volume + ws.Cells(i, 7).Value
         
         End If
         
         'If it is the last row of that ticker
         If ws.Cells((i + 1), 1).Value <> ws.Cells(i, 1).Value Then
         
         'Pull Year Close value
         year_close = ws.Cells((i + 1), 6).Value
         
         'Calculate Year Change value
         year_change = year_close - year_open
         
         ' Color cell based on Year Change
                For j = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
                    If ws.Cells(j, 10).Value >= 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 4
                    
                    ElseIf ws.Cells(j, 10).Value < 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                    End If
                    
                    Next j

       
         'Calculate Percent Change
         If year_open = 0 Then
            percent_change = year_change
            
        Else
            percent_change = year_change / year_open
            
        End If

     
     'Print the ticker summary in Row 9
     ws.Range("I" & Summary_Table_Row).Value = Ticker
     
     'Print the Year Change in Row 10
     ws.Range("J" & Summary_Table_Row).Value = year_change
     
     'Print the Percent Change in Row 11
     ws.Range("K" & Summary_Table_Row).Value = percent_change
     
     'Format to two decimal places
     ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
     
    'Print the volume amount in Row 12
    ws.Range("L" & Summary_Table_Row).Value = Volume
    
    'Add one to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1

    
    'Reset the ticker total
    Ticker = 0
    
    'If the next cell is the same ticker
    Else
    
    'Add to the volume total
    Volume = Volume + ws.Cells(i, 7).Value
    
    End If
    
    Next i
    
    Next ws
     
  End Sub
