Attribute VB_Name = "Module1"
Sub stock_stats():

'Loop over worksheets
For Each ws In Worksheets


	ws.Range("O1").Value = "Ticker"
	ws.Range("P1").Value = "Value"
	ws.Range("N2").Value = "Greatest % Increase"
	ws.Range("N3").Value = "Greatest % Decrease"
	ws.Range("N4").Value = "Greatest Total Volume"


	ws.Range("I1").Value = "Ticker"
	ws.Range("J1").Value = "Yearly Change"
	ws.Range("K1").Value = "Percent Change"
	ws.Range("L1").Value = "Total Stock Volume"


    Dim ticker_name As String
    Dim total_vol As Double
    total_vol = 0
    Dim table_row As Integer
    table_row = 2
    Dim maxInc As Integer
    Dim maxDec As Integer
    Dim maxVol As Integer

    
    'define variables
    Dim yearOpen, yearClose, yearlyChange, percentChange As Double
    yearOpen = ws.Cells(2, 3).Value
              
    'find last row
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
    
        'check for last of ticker value
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker = ws.Cells(i, 1).Value
        total_vol = total_vol + ws.Cells(i, 7).Value

        ws.Range("I" & table_row).Value = ticker
        ws.Range("L" & table_row).Value = total_vol
		'calculate yearly change
        yearClose = ws.Cells(i, 6).Value
        yearlyChange = yearClose - yearOpen
        ws.Range("J" & table_row).Value = yearlyChange
                
        'check yearly change against 0 threshold, then fill cells accordingly
        If yearlyChange < 0 Then
			ws.Range("J" & table_row).Interior.ColorIndex = 3 'red
        Else
			ws.Range("J" & table_row).Interior.ColorIndex = 4 'green    
        End If
                
        'calculate percent change
        If yearOpen = 0 Then
			percentChange = yearClose - yearOpen
        Else
			percentChange = (yearlyChange / yearOpen)
        End If
        ws.Range("K" & table_row).Value = percentChange
                                
        'Move down a row & reset total volume
        table_row = table_row + 1
        total_vol = 0
            
        yearOpen = ws.Cells(i + 1, 3).Value
                             
        Else
            'add to running total
            total_vol = total_vol + ws.Cells(i, 7).Value
        
        End If
            
    Next i
    
    'report metrics
	
	'ticker with greatest increase
    ws.Range("P2").Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & table_row))
    ws.Range("P2").NumberFormat = "0.00%"
    maxInc = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & table_row)), ws.Range("K2:K" & table_row), 0)
    
    'ticker with greatest decrease
    ws.Range("P3").Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & table_row))
    ws.Range("P3").NumberFormat = "0.00%"
    maxDec = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & table_row)), ws.Range("K2:K" & table_row), 0)
    
    'ticker with greatest volume
    ws.Range("P4") = WorksheetFunction.Max(ws.Range("L2:L" & table_row))
    maxVol = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & table_row)), ws.Range("L2:L" & table_row), 0)
    
    'print associated tickers
    ws.Range("O2") = ws.Cells(maxInc + 1, 9).Value
    ws.Range("O3") = ws.Cells(maxDec + 1, 9).Value
    ws.Range("O4") = ws.Cells(maxVol + 1, 9).Value
            
Next ws
        
End Sub
