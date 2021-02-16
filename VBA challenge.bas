Attribute VB_Name = "Module1"
Sub testdata()

'Loop through all worksheets
For Each ws In ThisWorkbook.Worksheets
ws.Activate

Dim ticker As String
Dim year_change As Double
Dim percent_change As Double
Dim total_volume As Double
Dim summary_table As Double
Dim year_open As Double
Dim year_close As Double
Dim greatest_increase As Double
Dim greatest_decrease As Double
Dim greatest_total_volume As Double
Dim value As Double


ws.Cells(1, 9).value = "Ticker"
ws.Cells(1, 10).value = "Year Change"
ws.Cells(1, 11).value = "Percent Change"
ws.Cells(1, 12).value = "Total Stock Volume"
ws.Cells(2, 14).value = "Greatest % Increase"
ws.Cells(3, 14).value = "Greatest % Decreased"
ws.Cells(4, 14).value = "Greatest Total Volume"
ws.Cells(1, 15).value = "Ticker Code"
ws.Cells(1, 16).value = "Value"

total_volume = 0
summary_table = 2

'set year open value for every worksheet
year_open = ws.Cells(2, 3).value

'Determine the lastrow
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastRow

    If ws.Cells(i - 1, 1).value = ws.Cells(i, 1).value And ws.Cells(i, 1).value <> ws.Cells(i + 1, 1).value Then

    'set the close price
    year_close = ws.Cells(i, 6).value

    'calculate year change
    year_change = year_close - year_open

    'display year_change
    ws.Cells(summary_table, 10).value = year_change
    
        If ws.Cells(summary_table, 10).value > 0 Then
        
        'set color for year change
        ws.Cells(summary_table, 10).Interior.Color = vbGreen
    
        Else
        ws.Cells(summary_table, 10).Interior.Color = vbRed
    
        End If
        
        If year_open = 0 And year_close <> 0 Then
        percent_change = 1
        
        ElseIf year_open = 0 And year_close = 0 Then
        percent_change = 0
        
        Else
        percent_change = (year_close - year_open) / year_open
    
        'display percent change
        ws.Cells(summary_table, 11).value = percent_change
        
        'percent format
        ws.Cells(summary_table, 11).NumberFormat = "%0.00"
        
            
        End If
        
        'get open price for next year
        year_open = ws.Cells(i + 1, 3).value
    
    End If

    If ws.Cells(i + 1, 1).value <> ws.Cells(i, 1).value Then

    'set the ticker name
    ws.Cells(summary_table, 9).value = ws.Cells(i, 1).value

    'set the total volume
    ws.Cells(summary_table, 12).value = total_volume + ws.Cells(i, 7).value

    'add 1 to summary_table
    summary_table = summary_table + 1

    'reset total_volume to 0
    total_volume = 0

    Else

    'add total volume if the tickers are the same
    total_volume = total_volume + ws.Cells(i, 7).value

    End If

   Next i
    Dim strng As String
    Dim strng1 As String
    Dim strng2 As String

    'find ticker code of the greatest percent increase or MAX value
    strng = ws.Evaluate("INDEX(I:I,MATCH(MAX(K:K),K:K,0))")
    
    'display ticker code
    ws.Cells(2, 15).value = strng
    
    'find ticker code of the greatest percent decrease or MIN value
    strng1 = ws.Evaluate("INDEX(I:I,MATCH(MIN(K:K),K:K,0))")
    
    'display ticker code
    ws.Cells(3, 15).value = strng1
    
    'find ticker code of the greatest total stock volume
    strng2 = ws.Evaluate("INDEX(I:I,MATCH(MAX(L:L),L:L,0))")
    
    'display ticker code
    ws.Cells(4, 15).value = strng2
    
    
    ws.Cells(2, 16).value = WorksheetFunction.Max(Range("K2" & ":" & "K" & lastRow))
    ws.Cells(2, 16).NumberFormat = "%0.00"
    
    ws.Cells(3, 16).value = WorksheetFunction.Min(Range("K2" & ":" & "K" & lastRow))
    ws.Cells(3, 16).NumberFormat = "%0.00"
    
    ws.Cells(4, 16).value = WorksheetFunction.Max(Range("L2" & ":" & "L" & lastRow))
    
  Next ws

End Sub

