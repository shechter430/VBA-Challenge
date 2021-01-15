Attribute VB_Name = "Module1"
Sub stock_output()

'Set loop to activate through all worksheets
For Each ws In Worksheets
ws.Activate

'Set variables
Dim ticker As String
Dim open_price As Double
    'open_price = 0
Dim close_price As Double
    close_price = 0
Dim yearly_change As Double
    yearly_change = 0
Dim percent_change As Double
    percent_change = 0
Dim total_vol As Double
    total_vol = 0
Dim summary_tablerow As Double
    summary_tablerow = 2
Dim column As Integer
    column = 1
Dim i As Long
    
'Set column headers
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
    
ws.Columns("I:L").AutoFit
    
'Find last row
Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Set initial open price
open_price = Cells(2, column + 2).Value


'Loop through all stocks
For i = 2 To lastrow

    'Check if we are still within the same ticker symbol
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        'Set ticker symbol
        ticker = Cells(i, 1).Value
        'Calculate yearly change in stock prices
        close_price = Cells(i, 6).Value
        yearly_change = close_price - open_price
        'Calculate percent change in stock prices
            If open_price = 0 And close_price = 0 Then
               percent_change = 0
               ElseIf open_price = 0 And close_price <> 0 Then
               percent_change = 1
               Else
               percent_change = yearly_change / open_price
            End If
        'Add ticker symbol to total volume
        total_vol = total_vol + Cells(i, 7).Value
        
        'Print values in summary table
        Range("I" & summary_tablerow).Value = ticker
        Range("J" & summary_tablerow).Value = yearly_change
        Range("K" & summary_tablerow).Value = percent_change
            Range("K" & summary_tablerow).NumberFormat = "0.00%"
        Range("L" & summary_tablerow).Value = total_vol
        
        'Conditional formatting for yearly change values(green = positive , red = negative)
            If Range("J" & summary_tablerow).Value >= 0 Then
                Range("J" & summary_tablerow).Interior.ColorIndex = 4
            Else
                Range("J" & summary_tablerow).Interior.ColorIndex = 3
            End If
            
        'Adding one to the summary table row
        summary_tablerow = summary_tablerow + 1
        
        'Reset total volume
        total_vol = 0
    
        'Reset open price
        open_price = Cells(i + 1, column + 2).Value
        
    Else
        'Add to total volume
        total_vol = total_vol + Cells(i, 7).Value
        
    End If
    
Next i
Next ws


End Sub



    
