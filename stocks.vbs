Sub stocks()

'   Set variables
Dim ticker As String
Dim open_price As Double
Dim close_price As Double
Dim delta_price As Double
Dim delta_pct As Double
Dim volume As LongLong

'   Run this code through every worksheet in the workbook
For Each ws In Worksheets

    '   Set table headers
    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "Yrly Change"
    ws.Range("L1").Value = "Pct Change"
    ws.Range("M1").Value = "Total Volume"

    '   Set initial value of variables to zero for calculating totals by stock ticker symbol
    open_price = ws.Cells(2, 3)
    volume = 0

    '   Keep track of each stock ticker symbol
    Dim summary As Integer
    summary = 2

    '   Find last row
    Dim last_row As Long
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

        '   Loop through all stock ticker symbols
        For i = 2 To last_row
    
            '   Check if working on the same stock ticker symbol
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                '   Set values
                ticker = ws.Cells(i, 1).Value
                close_price = ws.Cells(i, 6).Value
    
                '   Calculate change
                delta_price = close_price - open_price
                
                '   Avoid divide by zero error
                If open_price <> 0 Then
                    delta_pct = (delta_price / open_price) * 100
                Else
                    delta_pct = 0
                End If
                
                '   Add to total volume of each stock ticker symbol
                volume = volume + ws.Cells(i, 7).Value
                
                '   Print values in the Summary Table
                ws.Range("J" & summary).Value = ticker
                ws.Range("K" & summary).Value = delta_price
                
                    '   Set conditional formatting to highlight positive change in green and negative change in red
                    If delta_price < 0 Then
                        ws.Range("K" & summary).Interior.ColorIndex = 3
                    Else
                        ws.Range("K" & summary).Interior.ColorIndex = 4
                    End If
                
                ws.Range("L" & summary).Value = Round(delta_pct, 2) & "%"
                
                ws.Range("M" & summary).Value = volume
                
                '   Loop through each stock ticker symbol
                summary = summary + 1
                open_price = ws.Cells(i + 1, 3).Value
                volume = 0
            Else
                volume = volume + ws.Cells(i, 7).Value
            End If
        Next i
Next ws
End Sub
