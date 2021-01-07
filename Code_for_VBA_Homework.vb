Option Explicit
Sub vba_homework()
' Double: 1.2
' Integer: 1
' Long: is an integer that's long

Dim ticker As String
Dim price_change As Double
Dim percent_change As Double
Dim stock_vol As Double
Dim last_row As Double
Dim start As Long
Dim difference As Double
Dim i As Long
Dim j As Integer
Dim ws As Worksheet
Dim open_price As Double
Dim close_price As Double
Dim volume As Double
Dim startprice As Double

'to help it not crash
Application.ScreenUpdating = False

For Each ws In Worksheets:

    'labels
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    stock_vol = 0
    'to keep track of where to start open prices
    start = 2
    'to keep track of new chart created with solutions
    j = 2
    
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'looping
    For i = 2 To last_row
    
        'to check if sticker is still the same
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            ' stock_vol calculation
            stock_vol = stock_vol + ws.Cells(i, 7).Value
            
            ' end of ticker calculations
            ticker = ws.Cells(i, 1).Value
            ws.Cells(j, 9).Value = ticker
            ws.Cells(j, 12).Value = stock_vol
            
            open_price = ws.Cells(start, 3).Value
            close_price = ws.Cells(i, 6).Value
            
            difference = close_price - open_price
            
            'color in yearly price change column
            If difference < 0 Then
                ws.Range("J" & j).Interior.ColorIndex = 3
            Else
                ws.Range("J" & j).Interior.ColorIndex = 4
            End If
            
            ws.Range("J" & j).Value = difference
            
            'to avoid dividing by zero
            If open_price = 0 Then
                open_price = 1
            End If
            
            ws.Range("K" & j).Value = difference / open_price
            
            ' adjust variables for next ticker
            start = i + 1
            stock_vol = 0
            j = j + 1
            
        Else
            ' stock_vol calculation
            stock_vol = stock_vol + ws.Cells(i, 7).Value
            
        End If
    Next i
Next ws
'to help it not crash
Application.ScreenUpdating = True
End Sub