Attribute VB_Name = "Módulo2"
Sub Greatest()

'Variables declaration
Dim index_one As Integer
Dim index_zero As Integer
Dim last_row As Long
Dim lowest As Double
Dim highest As Double
Dim max_stock As Double
Dim ti_hi As String
Dim ti_low As String
Dim ti_stock As String
Dim index As Double

'Set the header
Range("P1:Q1") = Array("Ticker", "Value")

'Set the left column values
left_values = Array("Greatest % increase", "Greatest % decrease", "Greatest Total Volume")

index_one = 0
For index_zero = 2 To 4
    Cells(index_zero, 15) = left_values(index_one)
    index_one = index_one + 1
Next

'To get the last row for the following for loop
last_row = Cells(Rows.Count, 9).End(xlUp).Row

'Iterate over the tickers' yearly values to get highest, lowest and max_stock
highest = Cells(2, 11)
lowest = Cells(2, 11)
max_stock = Cells(2, 12)

'One if statement for each variable to evaluate versus the current value assignated
For index = 2 To last_row
    
    If highest < Cells(index, 11) Then
        highest = Cells(index, 11)
        ti_hi = Cells(index, 9)
    End If

    If lowest > Cells(index, 11) Then
        lowest = Cells(index, 11)
        ti_low = Cells(index, 9)
    End If
    
    If max_stock < Cells(index, 12) Then
        max_stock = Cells(index, 12)
        ti_stock = Cells(index, 9)
    End If

Next

'Return final values for highest, lowest and max_stock

Range("P2:Q2") = Array(ti_hi, FormatPercent(highest))
Range("P3:Q3") = Array(ti_low, FormatPercent(lowest))
Range("P4:Q4") = Array(ti_stock, max_stock)

End Sub
