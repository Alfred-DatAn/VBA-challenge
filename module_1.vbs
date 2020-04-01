Sub yearly_overview()


'Variables declaration
Dim index As Long
Dim index_two As Long
Dim ticker As String
Dim total_stock As Double
Dim open_price As Double
Dim close_price As Double
Dim index_three As Integer

'Set the header for the yearly overview
Range("I1:L1") = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")

'Set the values for the loop
index = 2
index_two = 2
ticker = Cells(2, 1)
total_stock = 0
open_price = Cells(2, 3)
close_price = 0

'This while loop works as long as the current cell has any value
Do While IsEmpty(Cells(index, 1)) = False
'If the current cell value has the same value as the "ticker" variable...
    If Cells(index, 1) = ticker Then
        total_stock = total_stock + Cells(index, 7)
'This entire second if statement works only for the very last value of the table.
'It prints all the variables before the while loop finishes
        If IsEmpty(Cells(index + 1, 1)) = True Then
            Cells(index_two, 9) = ticker
            close_price = Cells(index, 6)
            Cells(index_two, 10) = close_price - open_price
'This block prevents doing division by zero when getting the stock percent change
            If open_price = 0 Then
            open_price = 1
            End If
            Cells(index_two, 11) = FormatPercent(Cells(index_two, 10) / open_price, [2])
            Cells(index_two, 12) = total_stock
        End If
'This else block contains the exact same code as the previous if statement
'It runs every time the "ticker" value changes
    Else
        Cells(index_two, 9) = ticker
        close_price = Cells(index - 1, 6)
        Cells(index_two, 10) = close_price - open_price
        If open_price = 0 Then
            open_price = open_price + 1
        End If
        Cells(index_two, 11) = FormatPercent(Cells(index_two, 10) / open_price, [2])
        Cells(index_two, 12) = total_stock
        
        ticker = Cells(index, 1)
        open_price = Cells(index, 3)
        total_stock = Cells(index, 7)
        index_two = index_two + 1
            
    End If
    index = index + 1
   Loop

'After the first while loop is finished, a new while loop runs to evaluate every
'value of the yearly change column and apply the corresponding format
index_three = 2

Do While IsEmpty(Cells(index_three, 10)) = False
    If Cells(index_three, 10) >= 0 Then
        Cells(index_three, 10).Interior.Color = RGB(170, 198, 27)
    Else
        Cells(index_three, 10).Interior.Color = RGB(207, 45, 59)
    End If
    index_three = index_three + 1
Loop


End Sub