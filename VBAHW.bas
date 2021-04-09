Attribute VB_Name = "StockAnalyze"
Sub StockAnalyze()
'Loops through all worksheets
For Each ws In Worksheets
ws.Activate

Cells(1, 8).Value = "Ticker Symbol"
Cells(1, 9).Value = "Total Stock"
Cells(1, 10).Value = "Price Change"
Cells(1, 11).Value = "Percentage Change"

'cast variables

Dim ticker As String
Dim summarytable As Long
Dim lastrow As Long
Dim totalstock As Double
Dim openprice As Double
Dim closeprice As Double
Dim percentchange As Double



totalstock = 0

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

summarytable = 2


'Loops, engage

For I = 2 To lastrow


If Cells((I - 1), 1).Value <> Cells(I, 1).Value Then

openprice = Cells(I, 3)
End If

'If statement to separate stocks

If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

'Retrieve ticker

ticker = Cells(I, 1).Value

closeprice = Cells(I, 6)

Range("J" & summarytable).Value = closeprice - openprice

Range("J" & summarytable).NumberFormat = "$0.00"

If openprice = 0 Then

percentchange = 0

Else
'Retrieve percentage change
percentchange = (-1 * ((openprice - closeprice) / ((openprice + closeprice) / 2)))

Range("K" & summarytable).Value = percentchange

End If
'Format Percentage Change
Range("K" & summarytable).NumberFormat = "0.00%"

Range("H" & summarytable).Value = ticker

'retrieve total stock volume

totalstock = totalstock + Cells(I, 7).Value

Range("I" & summarytable).Value = totalstock

summarytable = summarytable + 1

totalstock = 0

'Else statement for tallying (adds up when tickers are not unequal to next row)

Else

totalstock = totalstock + Cells(I, 7).Value

Range("I" & summarytable).Value = totalstock

ticker = Cells(I, 1).Value

Range("H" & summarytable).Value = ticker

End If

Next I

'conditional formatting through loop and If statements. Feels a bit brute-force

For J = 2 To lastrow

If Cells(J, 10).Value > 0 Then
Cells(J, 10).Interior.Color = vbGreen

End If

If Cells(J, 10).Value < 0 Then
Cells(J, 10).Interior.Color = vbRed
End If

If Cells(J, 11).Value > 0 Then
Cells(J, 11).Interior.Color = vbGreen
End If

If Cells(J, 11).Value < 0 Then
Cells(J, 11).Interior.Color = vbRed
End If

Next J

Next

End Sub
