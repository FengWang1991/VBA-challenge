Sub Stock_Performance()

'variables
'i - variable for loops, v - total volumn of each stock, n - number of records on each worksheet, m - count the number of stocks on each worksheet
'op - open price of each stock, cp - closing price of each stock, gpi - greatest % increase of each worksheet, gpd - greatest % decrease of each worksheet, gtv - greatest total volume of each worksheet
'j - variable to help stock summary
'ticker(2) - array of string to save names of gpi, gpd, and gtv
Dim i, v, n, m As Long
Dim op, cp, gpi, gpd, gtv As Double
Dim j As Integer
Dim ticker(2) As String

'Get number of records on the current worksheet
n = ActiveSheet.UsedRange.Rows.Count
MsgBox ("There are " & n & " results!")

'Print the column names
Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percent Change"
Range("L1") = "Total Stock Volume"

'Assign values of the first stock to variables
j = 2
Cells(j, 9) = Cells(2, 1)
op = Cells(j, 3)
cp = Cells(j, 6)
v = 0
m = 1

'For-loop to compare each record to the first stock
For i = 2 To n
    
'If the stock name is the same as the name in the summary cell, update the closing price to the new closing price, and add the volume to the variable v.
'If it is a new stock name, get the price change, percent change (if open price is 0, percent change is unavailable), and total volume. Also, in the summary, move to the next row, print the name of stock, and assign new values to variables.
'The count variable m will be added 1.
    If Cells(i, 1) = Cells(j, 9) Then
        cp = Cells(i, 6)
        v = v + Cells(i, 7)
    Else
        Cells(j, 10) = cp - op
        If op = 0 Then
            Cells(j, 11) = "N/A"
        Else
            Cells(j, 11) = Format((cp / op - 1), "Percent")
        End If
        Cells(j, 12) = v
        m = m + 1
        j = j + 1
        Cells(j, 9) = Cells(i, 1)
        op = Cells(i, 3)
        cp = Cells(i, 6)
        v = Cells(i, 7)
    End If
Next i

'After the loop, the last stock values should be printed out.
Cells(j, 10) = cp - op
If op = 0 Then
    Cells(j, 11) = "N/A"
Else
    Cells(j, 11) = Format((cp / op - 1), "Percent")
End If
Cells(j, 12) = v

'Assign the first values of gpi, gpd, gv
gpi = Cells(2, 11)
gpd = Cells(2, 11)
gtv = Cells(2, 12)
ticker(0) = Cells(2, 9)
ticker(1) = Cells(2, 9)
ticker(2) = Cells(2, 9)

'Give green color if the change is positive, otherwise red color.
'Update the value of gpi, gpd, gtv, and ticker if necessary.
For i = 1 To m
    If Cells(i + 1, 10) > 0 Then
        Cells(i + 1, 10).Interior.Color = vbGreen
    ElseIf Cells(i + 1, 10) < 0 Then
        Cells(i + 1, 10).Interior.Color = vbRed
    End If

    If Cells(i + 1, 11) > gpi And Cells(i + 1, 11) <> "N/A" Then
        gpi = Cells(i + 1, 11)
        ticker(0) = Cells(i + 1, 9)
    End If

    If Cells(i + 1, 11) < gpd And Cells(i + 1, 11) <> "N/A" Then
        gpd = Cells(i + 1, 11)
        ticker(1) = Cells(i + 1, 9)
    End If

    If Cells(i + 1, 12) > gtv Then
        gtv = Cells(i + 1, 12)
        ticker(2) = Cells(i + 1, 9)
    End If

Next i

'Result of the bonus
Cells(2, 15) = "Greatest % Increase"
Cells(3, 15) = "Greatest % Decrease"
Cells(4, 15) = "Greatest Total Volume"
Cells(1, 16) = "Ticker"
Cells(1, 17) = "Value"
Cells(2, 16) = ticker(0)
Cells(2, 17) = Format(gpi, "Percent")
Cells(3, 16) = ticker(1)
Cells(3, 17) = Format(gpd, "Percent")
Cells(4, 16) = ticker(2)
Cells(4, 17) = gtv

End Sub

