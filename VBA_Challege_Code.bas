Attribute VB_Name = "Module1"
Sub Stocks()
Dim firstX As Double, lastX As Double, rowList As Double, i As Double, Brand_Total As Double, divisor As Double


    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

firstX = 2

rowList = 2
'rowList is a list of distinct Ticker of Stocks

LastRow = Cells(Rows.Count, 1).End(xlUp).Row
'Cells(1, 13).Value = lastrow
    Brand_Total = 0
    For i = 2 To LastRow
    If Cells(i, 1).Value = Cells(firstX, 1).Value Then
    
    Brand_Total = Brand_Total + Cells(i, 7).Value
   
        ElseIf (Cells(i, 1).Value <> Cells(firstX, 1).Value) Then
            lastX = i - 1
            divisor = Cells(firstX, 3).Value
                Cells(rowList, 9).Value = Cells(lastX, 1).Value
                Cells(rowList, 10).Value = (Cells(lastX, 6).Value - (Cells(firstX, 3).Value))
                        If divisor = 0 Then
                        Cells(rowList, 11).Value = "undefined"
                         Else: Cells(rowList, 11).Value = (Cells(lastX, 6).Value - (Cells(firstX, 3).Value)) / divisor
                        Cells(rowList, 11).NumberFormat = "0.00%"
                        End If
                    If Cells(rowList, 11).Value > 0 Then
                         Cells(rowList, 11).Interior.ColorIndex = 4
                    Else: Cells(rowList, 11).Interior.ColorIndex = 3
                    End If
                    
                Cells(rowList, 12).Value = Brand_Total
                Brand_Total = 0
                
firstX = i
rowList = rowList + 1

        End If
    Next i


End Sub

