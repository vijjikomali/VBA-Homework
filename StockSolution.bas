Attribute VB_Name = "StockSolution"
Sub stock_analysis()

Dim ttl As Double, i As Long, change As Single, j As Integer, strt As Long, rowNo As Integer
Dim perChange As Single, days As Integer, dailyChange As Single, avgChange As Single

'Dim ws1 As Worksheet
Dim ws2 As Worksheet

' results can be changed to whatever worksheet is needed
Set ws2 = Worksheets("Results")

j = 0
ttl = 0
change = 0
strt = 2
dailyChange = 0

' get the row number of the last row with data
rowcount = Cells(Rows.Count, "A").End(xlUp).Row


' 701937 and 78
For i = 2 To rowcount
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        ' Stores results in variables
        ttl = ttl + Cells(i, 7).Value
        change = (Cells(i, 6) - Cells(strt, 3))
        perChange = Round((change / Cells(strt, 3) * 100), 2)
        dailyChange = dailyChange + (Cells(i, 4) - Cells(i, 5))

        ' Average change
        days = (i - strt) + 1
        avgChange = dailyChange / days
        
        ' start of the next stock ticker
        strt = i + 1

        ' print the results to a seperate worksheet
        ws2.Range("A" & 2 + j).Value = Cells(i, 1).Value
        ws2.Range("B" & 2 + j).Value = Round(change, 2)
        ws2.Range("C" & 2 + j).Value = "%" & perChange
        ws2.Range("D" & 2 + j).Value = avgChange
        ws2.Range("E" & 2 + j).Value = ttl

        ' colors positives green and negatives red
        Select Case change
            Case Is > 0
               ws2.Range("B" & 2 + j).Interior.ColorIndex = 4
            Case Is < 0
                ws2.Range("B" & 2 + j).Interior.ColorIndex = 3
            Case Else
                ws2.Range("B" & 2 + j).Interior.ColorIndex = 0
        End Select
        
        ' reset variables for new stock ticker
        ttl = 0
        change = 0
        j = j + 1
        days = 0
        dailyChange = 0
   
    Else
        ttl = ttl + Cells(i, 7).Value
        change = change + (Cells(i, 6) - Cells(i, 3))

        ' change in high and low
        dailyChange = dailyChange + (Cells(i, 4) - Cells(i, 5))


    End If
Next i

' take the max and min and place them in a separate part in the worksheet
ws2.Cells(2, 8) = WorksheetFunction.Max(ws2.Range("E2:E" & rowcount))
ws2.Cells(5, 8) = "%" & WorksheetFunction.Max(ws2.Range("C2:C" & rowcount)) * 100
ws2.Cells(8, 8) = "%" & WorksheetFunction.Min(ws2.Range("C2:C" & rowcount)) * 100
ws2.Cells(11, 8) = WorksheetFunction.Max(ws2.Range("D2:D" & rowcount))


' returns one less because header row not a factor
volNo = WorksheetFunction.Match(WorksheetFunction.Max(ws2.Range("E2:E" & rowcount)), ws2.Range("E2:E" & rowcount), 0)
incrNo = WorksheetFunction.Match(WorksheetFunction.Max(ws2.Range("C2:C" & rowcount)), ws2.Range("C2:C" & rowcount), 0)
dcrNo = WorksheetFunction.Match(WorksheetFunction.Min(ws2.Range("C2:C" & rowcount)), ws2.Range("C2:C" & rowcount), 0)
avgNo = WorksheetFunction.Match(WorksheetFunction.Max(ws2.Range("D2:D" & rowcount)), ws2.Range("D2:D" & rowcount), 0)


' final ticker symbol for  total, greatest % of increase and decrease, and average
ws2.Cells(2, 9) = ws2.Cells(volNo + 1, 1)
ws2.Cells(5, 9) = ws2.Cells(incrNo + 1, 1)
ws2.Cells(8, 9) = ws2.Cells(dcrNo + 1, 1)
ws2.Cells(11, 9) = ws2.Cells(avgNo + 1, 1)


End Sub

