Attribute VB_Name = "Module1"
Sub tracker()

For Each ws In Worksheets

Dim ticker As String
Dim open_tick As Double
Dim high_tick As Double
Dim low_tick As Double
Dim close_tick As Double
Dim Volume As Integer

Summary_Table_Row = (2)
ws.Cells(1, "J") = "Ticker"
ws.Cells(1, "K") = "Total Volume"
ws.Cells(1, "L") = "Dollar Change"
ws.Cells(1, "M") = "Percent Change"
start_value = ws.Cells(2, "C").Value

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow

If ws.Cells(i + 1, "A").Value <> ws.Cells(i, "A").Value Then
    
    ticker = ws.Cells(i, "A").Value
    totalcount = totalcount + ws.Cells(i, "G").Value
    
    end_value = ws.Cells(i, "F").Value
    'Debug.Print ("This is the start" + Str(start_value))
    'Debug.Print ("This is the end" + Str(end_value))
    
    
    
    
    ws.Range("J" & Summary_Table_Row).Value = ticker
    ws.Range("K" & Summary_Table_Row).Value = totalcount
    ws.Range("L" & Summary_Table_Row).Value = (end_value - start_value)
    If start_value > 0 Then
    ws.Range("M" & Summary_Table_Row).Value = ((end_value - start_value) / start_value)
        End If
    Summary_Table_Row = Summary_Table_Row + 1
        totalcount = 0
    start_value = ws.Cells(i + 1, "C").Value
            
    
    Else
        totalcount = totalcount + ws.Cells(i, "G").Value
        
    End If
    
    Next i

'Found this code at bluepecantraining.com

Dim rg As Range
Dim cond1 As FormatCondition, cond2 As FormatCondition, cond3 As FormatCondition
Set rg = Range("L2", Range("L2").End(xlDown))

Set cond1 = rg.FormatConditions.Add(xlCellValue, xlGreater, 0)
Set cond2 = rg.FormatConditions.Add(xlCellValue, xlEqual, 0)
Set cond3 = rg.FormatConditions.Add(xlCellValue, xlLess, 0)

With cond1
.Interior.Color = vbGreen
End With

With cond2
.Interior.Color = vbYellow
End With

With cond3
.Interior.Color = vbRed
End With

Dim rg1 As Range
Dim cond4 As FormatCondition, cond5 As FormatCondition, cond6 As FormatCondition
Set rg1 = Range("m2", Range("m2").End(xlDown))

Set cond4 = rg1.FormatConditions.Add(xlCellValue, xlGreater, 0)
Set cond5 = rg1.FormatConditions.Add(xlCellValue, xlEqual, 0)
Set cond6 = rg1.FormatConditions.Add(xlCellValue, xlLess, 0)

With cond4
.Interior.Color = vbGreen
End With

With cond5
.Interior.Color = vbYellow
End With

With cond6
.Interior.Color = vbRed
End With


Next ws

End Sub
