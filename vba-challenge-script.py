Sub StockData()

Dim ticker As String
Dim volume As Integer
Dim YearOpen As Double
Dim YearClose As Double
Dim YearlyChange As Double
Dim summary_table As Integer
Dim ws As Worksheet







'headers'

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Value"

For Each ws In ThisWorkbook.Worksheets

 

''values'

        ticker = ws.Cells(i, 1).Value
        volume = ws.Cells(i, 7).Value
        YearOpen = ws.Cells(i, 3).Value
        YearClose = ws.Cells(i, 6).Value
        
        
 'the formula'
    

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
     'integers for loop'
     
    summary_table = 2
    
    'loop time again'
    For i = 2 To ws.UsedRange.Rows.Count
    
    
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        ticker = ws.Cells(i, 1).Value
        volume = ws.Cells(i, 7).Value
        YearOpen = ws.Cells(i, 3).Value
        YearClose = ws.Cells(i, 6).Value
        
    
    
    YearlyChange = YearClose - YearOpen
    PercentChange = (YearClose - YearOpen) / YearClose

'summary'

    ws.Cells(summary_table, 9).Value = ticker
    ws.Cells(summary_table, 10).Value = YearlyChange
    ws.Cells(summary_table, 11).Value = PercentChange
    ws.Cells(summary_table, 12).Value = volume
    summary_table = summary_table + 1

    volume = 0
End If

'close the loop'
    Next i
ws.Columns("J").NumberFormat = "0.00%"

Dim rg As Range
Dim g As Long
Dim c As Long
Dim ColorCells As Range

Set rg = ws.Range("J2", Range("J2").End(x1down))
    c = rg.Cells.Count
    
For g = 1 To c
Set ColorCells = rg(g)
Select Case ColorCells
    Case Is >= 0
        With ColorCells
        .Interior.Color = vbGreen
        End With
    Case Is < 0
        With ColorCells
        .Interior.Color = vbRed
        End With
    End Select
    Next g
End If
Next ws

End Sub

