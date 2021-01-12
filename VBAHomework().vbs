Sub VBAHomework()

For Each ws In Worksheets

Dim WorksheetName As String
Dim Ticker As String
Dim RowNumber As Integer
Dim CloseValue As Double
Dim OpenValue As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim Volume As Variant

WorksheetName = ws.Name

RowNumber = 2
OpenValue = ws.Cells(2, 3)
Volume = 0

    

For i = 2 To 705714
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    
    Ticker = ws.Cells(i, 1).Value
    ClosedValue = ws.Cells(i, 6)
    YearlyChange = ClosedValue - OpenValue
    If OpenValue = 0 And YearlyChange = 0 Then
    PercentChange = 0
    Else
    PercentChange = (YearlyChange / OpenValue)
    End If
    Volume = ws.Cells(i, 7) + Volume
    
    
    
    ws.Cells(RowNumber, 9) = Ticker
    ws.Cells(RowNumber, 10) = YearlyChange
    ws.Cells(RowNumber, 11) = PercentChange
    ws.Cells(RowNumber, 12) = Volume
    
    Volume = 0
    
    
    
    RowNumber = RowNumber + 1
    OpenValue = ws.Cells(i + 1, 3)
    
    Else
    
    Volume = ws.Cells(i, 7) + Volume
    
    If ws.Cells(RowNumber, 10) >= 0 Then
    ws.Cells(RowNumber, 10).Interior.ColorIndex = 4
    Else
    ws.Cells(RowNumber, 10).Interior.ColorIndex = 3
    
    
    
    End If

    End If
    
 
    Next i
    
    For j = 2 To 2836
    
    Dim Max As Double
    Dim Min As Double
    Dim TotalValue As Variant
    
    If ws.Cells(j, 11) > Max Then
    Max = ws.Cells(j, 11)
    Ticker = ws.Cells(j, 9)
    
    End If
    
    If ws.Cells(j, 11) < Min Then
    Min = ws.Cells(j, 11)
    Ticker2 = ws.Cells(j, 9)
    
    End If
    
     
    If ws.Cells(j, 12) > TotalValue Then
    TotalValue = ws.Cells(j, 12)
    Ticker3 = ws.Cells(j, 9)
    
    End If
    
    
    ws.Cells(2, 15) = Ticker
    ws.Cells(3, 15) = Ticker2
    ws.Cells(4, 15) = Ticker3
    ws.Cells(2, 16) = Max
    ws.Cells(3, 16) = Min
    ws.Cells(4, 16) = TotalValue
    
    Next j
    
    Next ws
    
End Sub
