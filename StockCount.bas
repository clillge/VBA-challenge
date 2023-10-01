Attribute VB_Name = "StockCount"
Sub StockCount()

For Each ws In Worksheets

    Dim total As LongLong
    Dim openprice, closeprice As Double
    Dim ticker As Integer
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ticker = 2
    total = 0

    openprice = Cells(2, 3).Value

    For i = 2 To LastRow

        total = total + Cells(i, 7).Value
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            closeprice = ws.Cells(i, 6).Value
        
            ws.Cells(ticker, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(ticker, 10).Value = (closeprice - openprice)
            ws.Cells(ticker, 11).Value = (ws.Cells(ticker, 10).Value / openprice)
            ws.Cells(ticker, 12).Value = total
        
            openprice = ws.Cells(i + 1, 3).Value
        
            total = 0
        
            ticker = ticker + 1
        
        End If
    
    Next i

    ws.Range("P2").Value = Application.WorksheetFunction.Max(ws.Range("K2:K3001").Value)
    ws.Range("P3").Value = Application.WorksheetFunction.Min(ws.Range("K2:K3001").Value)
    ws.Range("P4").Value = Application.WorksheetFunction.Max(ws.Range("L2:L3001").Value)

    For i = 2 To 3001

        If ws.Cells(i, 11).Value = ws.Range("P2").Value Then
            ws.Range("O2").Value = ws.Cells(i, 9).Value
        ElseIf ws.Cells(i, 11).Value = ws.Range("P3").Value Then
            ws.Range("O3").Value = ws.Cells(i, 9).Value
        ElseIf ws.Cells(i, 12).Value = ws.Range("P4").Value Then
            ws.Range("O4").Value = ws.Cells(i, 9).Value
        End If
    Next i
    
MsgBox (ws.Name)
MsgBox (LastRow)

Next ws

End Sub
