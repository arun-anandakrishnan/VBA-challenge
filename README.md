# VBA-challenge

'TickerSymbol
Sub StockSymbols()
    Dim lastRow As Long
    Dim StockSymbols As Object
    Dim cell As Range
    Dim symbol As String
    Dim outputRow As Long
    
    Set StockSymbols = CreateObject("Scripting.Dictionary")
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    outputRow = 2 ' Output starting row in Column I
    
    For Each cell In Range("A2:A" & lastRow)
        symbol = cell.Value
        
        If Not StockSymbols.Exists(symbol) Then
            StockSymbols(symbol) = True
            Cells(outputRow, "I").Value = symbol
            outputRow = outputRow + 1
        End If
    Next cell
End Sub

'Total Volume
Sub TotalStockVolume()
    Dim lastRow As Long
    Dim ticker As String
    Dim total As Double
    Dim currentRow As Long

   ' Find the last row in column A
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row

   ' Start from row 2
    ticker = Range("A2").Value
    total = 0
    currentRow = 2

   ' Loop through the rows
    For i = 2 To lastRow
        If Range("A" & i).Value = ticker Then
            ' Sum the totals for
            total = total + Range("G" & i).Value
        Else
            ' Output the total in column L and reset variables
            Range("L" & currentRow).Value = total
            currentRow = currentRow + 1
            ticker = Range("A" & i).Value
            total = Range("G" & i).Value
        End If
    Next i

    ' Output the last total in column L
    Range("L" & currentRow).Value = total
End Sub
