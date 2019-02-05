Sub Stocks()

'STEP 1: Definitions
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
    
Dim openPrice As Double
Dim closePrice As Double
Dim yearlyChange As Double
Dim tickerName As String
Dim percentChange As Double
Dim Volume As Double

Volume = 0
Dim Row As Double
Row = 2
Dim Column As Integer
Column = 1
Dim i As Long
lastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

'STEP 2: Define Headers and format cells
    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "Yearly Change"
    Cells(1, "K").Value = "Percent Change"
    Cells(1, "L").Value = "Total Stock Volume"
    Cells(2, "O").Value = "Greatest % Increase"
    Cells(3, "O").Value = "Greatest % Decrease"
    Cells(4, "O").Value = "Greatest Total Volume"
    Cells(1, "P").Value = "Ticker"
    Cells(1, "Q").Value = "Value"
    
    Cells(2, "Q").NumberFormat = "0.00%"
    Cells(3, "Q").NumberFormat = "0.00%"
     
'STEP 3: Loop through Ticker Symbols
        
        For i = 2 To lastRow
            'Only count if Same Symbol Stops Counting if the Symbol Changes
            If Cells(i + 1, "A").Value <> Cells(i, "A").Value Then
                'Ticker name
                tickerName = Cells(i, "A").Value
                Cells(Row, "I").Value = tickerName
                
                'Yearly Change
                openPrice = Cells(2, "C").Value
                closePrice = Cells(i, "F").Value
                yearlyChange = closePrice - openPrice
                Cells(Row, "J").Value = yearlyChange
                
                'Percent Change
                If (openPrice = 0 And closePrice = 0) Then
                    percentChange = 0
                ElseIf (openPrice = 0 And closePrice <> 0) Then
                    percentChange = 1
                Else
                    percentChange = yearlyChange / openPrice
                    Cells(Row, "K").Value = percentChange
                End If
                
                'Total Stock Volumn
                Volume = Volume + Cells(i, "G").Value
                Cells(Row, "L").Value = Volume
               
                openPrice = Cells(i + 1, "C")
                'Counters
                Volume = 0
                Row = Row + 1
                Cells(Row, "K").NumberFormat = "0.00%"
            
            Else
                Volume = Volume + Cells(i, "G").Value
            End If
        Next i
        
'STEP 4: Yearly Change Color
        yearLastRow = WS.Cells(Rows.Count, "I").End(xlUp).Row
        For j = 2 To yearLastRow
            If (Cells(j, "J").Value > 0 Or Cells(j, "J").Value = 0) Then
                Cells(j, "J").Interior.ColorIndex = 10
            ElseIf Cells(j, "J").Value < 0 Then
                   Cells(j, "J").Interior.ColorIndex = 3
            End If
        Next j
        

'STEP 5: Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
        For k = 2 To yearLastRow
            If Cells(k, "K").Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & yearLastRow)) Then
                Cells(2, "P").Value = Cells(k, "I").Value
                Cells(2, "Q").Value = Cells(k, "K").Value  
            ElseIf Cells(k, "K").Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & yearLastRow)) Then
                Cells(3, "P").Value = Cells(k, "I").Value
                Cells(3, "Q").Value = Cells(k, "K").Value
            ElseIf Cells(k, "L").Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & yearLastRow)) Then
                Cells(4, "P").Value = Cells(k, "I").Value
                Cells(4, "Q").Value = Cells(k, "L").Value
            End If
        Next k
        
    Next WS
        
End Sub

