Sub stockprices2()

    Dim i, j As Integer
    Dim nrows, nrows2 As Integer
    Dim openVal, closeVal As Integer
    Dim ws As Worksheet
    Dim totalVol As Variant
    Dim perIncrease, perDecrease As Double
    Dim greatVol As Variant
    Dim tickInc, tickDec, tickVol As String
    
    
    'Go through each worksheet
    For Each ws In Worksheets
        ws.Activate

        'clean worksheets to run
        Columns("I:Q").Delete Shift:=xlToLeft
    
        'headers
        Range("I1") = "Ticker"
        'set first values to create a range
        Range("I2") = Range("A2")
        Range("J1") = "Yearly Change"
        Range("K1") = "Percent Change"
        Range("L1") = "Total Stock Volume"
        Range("O2") = "Greatest % Increase"
        Range("O3") = "Greatest % Decrease"
        Range("O4") = "Greatest Total Volume"
        Range("P1") = "Ticker"
        Range("Q1") = "Value"
        
        openVal = Range("C2").Value
        totalVol = 0
        
        'row counter
        nrows = Range("A1", Range("A1").End(xlDown)).Rows.Count
        nrows2 = Range("I1", Range("I1").End(xlDown)).Rows.Count
        
        'assign rows once they change value
        For i = 2 To nrows
            totalVol = totalVol + Cells(i, 7).Value
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                Cells(nrows2 + 1, 9).Value = Cells(i + 1, 1).Value
                Cells(nrows2, 10).Value = Cells(i, 6).Value - openVal
                'color format for change
                If Cells(nrows2, 10).Value >= 0 Then
                    'green
                    Cells(nrows2, 10).Interior.ColorIndex = 4
                Else
                    'red
                    Cells(nrows2, 10).Interior.ColorIndex = 3
                End If
                '0 check for open values. Causes error if not
                If openVal = 0 Then
                    Cells(nrows2, 11).Value = 0
                Else
                    Cells(nrows2, 11).Value = (Cells(i, 6).Value - openVal) / openVal
                End If
                'total counter
                Cells(nrows2, 12).Value = totalVol
                totalVol = 0
                'new open value for next item
                openVal = Cells(i + 1, 3).Value
                nrows2 = nrows2 + 1
            End If
            
        Next i
        
        perIncrease = 0
        greatVol = 0
        
        'assigns new values if the cell is greater than current
        For j = 2 To nrows2 - 1
            If Cells(j, 11).Value > perIncrease Then
                perIncrease = Cells(j, 11).Value
                tickInc = Cells(j, 9).Value
            End If
            If Cells(j, 12).Value > greatVol Then
                greatVol = Cells(j, 12).Value
                tickVol = Cells(j, 9).Value
            End If
        Next j
        
        'starts at max
        perDecrease = perIncrease

        'assigns value if cell is less than current       
        For j = 2 To nrows2 - 1
            If Cells(j, 11).Value < perDecrease Then
                perDecrease = Cells(j, 11).Value
                tickDec = Cells(j, 9).Value
            End If
        Next j
       'Headers for max and min
        Range("P2") = tickInc
        Range("Q2") = perIncrease
        Range("P3") = tickDec
        Range("Q3") = perDecrease
        Range("P4") = tickVol
        Range("Q4") = greatVol
        
        'format
        Range("J1", Range("J1").End(xlDown)).NumberFormat = "0.00"
        Range("K1", Range("K1").End(xlDown)).NumberFormat = "0.00%"
        Range("Q2", "Q3").NumberFormat = "0.00%"
        Cells.EntireColumn.AutoFit
        Range("A1").Select
    Next ws
    
End Sub