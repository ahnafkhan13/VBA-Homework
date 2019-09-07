Sub market_analysis():

    Dim ws As Worksheet
    Dim tick As String
    Dim openT As Double
    Dim closeT As Double
    Dim vol As Double
    Dim rowCounter As Integer

    For Each ws In ThisWorkbook.Worksheets
    ws.Activate
  

    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"

    openT = Range("C2").Value
    vol = 0
    rowCounter = 2

    LR = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To LR
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            tick = Cells(i, 1).Value
            Cells(rowCounter, 9).Value = tick
            closeT = Cells(i, 6).Value
            Cells(rowCounter, 10).Value = closeT - openT
            
            If Cells(rowCounter, 10).Value > 0 Then
                Cells(rowCounter, 10).Interior.Color = RGB(0, 255, 0)
            Else
                Cells(rowCounter, 10).Interior.Color = RGB(255, 0, 0)
            End If
            
            If openT > 0 Then
                Cells(rowCounter, 11).Value = (closeT / openT) - 1
                Cells(rowCounter, 11).NumberFormat = "0.00%"
            Else
                Cells(rowCounter, 11).Value = 0
            End If
            
            vol = vol + Cells(i, 7).Value
            Cells(rowCounter, 12).Value = vol
            rowCounter = rowCounter + 1
            vol = 0
            openT = Cells(i + 1, 3).Value
          Else
            vol = vol + Cells(i, 7).Value
        End If
    Next i
    
    LR2 = Cells(Rows.Count, 9).End(xlUp).Row
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 17).NumberFormat = "0.00%"
    
    For i = 2 To LR2
        If Cells(i, 11).Value > Cells(2, 17).Value Then
            Cells(2, 17).Value = Cells(i, 11).Value
            Cells(2, 16).Value = Cells(i, 9).Value
        End If
        If Cells(i, 11).Value < Cells(3, 17).Value Then
            Cells(3, 17).Value = Cells(i, 11).Value
            Cells(3, 16).Value = Cells(i, 9).Value
        End If
        If Cells(i, 12).Value > Cells(4, 17).Value Then
            Cells(4, 17).Value = Cells(i, 12).Value
            Cells(4, 16).Value = Cells(i, 9).Value
        End If
    Next i
    Columns("O:O").EntireColumn.AutoFit
    Next ws
End Sub



