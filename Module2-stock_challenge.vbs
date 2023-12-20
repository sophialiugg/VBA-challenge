Sub ticker()
    For Each ws In Worksheets
        Dim ticker As String
        Dim yearlychange As Double
        Dim percentchange As Double
        Dim totalvolume As Double
        totalvolume = 0
        Dim SummaryTableRow As Integer
        SummaryTableRow = 2
        
        
        openprice = ws.Cells(2, 3).Value
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
       
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        
        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            ws.Cells(SummaryTableRow, 9).Value = ticker
            
            closingprice = ws.Cells(i, 6).Value
            yearlychange = ws.Cells(i, 6).Value - openprice
            ws.Cells(SummaryTableRow, 10).Value = yearlychange
            If ws.Cells(SummaryTableRow, 10).Value < 0 Then
                ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
            End If
            
                If openprice <> 0 Then
                percentchange = (ws.Cells(i, 6).Value - openprice) / openprice
                ws.Cells(SummaryTableRow, 11).Value = FormatPercent(percentchange, 2)
                
                Else
                ws.Cells(SummaryTableRow, 11).Value = 0
                
                End If
            
                
                If ws.Cells(SummaryTableRow, 11).Value < 0 Then
                    ws.Cells(SummaryTableRow, 11).Interior.ColorIndex = 3
                Else
                    ws.Cells(SummaryTableRow, 11).Interior.ColorIndex = 4
                End If
            
            
            openprice = ws.Cells(i + 1, 3).Value
            
            totalvolume = totalvolume + ws.Cells(i, 7).Value
            ws.Cells(SummaryTableRow, 12).Value = totalvolume
            totalvolume = 0
            
            SummaryTableRow = SummaryTableRow + 1
            Else
            totalvolume = totalvolume + ws.Cells(i, 7).Value
            
            End If
            
        Next i

        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        MaxValue = Range("K2")
        MinValue = Range("K2")
        MaxVolumeValue = Range("L2")
        Maxindex = 2
        Minindex = 2
        MaxVolumeindex = 2
        
        
        For i = 2 To lastrow2
        
            If ws.Cells(i, 11).Value > MaxValue Then
            MaxValue = ws.Cells(i, 11).Value
            Maxindex = i
            ws.Range("P2").Value = ws.Cells(Maxindex, 9).Value
            ws.Range("Q2").Value = FormatPercent(MaxValue, 2)
            End If
         
            
            If ws.Cells(i, 11).Value < MinValue Then
            MinValue = ws.Cells(i, 11).Value
            Minindex = i
            ws.Range("P3").Value = ws.Cells(Minindex, 9).Value
            ws.Range("Q3").Value = FormatPercent(MinValue, 2)
            
            End If
            
            If ws.Cells(i, 12).Value > MaxVolumeValue Then
            MaxVolumeValue = ws.Cells(i, 12).Value
            MaxVolumeindex = i
            ws.Range("P4").Value = ws.Cells(MaxVolumeindex, 9).Value
            ws.Range("Q4").Value = MaxVolumeValue

            End If
        Next i
        
        Next ws
        
End Sub