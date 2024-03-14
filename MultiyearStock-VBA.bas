Attribute VB_Name = "Module1"
Sub MultipleYearStockData()
    Dim ws As Worksheet
    Dim WorksheetName As String
    Dim LastRowA As Long, LastRowI As Long
    Dim i As Long, j As Long, TickCount As Long
    Dim PerChange As Double, GreatIncr As Double, GreatDecr As Double, GreatVol As Double
    
    For Each ws In Worksheets
        WorksheetName = ws.Name
        
        With ws
            ' Set up column headers
            .Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
            .Range("O2:O4").Value = Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume")
            .Range("P1:Q1").Value = Array("Ticker", "Value")
            TickCount = 2
            j = 2
            
            ' Find the last non-blank row in column A
            LastRowA = .Cells(.Rows.Count, 1).End(xlUp).Row
            
            ' Loop through all rows
            For i = 2 To LastRowA
                If .Cells(i + 1, 1).Value <> .Cells(i, 1).Value Then
                    ' Write ticker in column I (#9)
                    .Cells(TickCount, 9).Value = .Cells(i, 1).Value
                    
                    ' Calculate and write yearly change in column J (#10)
                    .Cells(TickCount, 10).Value = .Cells(i, 6).Value - .Cells(j, 3).Value
                    
                    ' Conditional formatting for yearly change
                    If .Cells(TickCount, 10).Value < 0 Then
                        .Cells(TickCount, 10).Interior.ColorIndex = 3 ' Set cell background color to red
                    Else
                        .Cells(TickCount, 10).Interior.ColorIndex = 4 ' Set cell background color to green
                    End If
                    
                    ' Calculate and write percent change in column K (#11)
                    If .Cells(j, 3).Value <> 0 Then
                        PerChange = ((.Cells(i, 6).Value - .Cells(j, 3).Value) / .Cells(j, 3).Value)
                        .Cells(TickCount, 11).Value = Format(PerChange, "Percent")
                    Else
                        .Cells(TickCount, 11).Value = Format(0, "Percent")
                    End If
                    
                    ' Calculate and write total volume in column L (#12)
                    .Cells(TickCount, 12).Value = WorksheetFunction.Sum(.Range(.Cells(j, 7), .Cells(i, 7)))
                    
                    ' Increase TickCount by 1
                    TickCount = TickCount + 1
                    
                    ' Set new start row of the ticker block
                    j = i + 1
                End If
            Next i
            
            ' Find last non-blank cell in column I
            LastRowI = .Cells(.Rows.Count, 9).End(xlUp).Row
            
            ' Prepare for summary
            GreatVol = .Cells(2, 12).Value
            GreatIncr = .Cells(2, 11).Value
            GreatDecr = .Cells(2, 11).Value
            
            ' Loop for summary
            For i = 2 To LastRowI
                ' For greatest total volume
                If .Cells(i, 12).Value > GreatVol Then
                    GreatVol = .Cells(i, 12).Value
                    .Cells(4, 16).Value = .Cells(i, 9).Value
                End If
                
                ' For greatest increase
                If .Cells(i, 11).Value > GreatIncr Then
                    GreatIncr = .Cells(i, 11).Value
                    .Cells(2, 16).Value = .Cells(i, 9).Value
                End If
                
                ' For greatest decrease
                If .Cells(i, 11).Value < GreatDecr Then
                    GreatDecr = .Cells(i, 11).Value
                    .Cells(3, 16).Value = .Cells(i, 9).Value
                End If
            Next i
            
            ' Write summary results in ws.Cells
            .Cells(2, 17).Value = Format(GreatIncr, "Percent")
            .Cells(3, 17).Value = Format(GreatDecr, "Percent")
            .Cells(4, 17).Value = Format(GreatVol, "Scientific")
            
            ' Adjust column width automatically
            .Columns("A:Z").AutoFit
        End With
    Next ws
End Sub

