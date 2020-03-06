Sub Reset():
    For Each sheet In Worksheets
        sheet.Columns("I:Q").Clear
    Next
End Sub

Sub PrintHeader(ByVal sheet As Worksheet):
    sheet.Range("I1").Value = "Ticker"
    sheet.Range("J1").Value = "Yearly Change"
    sheet.Range("K1").Value = "Percent Change"
    sheet.Range("L1").Value = "Total Stock Volume"
    
    sheet.Range("O1").Value = "Ticker"
    sheet.Range("P1").Value = "Value"
    sheet.Range("N2").Value = "Greatest % Increase"
    sheet.Range("N3").Value = "Greatest % Descrease"
    sheet.Range("N4").Value = "Greatest Total Volume"
    
    ' debug information
    If sheet.Index = 1 Then
        sheet.Range("N25").Value = "Debug Information:"
        sheet.Range("N26").Value = "Worksheet Name"
        sheet.Range("O26").Value = "RowIndex"
        sheet.Range("P26").Value = "Ticker"
        sheet.Range("Q26").Value = "TotalVolume"
    End If
End Sub

Sub FormatSummaryTable(ByVal sheet As Worksheet):
    ' Red highlight for negative change
    With sheet.Range("J2", sheet.Range("J2").End(xlDown)).FormatConditions _
            .Add(xlCellValue, xlLess, 0)
        With .Interior
            .ColorIndex = 3
        End With
    End With
    
    ' Green highlight for positive change
    With sheet.Range("J2", sheet.Range("J2").End(xlDown)).FormatConditions _
            .Add(xlCellValue, xlGreater, 0)
        With .Interior
            .ColorIndex = 4
        End With
    End With
    
    ' Percentage format for Percent Change column
    sheet.Columns("K").NumberFormat = "0.00%"
    sheet.Range("P2:P3").NumberFormat = "0.00%"
    sheet.Columns("L").NumberFormat = "#,##0"
    sheet.Range("P4").NumberFormat = "#,##0"
    
    ' format heading and ticker
    ' sheet.Range("I1:L1").Style = "Accent3"
    sheet.Range("I1:L1").Font.Bold = True
    sheet.Range("I1:L1").Interior.ColorIndex = 15
    sheet.Range("I1:L1").Font.Size = 14
    Columns("I").Font.Bold = True
    
    sheet.Range("I1:Q30").Columns.AutoFit
End Sub

Sub GetStats(ByVal sheet As Worksheet):
    ' Greatest % Increase Value
    sheet.Range("P2").Value = WorksheetFunction.Max(sheet.Range("K2", sheet.Range("K2").End(xlDown)))
        
    ' Greatest % Increase Ticker
    sheet.Range("O2").Value = Cells(sheet.Range("K2", sheet.Range("K2").End(xlDown)).Find(sheet.Range("P2").Value).Row, 9).Value
    
    ' Greatest % Decrease Value
    sheet.Range("P3").Value = WorksheetFunction.Min(sheet.Range("K2", sheet.Range("K2").End(xlDown)))
    
    ' Greatest % Decrease Ticker
    sheet.Range("O3").Value = Cells(sheet.Range("K2", sheet.Range("K2").End(xlDown)).Find(sheet.Range("P3").Value).Row, 9).Value
    
    ' Greatest Total Volume Value
    sheet.Range("P4").Value = WorksheetFunction.Max(sheet.Range("L2", sheet.Range("L2").End(xlDown)))
    
    ' Greatest Total Volume Ticker
    sheet.Range("O4").Value = Cells(sheet.Range("L2", sheet.Range("L2").End(xlDown)).Find(sheet.Range("P4").Value).Row, 9).Value
End Sub


' Main program
Sub GetSummary():
    
    Dim previousTicker, currentTicker As String
    Dim RowIndex As Long
    Dim summaryRowIndex As Long
    Dim beginingPrice, endingPrice As Double
    Dim totalVolume As LongLong

    For Each sheet In Worksheets
        Call PrintHeader(sheet)
    
        ' initial value for each sheet
        summaryRowIndex = 2
        RowIndex = 2
        previousTicker = sheet.Cells(RowIndex, 1).Value
        currentTicker = previousTicker
        beginningPrice = sheet.Cells(RowIndex, 3).Value
        totalVolume = 0
        endingPrice = 0
        
        
        Do While Not currentTicker = ""
        
            ' debug exit after limited tickers
            'If summaryRowIndex = 31 Then
            '    Exit Do
            'End If
            
     
            
            totalVolume = totalVolume + sheet.Cells(RowIndex, 7).Value
            If (previousTicker <> currentTicker) Then
                
                ' write to summary table
                sheet.Cells(summaryRowIndex, 9).Value = previousTicker
    
                ' get ending price from previous row
                endingPrice = sheet.Cells(RowIndex - 1, 6)
                sheet.Cells(summaryRowIndex, 10).Value = endingPrice - beginningPrice
                
                If beginningPrice = 0 Then
                    sheet.Cells(summaryRowIndex, 11).Value = 1
                Else
                    sheet.Cells(summaryRowIndex, 11).Value = (endingPrice - beginningPrice) / beginningPrice
                End If
                
                sheet.Cells(summaryRowIndex, 12).Value = totalVolume
                ' move to next row in the Summary table
                summaryRowIndex = summaryRowIndex + 1
    
                
                beginningPrice = sheet.Cells(RowIndex, 3).Value
                totalVolume = sheet.Cells(RowIndex, 7).Value
            End If
    
            ' move to next row in data table
            RowIndex = RowIndex + 1
            
            ' debug information
            Worksheets(1).Range("N27").Value = sheet.Name
            Worksheets(1).Range("O27").Value = RowIndex
            Worksheets(1).Range("P27").Value = currentTicker
            Worksheets(1).Range("Q27").Value = totalVolume
            
            previousTicker = currentTicker
            currentTicker = sheet.Cells(RowIndex, 1).Value
            
            ' at the end of current worksheet
            If currentTicker = "" Then
                ' get ending price from previous row
                endingPrice = sheet.Cells(RowIndex - 1, 6)
                
                sheet.Cells(summaryRowIndex, 9).Value = previousTicker
                sheet.Cells(summaryRowIndex, 10).Value = endingPrice - beginningPrice
                If beginningPrice = 0 Then
                    sheet.Cells(summaryRowIndex, 11).Value = 1
                Else
                    sheet.Cells(summaryRowIndex, 11).Value = (endingPrice - beginningPrice) / beginningPrice
                End If
                sheet.Cells(summaryRowIndex, 12).Value = totalVolume
                ' move to next row in the Summary table
                summaryRowIndex = summaryRowIndex + 1
                
            End If
        Loop
        
        
        Call GetStats(sheet)
        Call FormatSummaryTable(sheet)
    Next
    
End Sub

