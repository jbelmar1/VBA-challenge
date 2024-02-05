Attribute VB_Name = "Module1"
Sub Multiple_Year_Stock_Data()

    'LOOPING THROUGH ALL WORKSHEETS IN THE WORKBOOK
    For Each ws In Worksheets

        'VARIABLES
        Dim Worksheet_Name As String
        Dim Ticker_Value As Long
        Dim LastRow_TV As Long
        Dim LastRow_T As Long
        Dim LastRow_YC As Long
        Dim Percent_Change As Double
        Dim MaxValue_GI As Double
        Dim MaxValue_GD As Double
        Dim MaxValue_TSV As Double
        Dim i As Long
        Dim j As Long

        Worksheet_Name = ws.Name

        'HEADERS
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        'TICKER
        Ticker_Value = 2
        j = 2
        LastRow_TV = ws.Cells(Rows.Count, 1).End(xlUp).Row
            For i = 2 To LastRow_TV
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'NEED TO PUT THE TICKER NAME IN COLUMN I
                ws.Cells(Ticker_Value, 9).Value = ws.Cells(i, 1).Value
                'YEARLY CHANGE
                ws.Cells(Ticker_Value, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                    'PERCENT CHANGE
                    If ws.Cells(j, 3).Value = 0 Then
                    ws.Cells(Ticker_Value, 11).Value = 0
                    Else
                    ws.Cells(Ticker_Value, 11).Value = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    ws.Cells(Ticker_Value, 11).NumberFormat = "0.00%"
                    End If
                ws.Cells(Ticker_Value, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                Ticker_Value = Ticker_Value + 1
                j = i + 1
    
                End If
    
            Next i

        '2nd CALCULATIONS
        LastRow_T = ws.Cells(Rows.Count, 9).End(xlUp).Row
        'MAXVALUES
        MaxValue_GI = ws.Cells(2, 11).Value
        MaxValue_GD = ws.Cells(2, 11).Value
        MaxValue_TSV = ws.Cells(2, 12).Value
            For i = 2 To LastRow_T
                'MAXVALUE for GREATEST DECREASE
                If ws.Cells(i, 11).Value > MaxValue_GI Then
                MaxValue_GI = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                Else
                MaxValue_GI = MaxValue_GI
                End If
                ws.Cells(2, 17).Value = MaxValue_GI
                ws.Cells(2, 17).NumberFormat = "0.00%"
                'MAXVALUE for GREATEST DECREASE
                If ws.Cells(i, 11).Value < MaxValue_GD Then
                MaxValue_GD = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                Else
                MaxValue_GD = MaxValue_GD
                End If
                ws.Cells(3, 17).Value = MaxValue_GD
                ws.Cells(3, 17).NumberFormat = "0.00%"
                'MAXVALUE for TOTAL STOCK VOLUME
                If ws.Cells(i, 12).Value > MaxValue_TSV Then
                MaxValue_TSV = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                Else
                MaxValue_TSV = MaxValue_TSV
                End If
                ws.Cells(4, 17).Value = MaxValue_TSV
            Next i

'CONDITIONAL FORMATTING - GREEN FOR POSTIVE & RED FOR NEGATIVE

        LastRow_YC = ws.Cells(Rows.Count, 10).End(xlUp).Row
            For i = 2 To LastRow_YC
                If ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
                ElseIf ws.Cells(i, 10).Value >= o Then
                ws.Cells(i, 10).Interior.ColorIndex = 4

                ElseIf ws.Cells(x, 10).Value < 0 Then
                ws.Cells(i, 11).Interior.ColorIndex = 3
                ElseIf s.Cells(i, 10).Value >= o Then
                ws.Cells(i, 11).Interior.ColorIndex = 4
                End If
            Next i
    
        ws.UsedRange.EntireColumn.AutoFit
        ws.UsedRange.EntireRow.AutoFit
        
    Next ws

End Sub

