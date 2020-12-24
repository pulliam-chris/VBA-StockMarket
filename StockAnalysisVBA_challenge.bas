Attribute VB_Name = "Module2"
Sub StockAnalysis_Challenge()

Dim summaryRowCount As Long

' Loop through all sheets
    For Each ws In Worksheets
        
        'MsgBox (ws.Name)
        
        'Instert Calculation Table Headers
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Row Count to determine min, max change and max volume
        summaryRowCount = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Find biggest increase and decrease changes in the summary along with largest total volume
        ws.Range("Q2").Formula = "=MAX(" & "K2" & ":" & "K" & summaryRowCount & ")"
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").Formula = "=MIN(" & "K2" & ":" & "K" & summaryRowCount & ")"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").Formula = "=MAX(" & "L2" & ":" & "L" & summaryRowCount & ")"
        ws.Range("Q4").NumberFormat = "General"
        
        'For loop through summary table to find tickers for players
        For j = 2 To summaryRowCount
        If ws.Cells(j, 11).Value = ws.Range("Q2").Value Then
            ws.Range("P2").Value = ws.Cells(j, 9).Value
        ElseIf ws.Cells(j, 11).Value = ws.Range("Q3").Value Then
            ws.Range("P3").Value = ws.Cells(j, 9).Value
        ElseIf CStr(ws.Cells(j, 12).Value) = CStr(ws.Range("Q4").Value) Then
            ws.Range("P4").Value = ws.Cells(j, 9).Value
        End If
        Next j
        
        ' Autofit to summary data
        ws.Columns("J:L").AutoFit
        ws.Columns("O:Q").AutoFit
     
        
    Next ws
     
    
End Sub

