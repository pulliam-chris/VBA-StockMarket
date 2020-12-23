Attribute VB_Name = "Module2"
Sub StockAnalysis_Challenge()

' Loop through all sheets
    For Each ws In Worksheets
        
        'Instert Calculation Table Headers
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        ' Autofit to summary data
        ws.Columns("J:L").AutoFit
        ws.Columns("O:Q").AutoFit
     
        'MsgBox (ws.Name)
    Next ws
     
    
End Sub

