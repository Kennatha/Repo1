
Sub stocks()
For Each ws In Worksheets
Dim lastrow As Long

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim volume As Double

volume = 0
Dim tablerow As Integer
tablerow = 2
Dim ticker As String
Dim Per_change As Double
Dim year_change As Double
Dim open_price As Double
Dim close_price As Double
Dim stockcount As Integer
stockcount = 0
Dim output As Double
Dim output2 As Double
Dim output3 As Double
Dim max_increase As String
Dim max_decrease As String
Dim max_volume As String

output = 0
output2 = 0
output3 = 0
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "% Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

For i = 2 To lastrow

open_price = ws.Cells(i - stockcount, 3).Value
close_price = ws.Cells(i, 6).Value
    If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
    ticker = ws.Cells(i, 1).Value
    ws.Range("I" & tablerow).Value = ticker
    volume = volume + ws.Cells(i, 7).Value
    ws.Range("L" & tablerow).Value = volume
   
    year_change = open_price - close_price
    Per_change = year_change / open_price
    ws.Range("j" & tablerow).Value = year_change
    ws.Range("K" & tablerow).Value = Per_change
        If ws.Range("j" & tablerow).Value <= 0 Then
        ws.Range("j" & tablerow).Interior.ColorIndex = 3
        ws.Range("k" & tablerow).Interior.ColorIndex = 3
       Else:
       ws.Range("j" & tablerow).Interior.ColorIndex = 4
        ws.Range("k" & tablerow).Interior.ColorIndex = 4
        End If
        
        
     tablerow = tablerow + 1
    volume = 0
    stockcount = 0

    Else: volume = volume + ws.Cells(i, 7).Value
    stockcount = stockcount + 1

    
    End If
    If ws.Cells(i, 11).Value > output Then
            output = ws.Cells(i, 11).Value
            max_increase = ws.Cells(i, 9).Value
            ws.Range("r2").Value = output
            ws.Range("s2").Value = max_increase
            
            Else: ws.Range("r2").Value = output
            ws.Range("s2").Value = max_increase
            
            End If
            
            If ws.Cells(i, 11).Value < output2 Then
            output2 = ws.Cells(i, 11).Value
            max_decrease = ws.Cells(i, 9).Value
            ws.Range("r3").Value = output2
            ws.Range("s2").Value = max_decrease
            
            Else: ws.Range("r3").Value = output2
            ws.Range("s3").Value = max_decrease
            End If
            If ws.Cells(i, 12).Value > output3 Then
            output3 = ws.Cells(i, 12).Value
            max_volume = ws.Cells(i, 9).Value
            ws.Range("r4").Value = output3
            ws.Range("s4").Value = max_volume
            
            Else: ws.Range("r4").Value = output3
            ws.Range("s4").Value = max_volume
            End If
    Next i
    
    ws.Range("r1").Value = "Ticker"
    ws.Range("s1").Value = "Value"
    ws.Range("q2").Value = "Greatest % Increase"
    ws.Range("q3").Value = "Greatest % Decrease"
    ws.Range("q4").Value = "Greatest Total Volume"
      
       
    Next ws
    
End Sub

