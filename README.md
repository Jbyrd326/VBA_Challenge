# VBA_Challenge

Sub multipleyearstockdata()

     For Each ws In Worksheets

        Dim Toatal_Volume As Integer
        ticker_row = 2
        Total = 0
      
       Dim Table As Integer
       Table = 2
      
       ws.Range("I1").Value = "Ticker"
       ws.Range("J1").Value = "Yearly Change"
       ws.Range("K1").Value = "Percent Change"
       ws.Range("L1").Value = "Total Stock Volume"
       
       lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       total_volume = 0
       
     'creating for loop to find stock names
       For i = 2 To lastrow:

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
               ticker = ws.Cells(i, 1).Value
               Year_Open = ws.Cells(i, 3).Value
               Year_Close = ws.Cells(i, 6).Value
               total_volume = total_volume + ws.Cells(i, 7).Value
               
               total_volume = 0
               
         Else
         
           total_volume = total_volume + ws.Cells(i, 7).Value
         
        
        If Year_Open = 0 Then
        
               Percent_Change = Year_Close
               
        Else
        
                  Yearly_Change = Year_Close - Year_Open
                    Percent_Change = Yearly_Change / Year_Open
               
               
            Else
            
              Total = Total + ws.Range("G" & i).Value
             
             End If
             
        Next i
    Next ws

End Sub


'_____________________________________________________________________

Dim K As Long
For K = 2 To lastrow
If Cells(K, 10).Value > 0 Then
Cells(K, 10).Interior.ColorIndex = 4
Else: Cells(K, 10).Interior.ColorIndex = 3
End If
