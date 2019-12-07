Sub Stock_Hard()

For Each ws In Worksheets

Dim ticker As String
Dim LastRow As Long
Dim ticker_vol As Double
ticker_vol = 0
Dim Summary_Table_Row As Long
Summary_Table_Row = 2
Dim openVolume, closeVolume, yearly_change, percentage_change As Double
Dim PreviousAmount As Long
PreviousAmount = 2
Dim GreatestIncrease As Double
GreatestIncrease = 0
Dim GreatestDecrease As Double
GreatestDecrease = 0
Dim GreatestTotalVolume As Double
GreatestTotalVolume = 0

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
For i = 2 To LastRow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ticker_vol = ticker_vol + ws.Cells(i, 7).Value
      
        ticker = ws.Cells(i, 1).Value

        ws.Range("I" & Summary_Table_Row).Value = ticker

        ws.Range("L" & Summary_Table_Row).Value = ticker_vol
        ticker_vol = 0

        openVolume = ws.Range("C" & PreviousAmount)

        closeVolume = ws.Range("F" & i)

        yearly_change = (closeVolume - openVolume)

        ws.Range("J" & Summary_Table_Row).Value = yearly_change

            If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4

            Else
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3

            End If

        
        If openVolume = 0 Then
            
            percentage_change = 0
        
        Else
            openVolume = ws.Range("C" & PreviousAmount)
            percentage_change = yearly_change / openVolume
            
        End If
        
        ws.Range("K" & Summary_Table_Row).Value = percentage_change
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
     

        ws.Range("I1").Value = "Ticker value"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"

        Summary_Table_Row = Summary_Table_Row + 1

        PreviousAmount = i + 1

        End If
 
Next i

    For i = 2 To LastRow
        If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
            ws.Range("Q2").Value = ws.Range("K" & i).Value
            ws.Range("P2").Value = ws.Range("I" & i).Value
        End If

        If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
            ws.Range("Q3").Value = ws.Range("K" & i).Value
            ws.Range("P3").Value = ws.Range("I" & i).Value
        End If

        If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
            ws.Range("Q4").Value = ws.Range("L" & i).Value
            ws.Range("P4").Value = ws.Range("I" & i).Value
        End If

    Next i
    
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").NumberFormat = "0.00%"
            
ws.Columns("I:Q").AutoFit

Next ws

End Sub
