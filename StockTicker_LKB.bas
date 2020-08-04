Attribute VB_Name = "Module1"
Sub StockTicker()

Dim Ticker As String
Dim TickerCounter As Integer
Dim YearOpen As Double
Dim YearClose As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim LastRow As Long
Dim i As Long
Dim Vol As Double
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestTotalVolume As Double
Dim GreatestIncreaseTicker As String
Dim GreatestDecreaseTicker As String
Dim GreatestTotalVolumeTicker As String
Dim OpenDate As Double

For Each ws In Worksheets

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Vol = 0
TickerCounter = 2
GreatestIncrease = 0
GreatestDecrease = 0
GreatestTotalVolume = 0
OpenDate = ws.Range("B2").Value

For i = 2 To LastRow

    Ticker = ws.Cells(i, 1).Value
    If ws.Cells(i, 2).Value = OpenDate Then
        YearOpen = ws.Cells(i, 3).Value
    End If

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        

        YearClose = ws.Cells(i, 6).Value
        If YearClose = 0 Then
            MsgBox ("Please note that the closing value for the following stock is zero (0), indicating that the data may be incomplete for the year " & ws.Name & ": " & ws.Cells(i, 1).Value & ".")
        End If
        
        YearlyChange = (YearClose - YearOpen)
        If YearOpen <> 0 Then
            PercentChange = (YearClose / YearOpen) - 1
        Else
            PercentChange = 1
            MsgBox ("Please note that the opening value for the following stock is zero (0), indicating the data may be incomplete for the year " & ws.Name & ": " & ws.Cells(i, 1).Value & ".")
        End If
        Vol = Vol + (ws.Cells(i, 7).Value)
        
        If PercentChange > GreatestIncrease Then
            GreatestIncrease = PercentChange
            GreatestIncreaseTicker = ws.Cells(i, 1).Value
        End If
        
        If PercentChange < GreatestDecrease Then
            GreatestDecrease = PercentChange
            GreatestDecreaseTicker = ws.Cells(i, 1).Value
        End If
        
        If Vol > GreatestTotalVolume Then
            GreatestTotalVolume = Vol
            GreatestTotalVolumeTicker = ws.Cells(i, 1).Value
        End If
                
        ws.Range("I" & TickerCounter).Value = Ticker
        ws.Range("J" & TickerCounter).Value = YearlyChange
        ws.Range("K" & TickerCounter).Value = Format(PercentChange, "Percent")
        ws.Range("L" & TickerCounter).Value = Vol
        
            If ws.Range("J" & TickerCounter).Value > 0 Then
            
                ws.Range("J" & TickerCounter).Interior.ColorIndex = 4
            
            Else
            
                ws.Range("J" & TickerCounter).Interior.ColorIndex = 3
            
            End If
            
            If ws.Range("K" & TickerCounter).Value > 0 Then
            
                ws.Range("K" & TickerCounter).Interior.ColorIndex = 4
            
            Else
            
                ws.Range("K" & TickerCounter).Interior.ColorIndex = 3
            
            End If
        
        Vol = 0
        TickerCounter = TickerCounter + 1
        
    Else
    
    Vol = Vol + (ws.Cells(i, 7).Value)
    
    End If

Next i
       
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percentage Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
ws.Range("O1").Value = "Ticker"
ws.Range("O2").Value = GreatestIncreaseTicker
ws.Range("O3").Value = GreatestDecreaseTicker
ws.Range("O4").Value = GreatestTotalVolumeTicker
ws.Range("P1").Value = "Value"
ws.Range("P2").Value = Format(GreatestIncrease, "Percent")
ws.Range("P3").Value = Format(GreatestDecrease, "Percent")
ws.Range("P4").Value = GreatestTotalVolume

ws.Columns("I:P").AutoFit

Next ws

End Sub

