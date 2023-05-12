VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stockaroo()

For Each ws In Worksheets

    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    'Dim OpenPrice As Double

    Dim SummaryRow As Integer
    SummaryRow = 2

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "YearlyChange"
    ws.Cells(1, 12).Value = "PercentChange"
    ws.Cells(1, 13).Value = "Total Stock Volume"
    ws.Cells(1, 17).Value = "Ticker"
    ws.Cells(1, 18).Value = "Value"
    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(3, 16).Value = "Greatest % Decrease"
    ws.Cells(3, 16).Value = "Greatest Total Volume"

        
        
        
        For i = 2 To LastRow
        
            For j = 2 To 3
            
            
                If ws.Cells(i, 2).Value = 20180102 Then
                
                OpenPrice = ws.Cells(i, 3).Value
                
                ElseIf ws.Cells(i, 2).Value = 20190102 Then
                
                OpenPrice = ws.Cells(i, 3).Value
                
                ElseIf ws.Cells(i, 2).Value = 20200102 Then
                
                OpenPrice = ws.Cells(i, 3).Value
                
                End If
                

            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                YearlyChange = (ws.Cells(i, 6).Value - OpenPrice)
                PercentChange = (((ws.Cells(i, 6).Value - OpenPrice) / OpenPrice) * 100)
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

                ws.Range("J" & SummaryRow).Value = Ticker
                ws.Range("K" & SummaryRow).Value = YearlyChange
                ws.Range("L" & SummaryRow).Value = Format(PercentChange, "0.00%")
                ws.Range("M" & SummaryRow).Value = TotalStockVolume

                SummaryRow = SummaryRow + 1
                YearlyChange = 0
                PercentChange = 0
                TotalStockVolume = 0
                'OpenPrice = Empty

                'ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value

            Else
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            End If
                
            Next j
            
            'ws.Cells(i, 1).Value = ws.Cells(i, 1).Value
            
        Next i
        'OpenPrice = 0
        

    Next ws

End Sub

Sub VolumeReset()

For Each ws In Worksheets

    Dim Ticker As String
    Dim TotalStockVolume As Double
    Dim Volume As Double
    Dim SummaryRow As Integer
    SummaryRow = 2

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 13).Value = "Total Stock Volume"

        For i = 2 To LastRow

        'OpenPrice = ws.Cells(i, 3).Value
        Volume = ws.Cells(i, 7).Value

            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                
                
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

                ws.Range("J" & SummaryRow).Value = Ticker
                ws.Range("M" & SummaryRow).Value = TotalStockVolume

                SummaryRow = SummaryRow + 2
                TotalStockVolume = 0

            Else
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            End If
        Next i

    Next ws

End Sub

Sub stockconditions()


For Each ws In Worksheets

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(3, 16).Value = "Greatest % Decrease"
    ws.Cells(1, 17).Value = "Ticker"
    ws.Cells(1, 18).Value = "Value"

       For i = 2 To LastRow
        For j = 10 To 11
        
       
        If ws.Cells(i, 11).Value > 0 Then
            ws.Cells(i, 11).Interior.ColorIndex = 4
        
        ElseIf ws.Cells(i, 11).Value <= 0 Then
            ws.Cells(i, 11).Interior.ColorIndex = 3
        Else
            ws.Cells(i, 11).Interior.ColorIndex = 2
        End If
        
        Next j
        Next i

    Next ws

End Sub

