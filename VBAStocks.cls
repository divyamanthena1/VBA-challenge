VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockData():

        For Each ws In Worksheets
    
        
        Dim Total As Variant
        Dim CurrentTicker As String
        Dim SummaryRow As Integer
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        SummaryRow = 2
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        OpenPrice = ws.Cells(2, 3).Value
        
        ' Loop through tickers
        For Row = 2 To LastRow
        
            ' Add current value to total
            Total = Total + ws.Cells(Row, 7).Value
            
            CurrentTicker = ws.Cells(Row, 1).Value
            
            'if next ticker row is not still in current ticker then
            
            If CurrentTicker <> ws.Cells(Row + 1, 1).Value Then
                
                ClosePrice = ws.Cells(Row, 6).Value
                
                ' print total and ticker
                ws.Cells(SummaryRow, 9).Value = CurrentTicker
                ws.Cells(SummaryRow, 12).Value = Total
                ws.Cells(SummaryRow, 10).Value = ClosePrice - OpenPrice
                    If ClosePrice - OpenPrice <= 0 Then
                        ws.Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0)
                    Else
                        ws.Cells(SummaryRow, 10).Interior.Color = RGB(0, 255, 0)
                   End If
                
                    If OpenPrice = 0 Then
                    ws.Cells(SummaryRow, 11).Value = 0
                    Else
                    ws.Cells(SummaryRow, 11).Value = (ClosePrice - OpenPrice) / (OpenPrice)
                    End If
                 
                SummaryRow = SummaryRow + 1
                
                ' reset total to next value
                Total = 0
                
                ' reset ticker to next ticker
                CurrentTicker = ws.Cells(Row + 1, 1).Value
                OpenPrice = ws.Cells(Row + 1, 3).Value
             
        End If
    Next Row
    Next ws
End Sub


