Option Explicit
Sub StockData()

    Const TICKER_COLUMN As Integer = 1
    Const OPEN_COLUMN As Integer = 3
    Const CLOSE_COLUMN As Integer = 6
    Const YEARLY_COLUMN As Integer = 10
    Const PERCENT_COLUMN As Integer = 11
    Const TICKER2_COLUMN As Integer = 16
    Const VALUE_COLUMN As Integer = 17
                  
    Dim ws As Worksheet
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As String
    Dim output_row As Integer
    Dim last_row As Long
    Dim input_row As Long
    Dim open_price As Double
    Dim close_price As Double
    Dim output2_row As Integer
    Dim percent_incr As Double
    Dim percent_decr As Double
    Dim total_value As String
        
    For Each ws In ThisWorkbook.Worksheets
    ws.Activate
           
        output_row = 2
        last_row = Cells(Rows.Count, TICKER_COLUMN).End(xlUp).Row
        'MsgBox (last_row)
        
        For input_row = 2 To last_row
            Ticker = Cells(input_row, TICKER_COLUMN).Value
    
            'First row of new ticker.
            If Cells(input_row - 1, TICKER_COLUMN).Value <> Ticker Then
                open_price = Cells(input_row, OPEN_COLUMN).Value
                Total_Stock_Volume = 0
            End If
    
            'Every row
            Total_Stock_Volume = Total_Stock_Volume + Cells(input_row, 7).Value
    
            'Last row of current ticker.
            If Cells(input_row + 1, TICKER_COLUMN).Value <> Ticker Then
                close_price = Cells(input_row, CLOSE_COLUMN).Value
                Yearly_Change = close_price - open_price
                Percent_Change = Yearly_Change / open_price
    
                'Output
                Range("I" & output_row).Value = Ticker
                Range("J" & output_row).Value = Yearly_Change
    
                If Cells(output_row, YEARLY_COLUMN).Value > 0 Then
                    Cells(output_row, YEARLY_COLUMN).Interior.Color = vbGreen
                    Else: Cells(output_row, YEARLY_COLUMN).Interior.Color = vbRed
                End If
    
                Range("K" & output_row).Value = Percent_Change
                Range("K" & output_row).NumberFormat = "0.00%"
                If Cells(output_row, PERCENT_COLUMN).Value > 0 Then
                    Cells(output_row, PERCENT_COLUMN).Interior.Color = vbGreen
                    Else: Cells(output_row, PERCENT_COLUMN).Interior.Color = vbRed
                End If
                Range("L" & output_row).Value = Total_Stock_Volume
                output_row = output_row + 1
    
            End If
    
        Next input_row
    
        output2_row = 2
        percent_incr = Application.WorksheetFunction.Max(Columns("K"))
        percent_decr = Application.WorksheetFunction.Min(Columns("K"))
        total_value = Application.WorksheetFunction.Max(Columns("L"))
        
        Range("Q" & output2_row).Value = percent_incr
        Range("Q" & output2_row).NumberFormat = "0.00%"
        Range("Q" & output2_row + 1).Value = percent_decr
        Range("Q" & output2_row + 1).NumberFormat = "0.00%"
        Range("Q" & output2_row + 2).Value = total_value
        
        Cells(2, 16).Value = Application.WorksheetFunction.Index(Range("I:I"), Application.WorksheetFunction.Match(Cells(2, 17).Value, Range("K:K"), 0), 1)
        Cells(3, 16).Value = Application.WorksheetFunction.Index(Range("I:I"), Application.WorksheetFunction.Match(Cells(3, 17).Value, Range("K:K"), 0), 1)
        Cells(4, 16).Value = Application.WorksheetFunction.Index(Range("I:I"), Application.WorksheetFunction.Match(Cells(4, 17).Value, Range("L:L"), 0), 1)
     
     Next ws
     
End Sub
