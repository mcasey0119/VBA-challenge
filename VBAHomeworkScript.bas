Attribute VB_Name = "Module1"
Sub vbaHomework()
    Dim ticker As String
    Dim open_price As Double
    Dim close_price As Double
    Dim pct_change As Integer
    Dim summary_table As Integer
    Dim last_row As Long
    Dim yearly_change As Double
    Dim ticker_count As Integer
    
    'goes to end of spreadsheet
    
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    'define ticker_count and summary_table start points
    
    ticker_count = 0
    summary_table = 2
    
    ' loop through all rows
    
    For i = 1 To last_row
        ' loop through close_price column 6
        close_price = Cells(i, 6).Value
        ' loop through ticker column 1
        ticker = Cells(i, 1).Value
        
        If ticker <> Cells(i + 1, 1).Value Then
            open_price = Cells(i + 1, 3).Value
            yearly_change = close_price - open_price
            
            'defining when open_price is not equal to 0
            If open_price <> 0 Then
                pct_change = yearly_change / open_price * 100
            Else
                pct_change = 100
            End If
            
            'got this help from instructor
            Range("I" & summary_table).Value = ticker
            Range("J" & summary_table).Value = yearly_change
            Range("K" & summary_table).Value = pct_change
            summary_table = summary_table + 1
            
        End If
    Next i
            

    
End Sub
