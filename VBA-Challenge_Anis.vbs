Attribute VB_Name = "Module1"
Sub VBAHomework()
    For Each ws In Worksheets
        ws.Activate
        Call SetTitle
    Next ws
    
End Sub

Sub SetTitle()
    Range("I:Q").Value = ""
    Range("I:Q").Interior.ColorIndex = 0
' Set title row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    'this is for challenge only
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("A:O").Columns.AutoFit
    
    Call CalculateSummary
    
End Sub

' Finding all the summaries of tickers through loops
Sub CalculateSummary()

    Dim ticker As String
    Dim ticker_vol As Double
    ticker_vol = 0
    
    Dim ticker_row As Integer
    ticker_row = 2
    
    Dim price_open As Long
    price_open = Cells(2, 3).Value
    
    Dim price_close As Double
    Dim change_year As Double
    Dim change_percent As Double
    
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim current_row As Long
    Dim current_col As Integer
    
    For current_row = 2 To lastrow
        If Cells(current_row + 1, 1).Value <> Cells(current_row, 1).Value Then
            ticker = Cells(current_row, 1).Value
            ticker_vol = ticker_vol + Cells(current_row, 7).Value
            Range("I" & ticker_row).Value = ticker
            Range("L" & ticker_row).Value = ticker_vol
            
            price_close = Cells(current_row, 6).Value
            change_year = (price_close - price_open)
            Range("J" & ticker_row).Value = change_year
            
            If price_open = 0 Then
                change_percent = 0
                
            Else
                change_percent = change_year / price_open
            
            End If
            
        Range("K" & ticker_row).Value = change_percent
        Range("K" & ticker_row).NumberFormat = "0.00%"
        
        ticker_row = ticker_row + 1
        
        ticker_vol = 0
        
        price_open = Cells(current_row + 1, 3)
        
    Else
        
        ticker_vol = ticker_vol + Cells(current_row, 7).Value
        
    End If
    
Next current_row

'Conditional Formatting
lastrow_summary = Cells(Rows.Count, 9).End(xlUp).Row

    For current_row = 2 To lastrow_summary
        
        If Cells(current_row, 10).Value > 0 Then
            Cells(current_row, 10).Interior.ColorIndex = 10
            
        
        Else
            Cells(current_row, 10).Interior.ColorIndex = 3
        
        End If
        
    Next current_row
        
'Extra Challege
For current_row = 2 To lastrow_summary

    If Cells(current_row, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_summary)) Then
        Cells(2, 16).Value = Cells(current_row, 9).Value
        Cells(2, 17).Value = Cells(current_row, 11).Value
        Cells(2, 17).NumberFormat = "0.00%"
        
    ElseIf Cells(current_row, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_summary)) Then
        Cells(3, 16).Value = Cells(current_row, 9).Value
        Cells(3, 17).Value = Cells(current_row, 11).Value
        Cells(3, 17).NumberFormat = "0.00%"
        
        ElseIf Cells(current_row, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_summary)) Then
            Cells(4, 16).Value = Cells(current_row, 9).Value
            Cells(4, 17).Value = Cells(current_row, 12).Value
        
        End If
    
    Next current_row
    
    Debug.Print ActiveSheet.Name
    
End Sub


