Attribute VB_Name = "Module1"
Sub test():

Dim last_row As Long
Dim ticker As String
Dim yearly_change As Double
Dim stock_vol As Double
Dim summary_row As Integer
Dim year_open As Double
Dim year_close As Double
Dim max As Double
Dim min As Double
Dim vol_max As Double
Dim ticker_inc As String
Dim ticker_dec As String
Dim ticker_vol As String
Dim percent_change As Double
    
    percent_change = 0
    year_open = Cells(2, 3).Value
    year_close = 0
    summary_row = 2
    yearly_change = 0
    stock_vol = 0
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    max = 0
    min = 0
    vol_max = 0
    
Range("J1").Value = "Ticker"
Range("K1").Value = "Yearly Change"
Range("L1").Value = "Percent Change"
Range("M1").Value = " Total Stock Volume"

Range("P2").Value = "Greatest % Increase"
Range("P3").Value = "Greatest % Decrease"
Range("P4").Value = "Greatest Total Volume"

Range("Q1").Value = "Ticker"
Range("R1").Value = "Value"

For i = 2 To last_row

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
    ticker = Cells(i, 1).Value
    year_close = Cells(i, 6).Value
        
        yearly_change = year_close - year_open
        
            If year_open = 0 Then
            
                yearly_change = 0
                percent_change = 0
            
            Else
            
                percent_change = (yearly_change / year_open)
            
            End If
            
        stock_vol = stock_vol + Cells(i, 7).Value
        
        Range("J" & summary_row).Value = ticker
        Range("K" & summary_row).Value = yearly_change
        Range("L" & summary_row).Value = FormatPercent(percent_change)
        Range("M" & summary_row).Value = stock_vol
            
            If Range("K" & summary_row).Value > 0 Then
                Range("K" & summary_row).Interior.ColorIndex = 4
            
            Else
                Range("K" & summary_row).Interior.ColorIndex = 3
                
            End If
        
        summary_row = summary_row + 1
        
        year_change = 0
        year_open = Cells(i + 1, 3).Value
        year_close = 0
        stock_vol = 0
    
    Else
        
        stock_vol = stock_vol + Cells(i, 7).Value
        
    End If
Next i

    For i = 2 To last_row
        If Cells(i, 12) > max Then
        
            max = Cells(i, 12).Value
            ticker_inc = Cells(i, 10).Value
            
        ElseIf Cells(i, 12).Value < min Then
            
            min = Cells(i, 12).Value
            ticker_dec = Cells(i, 10).Value
        
        End If
    Next i
    
    For i = 2 To last_row
        If Cells(i, 13).Value > vol_max Then
        
            vol_max = Cells(i, 13).Value
            ticker_vol = Cells(i, 10).Value
        End If
    Next i
    

    Range("Q2").Value = ticker_inc
    Range("R2").Value = max
    Range("R3").Value = FormatPercent(min)
    Range("Q3").Value = ticker_dec
    Range("Q4").Value = ticker_vol
    Range("R4").Value = vol_max
    
End Sub



