Attribute VB_Name = "Module1"

Option Explicit

Sub Stock_Data()
    Dim ws As Worksheet
    Dim i As Long, TickerRow As Long, lastrow As Long, lastrow_summary_table As Long
    Dim ticker As String
    Dim open_price As Double, close_price As Double, quarterly_change As Double, percent_change As Double
    Dim stockvolume As Double
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        
        ' Set headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Quarterly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        TickerRow = 2
        stockvolume = 0
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        open_price = Cells(2, 3).Value ' Initialize open price
        
        For i = 2 To lastrow
            ' Check if the ticker changes
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ticker = Cells(i, 1).Value
                close_price = Cells(i, 6).Value
                stockvolume = stockvolume + Cells(i, 7).Value
                
                ' Calculate changes
                quarterly_change = close_price - open_price
                If open_price <> 0 Then
                    percent_change = quarterly_change / open_price
                Else
                    percent_change = 0
                End If
                
                ' Output results
                Range("I" & TickerRow).Value = ticker
                Range("J" & TickerRow).Value = quarterly_change
                Range("K" & TickerRow).Value = percent_change
                Range("K" & TickerRow).NumberFormat = "0.00%"
                Range("L" & TickerRow).Value = stockvolume
                
                ' Reset for the next ticker
                TickerRow = TickerRow + 1
                stockvolume = 0
                If i + 1 <= lastrow Then
                    open_price = Cells(i + 1, 3).Value
                End If
            Else
                stockvolume = stockvolume + Cells(i, 7).Value
            End If
        Next i
        
        ' Highlight positive and negative quarterly changes
        lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
        For i = 2 To lastrow_summary_table
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.Color = RGB(0, 255, 0) ' Green
            Else
                Cells(i, 10).Interior.Color = RGB(255, 0, 0) ' Red
            End If
        Next i
        
        ' Find greatest % increase, % decrease, and total volume
        With Application.WorksheetFunction
            Cells(2, 16).Value = Cells(.Match(.Max(Range("K2:K" & lastrow_summary_table)), Range("K2:K" & lastrow_summary_table), 0) + 1, 9).Value
            Cells(2, 17).Value = .Max(Range("K2:K" & lastrow_summary_table))
            Cells(2, 17).NumberFormat = "0.00%"
            
            Cells(3, 16).Value = Cells(.Match(.Min(Range("K2:K" & lastrow_summary_table)), Range("K2:K" & lastrow_summary_table), 0) + 1, 9).Value
            Cells(3, 17).Value = .Min(Range("K2:K" & lastrow_summary_table))
            Cells(3, 17).NumberFormat = "0.00%"
            
            Cells(4, 16).Value = Cells(.Match(.Max(Range("L2:L" & lastrow_summary_table)), Range("L2:L" & lastrow_summary_table), 0) + 1, 9).Value
            Cells(4, 17).Value = .Max(Range("L2:L" & lastrow_summary_table))
        End With
    Next ws
End Sub

