Attribute VB_Name = "Module1"
    'General Comment: ws was passed as an argument in subs for the last sub
    'which itirates through all the sheets calling pervious subs

Sub stock(ws)

    'set variable to hold ticker name
Dim ticker As String

    'set variable to hold change yearly change value
Dim yearly_change As Double

    'set variable to hold percentage change value
Dim precentage_change As Double

    'set variable to hold volume value
Dim volume As Double

    'set initial volume value to 0
volume = 0

    'track location of each value in summary table
Dim summary_raw As Integer
summary_raw = 2

    'find the last raw in column A
last_raw = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'declare firs value for the open price for itiration
Dim open_price As Double
open_price = ws.Cells(2, 3).Value

    'create a headers for the summary table
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

    'itirate through all raws and populate summary table
For i = 2 To last_raw

    'add the voulme value of current ticker
    volume = volume + ws.Cells(i, 7).Value

    'check if the next raw is different ticker
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        'set current ticker name
        ticker = ws.Cells(i, 1).Value

        'set closing price for current ticker
        close_price = ws.Cells(i, 6).Value

        'find yearly change for current ticker
        yearly_change = close_price - open_price
        
        'find yearly change for current ticker
        precentage_change = ((close_price - open_price) / open_price)
        
        'populate summary table with found values
        ws.Range("I" & summary_raw).Value = ticker
        ws.Range("J" & summary_raw).Value = yearly_change
        ws.Range("K" & summary_raw).Value = precentage_change
        ws.Range("L" & summary_raw).Value = volume
        
        'move to next raw in summary table
        summary_raw = summary_raw + 1
        
        'reset vaolume for new ticker
        volume = 0
        
            'If not at the last row, set new open price for a new ticker
            If i <> last_raw Then

            open_price = ws.Cells(i + 1, 3).Value

            End If

    End If

Next i

End Sub

Sub format(ws)

    'find the last row with data in column J
Dim last_row2 As Long
last_raw2 = ws.Cells(Rows.Count, 10).End(xlUp).Row

    'fromat column k to show percentages
Range("K2:K" & last_raw2).NumberFormat = "0.00%"

    'itirate throug column J and apply conditinal formating
For j = 2 To last_raw2

    'if the yearly change is positive, make cell green
    If ws.Cells(j, 10).Value >= 0 Then
    
    ws.Cells(j, 10).Interior.ColorIndex = 4
    
    'if the yearly change is negative, make cell red
    ElseIf ws.Cells(j, 10).Value <= 0 Then
    
    ws.Cells(j, 10).Interior.ColorIndex = 3
    
    'if there is no change make cell yellow
    Else: ws.Cells(j, 10).Interior.ColorIndex = 6
    
    End If
    
Next j

End Sub

Sub functionality(ws)

    'set variables to track max increase/decrease and max volume
Dim max_increase As Double
Dim max_decrease As Double
Dim max_volume As Double

    'create headers for additional table
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % increase"
ws.Range("O3").Value = "Greatest % decrease"
ws.Range("O4").Value = "Greatest total volume"

    'find the last row with data in column K
last_raw3 = ws.Cells(Rows.Count, 11).End(xlUp).Row

    'set tracking virables as first values in summary table
max_increase = ws.Cells(2, 11).Value
max_decrease = ws.Cells(2, 11).Value
max_volume = ws.Cells(2, 12).Value

    'itirate through all raws in summary table and find
    'max increase/decrease and max volume
For k = 2 To last_raw3

    'start ittirating through summary table
    current_change = ws.Cells(k, 11)
    current_volume = ws.Cells(k, 12)

        'if current raw is greater rewritte max tracking variable
        If current_change >= max_increase Then
        max_increase = current_change
        
        'save the name of the ticker
        increase_ticker = ws.Cells(k, 9).Value
        
        End If
        
        'if current raw is lesser rewritte min racking variable
        If current_change <= max_decrease Then
        max_decrease = current_change
        
        'save the name of the ticker
        decrease_ticker = ws.Cells(k, 9).Value
        
        End If
        
        'if volume of the current raw bigger rewrite volume tracing variable
        If current_volume >= max_volume Then
        max_volume = current_volume
        
        'save the name of the ticker
        volume_ticker = ws.Cells(k, 9).Value
        
        End If
        
Next k

    'populate the new table with found tracking values
ws.Range("Q2").Value = max_increase
ws.Range("P2").Value = increase_ticker
ws.Range("Q3").Value = max_decrease
ws.Range("P3").Value = decrease_ticker
ws.Range("Q4").Value = max_volume
ws.Range("P4").Value = volume_ticker

    'adjust columns to automaticaly display data
ws.Columns("A:Q").AutoFit

    'show cells in percentage format
ws.Range("Q2:Q3").NumberFormat = "0.00%"

End Sub

Sub final_analysis()

    'declare variable to represent worksheets
Dim ws As Worksheet

    'itirate through all workseets calling subrutines
For Each ws In worksheets

    'call subrutine to create summary table
 stock ws
 
    'call subrutine to format summarty table
 format ws
 
    'call subrutine to create table for addtitional analysis
 functionality ws
 
 Next ws
 
    'dispaly message after compliting analysis across all worksheets
 MsgBox "Analysis Complete"

End Sub

