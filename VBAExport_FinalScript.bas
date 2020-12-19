Attribute VB_Name = "FinalScript"
Sub TestOnDecember17Test():

'Assign ws as common standard variable with global variable worksheet to analyze all sheets in workbook
    Dim ws As Worksheet
    
'Apply Loop so that script runs on every worksheet
    For Each ws In Worksheets

'Define ws variable for each column output on every worksheet
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

'Declare variables

    Dim ticker_name As String
    Dim open_price As Double
    Dim total_tickervolume As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim prior_total As Long
    Dim percent_change As Double
    Dim grt_percent_incr As Double
    Dim grt_percent_decr As Double
    Dim lastrow As Long
    Dim lastrowvalue As Long
    Dim sum_table As Long
    Dim grt_totalvolume As Double
    Dim i As Long
    
'Set variables as is appropriate for proper index placement &/or calculations
    total_tickervolume = 0
    prior_total = 2
    grt_percent_incr = 0
    grt_percent_decr = 0
    grt_totalvolume = 0
    open_price = 0
    close_price = 0
    sum_table = 2

'Set value for variable lastrow to find the last non-blank cell in column A
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
'Apply Loop where i represents the counter of iterations starting at row 2 to the last non-blank cell in column A (parameter set by variable lastrow)
    
    For i = 2 To lastrow
    ticker_name = ws.Cells(i, 1).Value
    
'Redefine variable total_tickervolume to include the values in column G (<vol>) to ticker total volume for summary table

    total_tickervolume = total_tickervolume + ws.Cells(i, 7).Value
    
'Apply an If...Then statement within the For...Next Loop to differentiate from values in cells a rowup for i in column A that do no equal values set by ticker_name variable

     If ws.Cells(i + 1, 1).Value <> ticker_name Then
    
'Populate the summary table with the ticker name in the summary table at column I and with value out at row 2 of summary table(recall '&', ampersand will concatenate values)

    ws.Range("I" & sum_table).Value = ticker_name
    
'Assign the total ticker volume to the summary table at column L and row 2 of summary table

    ws.Range("L" & sum_table).Value = total_tickervolume
    
'Reassign the variable for ticker volume

    total_tickervolume = 0

'Set variables for open_price, values at column C concatenate with variable prior_total set at 0 and close_price, values at column F concatenate with variable i

    open_price = ws.Range("C" & prior_total)
    close_price = ws.Range("F" & i)
    
'Set variable for yearly_change to calculation

    yearly_change = close_price - open_price

'Assign the output for yearly change to the summary table at column J and concatenate with variable sum_table set to row 2
    
    ws.Range("J" & sum_table).Value = yearly_change
    ws.Range("J" & sum_table).NumberFormat = "$0.00"

'Determine the percent change using an If...Else statement

    If open_price = 0 Then
        percent_change = 0
    Else
        open_price = ws.Range("C" & prior_total)
        percent_change = yearly_change / open_price
    End If
    
'Assign percent change output in column K at row 2 with variable sum_table

     ws.Range("K" & sum_table).Value = percent_change
    
'Apply conditional formatting highlights using an If...Then statement for yearly change in column J, row 2 where positive change is in green and negative change is in red

    If ws.Range("J" & sum_table).Value >= 0 Then
        ws.Range("J" & sum_table).Interior.ColorIndex = 4
    Else
        ws.Range("J" & sum_table).Interior.ColorIndex = 3
    End If
    
'Format specified data type to include the percent sign

    ws.Range("K" & sum_table).NumberFormat = "0.00%"

'Increase variable sum_table row by 1 and prior_total to include i + 1

    sum_table = sum_table + 1
    prior_total = i + 1
    
    End If
    
    Next i


'************************************BONUS******************************************************************'


'Set value for variable lastrowvalue to find the last non-blank cell in column L
    
    lastrowvalue = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
'Apply Loop where i represents the counter of iterations starting at row 2 to the last non-blank cell in column  L (parameter set by variable lastrowvalue)
    
    For i = 2 To lastrowvalue

'Apply an If...Then statement to populate bonus summary tables for greatest % increase

     If ws.Cells(i, 17).Value > ws.Cells(2, 17) Then
        ws.Cells(2, 17).Value = ws.Cells(i, 11)
        ws.Cells(2, 16).Value = ws.Cells(i, 9)
        
    End If
    
'Apply an If...Then statement to populate bonus summary table for greatest % decrease

    If ws.Cells(i, 11).Value < ws.Cells(3, 17).Value Then
        ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        
    End If
    
'Apply an If...Then statement to populate bonus summary table for greatest total volume

    If ws.Cells(i, 12).Value > ws.Cells(4, 17).Value Then
        ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        
    End If
    
    Next i
    
'Apply format change to include percentage within two decimal places for value output at column Q

    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"

    Next ws

End Sub

