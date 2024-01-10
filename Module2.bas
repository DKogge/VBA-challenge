Attribute VB_Name = "Module1"
Sub module2():

'definitions
    Dim summary_row As Integer
    Dim ticker As String
    Dim Jan As Double
    Dim Dec As Double
    Dim greatinc As Double
    Dim greatdec As Double
    Dim greatvol As Double
    
'worksheet loop
For Each ws In Worksheets

'lastrows
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    otherlastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
'baselines
    summary_row = 2
    Totsvol = 0
    Jan = ws.Cells(2, 3).Value
    Dec = 0
    
'set up summary table 1
    ws.Range("J1") = "Ticker"
    ws.Range("K1") = "Yearly Change"
    ws.Range("L1") = "Percent Change"
    ws.Range("M1") = "Total Stock Volume"
        
'setup summary table 2
    ws.Range("P2") = "Greatest % Increase"
    ws.Range("P3") = "Greatest % Decrease"
    ws.Range("P4") = "Greatest Total Volume"
    ws.Range("Q1") = "Ticker"
    ws.Range("R1") = "Value"

    
'regular loop time
    For i = 2 To lastrow
    
        'what to do when the ticker name changes
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
            'assign ticker name and input in summary
            ticker = ws.Cells(i, 1).Value
            ws.Range("j" & summary_row).Value = ticker
            
            'open/close yearly price change input
            Dec = ws.Cells(i, 6).Value
            ws.Range("k" & summary_row).Value = (Dec - Jan)
            
            'percentage change input
            ws.Range("l" & summary_row).Value = (Dec - Jan) / Jan
                
            'add last volume number and input in summary
            Totsvol = Totsvol + ws.Cells(i, 7).Value
            ws.Range("m" & summary_row).Value = Totsvol
              
            'reset for next loop
            summary_row = summary_row + 1
            Totsvol = 0
            Jan = ws.Cells(i + 1, 3).Value
            Dec = 0
            
        'what to do when the ticker stays the same
        Else
            Totsvol = Totsvol + ws.Cells(i, 7).Value
              
        End If
      
    Next i
    
'formatting loop. Sometimes this part doesn't run the first time, but will always show up
'if you run the macro again and I have no idea why

    For c = 2 To otherlastrow
        
        'red if value is less than 0
        If ws.Cells(c, 11).Value < 0 Then
        ws.Cells(c, 11).Interior.ColorIndex = 3
        
        'otherwise green
        Else: ws.Cells(c, 11).Interior.ColorIndex = 4
                
        End If
        
    Next c
    
'greatest inc/dec/vol using min/max function

    Set percentrange = ws.Range("L2:L3001")
    greatinc = Application.WorksheetFunction.Max(percentrange)
    ws.Range("R2").Value = greatinc
    
    greatdec = Application.WorksheetFunction.Min(percentrange)
    ws.Range("R3").Value = greatdec
    
    Set volrange = ws.Range("M2:M3001")
    greatvol = Application.WorksheetFunction.Max(volrange)
    ws.Range("R4").Value = greatvol
    
'loop to get the corresponding tickers - this also will often only run on the 2nd macro run. Why?
    For x = 2 To otherlastrow
    
        If ws.Cells(x, 12) = greatinc Then
        ws.Range("Q2") = ws.Cells(x, 10).Value
        
        ElseIf ws.Cells(x, 12) = greatdec Then
        ws.Range("Q3") = ws.Cells(x, 10).Value
        
        ElseIf ws.Cells(x, 13) = greatvol Then
        ws.Range("Q4") = ws.Cells(x, 10).Value
        
        End If
        
    Next x
    
Next ws

End Sub
