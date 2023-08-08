Attribute VB_Name = "Module1"
Sub Stocks()
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
    ' Create variables
    Dim ticker As String
    Dim TSV As Double
    Dim i, j, a, b, c, oprow As Integer
    Dim OpeningP, ClosingP, YC, PC As Double
    
    'Set the initial value for the Total stock volume to 0
    TSV = 0
    
    'Counts the number of rows
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Inserting data via Ranges
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Set the row number of initial opening price
    oprow = 2
    
    'Set the initial opening price
    OpeningP = ws.Cells(2, 3).Value
    
    ' Loop through rows in the column
    For i = 2 To lastrow
    
        ' Searches for when the value of the next cell is different than that of the current cell
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
           ws.Cells(oprow, 9).Value = ws.Cells(i, 1).Value
           
           'Set the closing price
           ClosingP = ws.Cells(i, 6).Value
           
           'Set the yearly change
           YC = ClosingP - OpeningP
           ws.Cells(oprow, 10).Value = YC
           
           'Set the percentage change
           PC = YC / OpeningP
           ws.Cells(oprow, 11).Value = PC
           
           'Change the format of cells to percentage
           ws.Range("K" & oprow).NumberFormat = "0.00%"
           
           'Set opening price
           OpeningP = ws.Cells(i + 1, 3).Value
           
           ' Set the cell colours to green when it is positive
           If ws.Cells(oprow, 10).Value > 0 Then
             ws.Cells(oprow, 10).Interior.ColorIndex = 4
             
             ' Set the cell colours to red when it is negative
             ElseIf ws.Cells(oprow, 10).Value < 0 Then
             ws.Cells(oprow, 10).Interior.ColorIndex = 3
             
             
            End If
            
            ' Set the cell colours to green when it is positive
            If ws.Cells(oprow, 11).Value > 0 Then
             ws.Cells(oprow, 11).Interior.ColorIndex = 4
             
             ' Set the cell colours to red when it is negative
             ElseIf ws.Cells(oprow, 11).Value < 0 Then
             ws.Cells(oprow, 11).Interior.ColorIndex = 3
            End If
                
           'Set total stock volume
           TSV = TSV + ws.Cells(i, 7).Value
           ws.Range("L" & oprow).Value = TSV
           
           'Reset the total stock volume
           TSV = 0
           
           'Add one to the opening price row
           oprow = oprow + 1
           
           
           'When the ticker is same
           Else
           'Add to total stock volume
           TSV = TSV + ws.Cells(i, 7).Value
           
        End If
        
        
    Next i

    'Bonus Questions
    
    
    'Inserting data via Ranges
    ws.Range("O2").Value = "Greatest % increase"
    ws.Range("O3").Value = "Greatest % decrease"
    ws.Range("O4").Value = "Greatest total volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'Use a worksheet function to find the greatest increase %
    ws.Range("Q2") = WorksheetFunction.Max(ws.Range("K:K"))
    
    'Use a worksheet function to find the greatest decrease %
    ws.Range("Q3") = WorksheetFunction.Min(ws.Range("K:K"))
    
    'Use a worksheet function to find the greatest total stock volume
    ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L:L"))
    
    'Change the format of cells to percentage
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    'Use a worksheet function to find the row of the greatest increase %
    a = WorksheetFunction.Match(ws.Range("Q2"), (ws.Range("K:K")), 0)
    'Locate the ticker name via row of the greatest increase and ticker colume
    ws.Range("P2") = ws.Cells(a, 9)
    
    'Use a worksheet function to find the row of the greatest decrease %
    b = WorksheetFunction.Match(ws.Range("Q3"), (ws.Range("K:K")), 0)
    'Locate the ticker name via row of the greatest decrease and ticker colume
    ws.Range("P3") = ws.Cells(b, 9)
    
    'Use a worksheet function to find the row of the greatest total stock volume
    c = WorksheetFunction.Match(ws.Range("Q4"), (ws.Range("L:L")), 0)
    'Locate the ticker name via row of the greatest total stock volume and ticker colume
    ws.Range("P4") = ws.Cells(c, 9)
Next ws
End Sub



