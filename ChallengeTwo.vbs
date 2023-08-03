Attribute VB_Name = "Module1"
Sub Stocks()
    ' Create variables
    Dim ticker As String
    Dim TSV As Double
    Dim i, j, a, b, c, oprow As Integer
    Dim OpeningP, ClosingP, YC, PC As Double
    
    'Set the initial value for the Total stock volume to 0
    TSV = 0
    
    'Counts the number of rows
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Inserting data via Ranges
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    'Set the row number of initial opening price
    oprow = 2
    
    'Set the initial opening price
    OpeningP = Cells(2, 3).Value
    
    ' Loop through rows in the column
    For i = 2 To lastrow
    
        ' Searches for when the value of the next cell is different than that of the current cell
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
           Cells(oprow, 9).Value = Cells(i, 1).Value
           
           'Set the closing price
           ClosingP = Cells(i, 6).Value
           
           'Set the yearly change
           YC = ClosingP - OpeningP
           Cells(oprow, 10).Value = YC
           
           'Set the percentage change
           PC = YC / OpeningP
           Cells(oprow, 11).Value = PC
           
           'Change the format of cells to percentage
           Range("K" & oprow).NumberFormat = "0.00%"
           
           'Set opening price
           OpeningP = Cells(i + 1, 3).Value
           
           ' Set the cell colours to green when it is positive
           If Cells(oprow, 10).Value > 0 Then
             Cells(oprow, 10).Interior.ColorIndex = 4
             
             ' Set the cell colours to red when it is negative
             ElseIf Cells(oprow, 10).Value < 0 Then
             Cells(oprow, 10).Interior.ColorIndex = 3
             
             
            End If
            
            ' Set the cell colours to green when it is positive
            If Cells(oprow, 11).Value > 0 Then
             Cells(oprow, 11).Interior.ColorIndex = 4
             
             ' Set the cell colours to red when it is negative
             ElseIf Cells(oprow, 11).Value < 0 Then
             Cells(oprow, 11).Interior.ColorIndex = 3
            End If
                
           'Set total stock volume
           TSV = TSV + Cells(i, 7).Value
           Range("L" & oprow).Value = TSV
           
           'Reset the total stock volume
           TSV = 0
           
           'Add one to the opening price row
           oprow = oprow + 1
           
           
           'When the ticker is same
           Else
           'Add to total stock volume
           TSV = TSV + Cells(i, 7).Value
           
        End If
        
        
    Next i

    'Bonus Questions
    
    
    'Inserting data via Ranges
    Range("O2").Value = "Greatest % increase"
    Range("O3").Value = "Greatest % decrease"
    Range("O4").Value = "Greatest total volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    'Use a worksheet function to find the greatest increase %
    Range("Q2") = WorksheetFunction.Max(Range("K:K"))
    
    'Use a worksheet function to find the greatest decrease %
    Range("Q3") = WorksheetFunction.Min(Range("K:K"))
    
    'Use a worksheet function to find the greatest total stock volume
    Range("Q4") = WorksheetFunction.Max(Range("L:L"))
    
    'Change the format of cells to percentage
    Range("Q2:Q3").NumberFormat = "0.00%"
    
    'Use a worksheet function to find the row of the greatest increase %
    a = WorksheetFunction.Match(Range("Q2"), (Range("K:K")), 0)
    'Locate the ticker name via row of the greatest increase and ticker colume
    Range("P2") = Cells(a, 9)
    
    'Use a worksheet function to find the row of the greatest decrease %
    b = WorksheetFunction.Match(Range("Q3"), (Range("K:K")), 0)
    'Locate the ticker name via row of the greatest decrease and ticker colume
    Range("P3") = Cells(b, 9)
    
    'Use a worksheet function to find the row of the greatest total stock volume
    c = WorksheetFunction.Match(Range("Q4"), (Range("L:L")), 0)
    'Locate the ticker name via row of the greatest total stock volume and ticker colume
    Range("P4") = Cells(c, 9)
    
End Sub

