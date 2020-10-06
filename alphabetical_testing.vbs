Sub Ticker_Symbol()
    Dim ws As Worksheet
    
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
    
       
            Dim wsName As String
            Dim LastRow As LongLong
            Dim tsCount As LongLong
            Dim I As LongLong
            Dim J As LongLong
            Dim yearlyChange As Variant
            Dim Year_Open As Variant
            Dim Year_Close As Variant
            Dim TotalStock As LongLong
            
        
        
        
        
        
        'find the last row in each workbook with data on it
        'to populate Long variable LastRow
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'tsCount variable used in loop to count the number of
        'unique ticker symbols. Count starts at 1 because
        'the loop will begin on row 2. tsCount will be used
        'to index the unique values in column "I"
        
        tsCount = 1
        
        'title the cells with each field
        
        ws.Range("I1") = "Ticker Symbol"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        
        
        
        

        
        'loop through A2 to the last row for conditionals
        
        For I = 2 To LastRow
        
            'since these are listed in aplphabetical order the first different
            'ticker symbol in column a is the next unique value
            
            If ws.Cells(I, 1).Value <> ws.Cells(I - 1, 1).Value Then
                tsCount = tsCount + 1
                TotalStock = ws.Cells(I, 7).Value 'first value to add to Total Stock
                ws.Cells(tsCount, 9).Value = ws.Cells(I, 1).Value    'puts ticker symbol in
                Year_Open = ws.Cells(I, 3).Value 'takes the first value of the year
            
            
                        
            ElseIf ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1) Then
                TotalStock = TotalStock + ws.Cells(I, 7).Value 'adds last value in year for volume
                ws.Cells(tsCount, 12).Value = TotalStock 'places it in the table
                Year_Close = ws.Cells(I, 6).Value  'takes last value of last day of year
                yearlyChange = Year_Close - Year_Open 'finds yearly change
                'nested if statement to avoid div/0 error for yearly change %
                If yearlyChange = 0 Then
                    ws.Cells(tsCount, 10).Value = yearlyChange
                    ws.Cells(tsCount, 11).Value = 0
             
                    
                Else
                    ws.Cells(tsCount, 10).Value = yearlyChange
                    If Year_Open = 0 Then
                        ws.Cells(tsCount, 11) = 1
                    Else
                        ws.Cells(tsCount, 11) = (Year_Close - Year_Open) / Year_Open 'yearly change percent
                    End If
                    
                    
                End If
            
            'Total Stock Volume of the rows inbetween the first and last unique value
            'Can't figure out why it's working on most but a few are not coming up right :(
            
            ElseIf ws.Cells(I, 1).Value = ws.Cells(I - 1, 1) Then
                TotalStock = TotalStock + ws.Cells(I, 7).Value
            
            End If
            
            
        Next I
 
               
        
       
        'Find top increase decrease and volume
        Dim inc As Range
        Dim vol As Range
    
        Dim topInc As Double
        Dim topDec As Double
        Dim topVol As LongLong
    
        Set inc = Range("K:K")
        Set vol = Range("L:L")
        
        topInc = Application.WorksheetFunction.Max(inc)
        topDec = Application.WorksheetFunction.Min(inc)
        topVol = Application.WorksheetFunction.Max(vol)
        topIncTS = Application.WorksheetFunction.Index(Range("I:I"), Application.WorksheetFunction.Match(topInc, Range("K:K"), 0))
        topDecTS = Application.WorksheetFunction.Index(Range("I:I"), Application.WorksheetFunction.Match(topDec, Range("K:K"), 0))
        topVolTS = Application.WorksheetFunction.Index(Range("I:I"), Application.WorksheetFunction.Match(topVol, Range("L:L"), 0))
    
        ws.Range("O3").Value = "Greatest % Increase"
        ws.Range("O4").Value = "Greatest % Decrease"
        ws.Range("O5").Value = "Greatest Total Volume"
        ws.Range("P3").Value = topIncTS
        ws.Range("P4").Value = topDecTS
        ws.Range("P5").Value = topVolTS
        ws.Range("Q3").Value = topInc
        ws.Range("Q4").Value = topDec
        ws.Range("Q5").Value = topVol
        
        
        'Format greatest values
        
        If ws.Range("Q3").Value >= 10 Then
            ws.Range("Q3").NumberFormat = "0000.00%"
        ElseIf ws.Range("Q3").Value >= 1 Then
            ws.Range("Q3").NumberFormat = "000.00%"
        End If
    
        If ws.Range("Q4").Value <= -1 Then
            ws.Range("Q4").NumberFormat = "000.00%"
        Else
            ws.Range("Q4").NumberFormat = "00.00%"
        End If
        
    
        
    
     yearChangeRows = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        For J = 2 To yearChangeRows
            'format the percentage in column K
            If ws.Cells(J, 11).Value >= 1 Or ws.Cells(J, 11).Value <= -1 Then
                ws.Cells(J, 11).NumberFormat = "000.00%"
            ElseIf ws.Cells(J, 11).Value < 1 And ws.Cells(J, 11).Value >= 0.1 Then
                ws.Cells(J, 11).NumberFormat = "00.00%"
            ElseIf ws.Cells(J, 11).Value > -1 And ws.Cells(J, 11).Value <= -0.1 Then
                ws.Cells(J, 11).NumberFormat = "00.00%"
            Else
                ws.Cells(J, 11).NumberFormat = "0.00%"
            End If
            'format the color in column J
            If ws.Cells(J, 10).Value > 0 Then
                ws.Cells(J, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(J, 10).Value < 0 Then
                ws.Cells(J, 10).Interior.ColorIndex = 3
            ElseIf ws.Cells(J, 10).Value = 0 Then
                ws.Cells(J, 10).Interior.ColorIndex = 0
            End If
            
            
            
        Next J
        ws.Range("I:Q").EntireColumn.AutoFit
      
    Next ws
    

   
End Sub
