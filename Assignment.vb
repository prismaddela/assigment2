Sub StockFinalMod()

'set variable for specifying which column

Dim LastRow
Dim total As Double
Dim Sum As Integer
Dim Yearopen As Double
Dim Yearclose As Double
Dim Yearchange As Double
Dim Percentchange As Double


For Each ws In ThisWorkbook.Worksheets

 'set headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
 
 LastRow = Cells(Rows.Count, "A").End(xlUp).Row

Sum = 2
total = 0

'loop through column
For i = 2 To LastRow


    'search when next cell is different from previous
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
          
          'total volume in table
           Yearclose = Cells(i - 1, 6).Value
           Yearopen = Cells(i, 3).Value
                
            'year and percent change calcs
            Yearchange = Yearclose - Yearopen
            Percentchange = (Yearclose - Yearopen) / Yearclose
                    
            'enter change values
            Cells(Sum, 10).Value = Yearchange
            Cells(Sum, 11).Value = Percentchange
            Cells(Sum, 11).NumberFormat = "0.00%"
            
            
        'total volume in table
        Cells(Sum, 12).Value = total
        Cells(Sum, 9).Value = Cells(i, 1).Value
        total = 0
        Sum = Sum + 1
        
            
                    
            'total volume calc
             total = total + Cells(i, 7).Value
        
        End If

Next i

 
Next ws


End Sub
