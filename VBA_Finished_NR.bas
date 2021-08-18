Attribute VB_Name = "Module1"
Sub Datacleaner()

'loop to cycle through sheets.

 For Each ws In Worksheets
 
' Creating Variables
Dim Ticker As String
Dim Dates As Double
Dim Openprice As Double
Dim closingprice As Double
Dim Volume As Double
Dim Lastrow As Double
Dim sumtable As Double

'Initiate starting values.

    sumtable = 2
    Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Volume = 0
    Openprice = ws.Cells(2, 3).Value
    
    'Creating headers.
  
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    'For loop to identify changes in ticker and do most of the calculations and print the values.
    
    For i = 2 To Lastrow
    
           If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value And Openprice <> 0 Then
                
                Ticker = ws.Cells(i, 1).Value
                
                Volume = Volume + ws.Cells(i, 7).Value
                
                closingprice = ws.Cells(i, 6).Value
                
                ws.Range("I" & sumtable).Value = Ticker
              
                ws.Range("L" & sumtable).Value = Volume
                
                ws.Range("J" & sumtable).Value = (closingprice - Openprice)
                
                ws.Range("K" & sumtable).Value = ((closingprice - Openprice) / Openprice)
                
                ws.Columns("K").NumberFormat = "0.00%"
                     
                sumtable = sumtable + 1
                
                Volume = 0
                
                Openprice = ws.Cells(i + 1, 3).Value
                
                
Else

Volume = Volume + ws.Cells(i, 7).Value
                
                   End If
                   
                   Next i
                   
    'Loops to do conditional formatting.
    
    For i = 2 To Lastrow
             
             If ws.Cells(i, 10) > 0 Then
             
             ws.Range("J" & i).Interior.ColorIndex = 4
             
             ElseIf ws.Cells(i, 10) < 0 Then
             
             ws.Range("J" & i).Interior.ColorIndex = 3
             
             End If
             
             Next i
             
    'Loops to do conditional formatting.
             
    For i = 2 To Lastrow
             
             If ws.Cells(i, 11) > 0 Then
             
             ws.Range("K" & i).Interior.ColorIndex = 4
             
             ElseIf ws.Cells(i, 11) < 0 Then
             
             ws.Range("K" & i).Interior.ColorIndex = 3
             
             End If
             
             Next i
             
  
    'Finding the stocks with the highest and lowest percentages and volume. Also formats the cells.
        
        ws.Cells(2, 17) = Application.WorksheetFunction.Max(ws.Range("K:K"))
        
        ws.Cells(3, 17) = Application.WorksheetFunction.Min(ws.Range("K:K"))
        
        ws.Cells(2, 17).NumberFormat = "0.00%"
       
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ws.Cells(4, 17) = Application.WorksheetFunction.Max(ws.Range("L:L"))
        
        ws.Cells(4, 17).NumberFormat = "#0.0000E+0"
        
        
             
    'Finding tickers for the highest value stocks.
        
            For i = 2 To Lastrow
              
                Newtick = ws.Cells(i, 9)
             
                 If ws.Cells(i, 11) = ws.Cells(2, 17) Then
             
                 ws.Cells(2, 16) = Newtick
                 
                 ElseIf ws.Cells(i, 11) = ws.Cells(3, 17) Then
             
                    ws.Cells(3, 16) = Newtick
                    
                 ElseIf ws.Cells(i, 12) = ws.Cells(4, 17) Then
             
             ws.Cells(4, 16) = Newtick

                  End If
             
            Next i
              
             
 'Corrects the format and resets all the variables to start next sheet.
              
 ws.Columns("A:P").AutoFit
             

Ticker = 0
tickertotal = 0
Dates = 0
Openprice = 0
closingprice = 0
Volume = 0
Lastrow = 0
sumtable = 2
             
                   Next ws
                   
                   
                    
                
End Sub


