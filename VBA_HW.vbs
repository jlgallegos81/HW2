Sub StockAveragePrice()

' Variables - Placeholders
  Dim Ticker As String
  Dim Total As Double
  Dim Location As Double
  Dim OPrice As Double
  Dim Cprice As Double
  Dim Pchange As Double
  
  
     
'Set starting values
    Total = 0
    Location = 2
    OPrice = 0
    Cprice = 0
    Range("i1").Value = "Ticker"
    Range("j1").Value = "Total"
    Range("k1").Value = "Opening Price"
    Range("l1").Value = "Closing Price"
    Range("m1").Value = "Change Percentage"
        
' Loop through all Tickers
    For i = 2 To 705714
        
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        OPrice = Cells(i, 3).Value
        End If
        
' Check when Ticker changes
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
        Total = Total + Cells(i, 7).Value
        Cprice = Cells(i, 6).Value
        
        
' Print the info
    Range("i" & Location).Value = Ticker
    Range("j" & Location).Value = Total
    Range("k" & Location).Value = OPrice
    Range("l" & Location).Value = Cprice
    Range("m" & Location).Value = ((Cprice - OPrice) / OPrice) * 100
        
    ' Continue moving down Ticker row
    Location = Location + 1
    
' Reset the total
    Total = 0
    
    Else
    
' If matches keep adding
    Total = Total + Cells(i, 7).Value
        
    End If
    
 Next i
  
 With Range("M2:M283")
    For Each Cell In Range("m2:M2835")
        If Cell.Value < 0 Then
            Cell.Interior.ColorIndex = 3
            Else
            Cell.Interior.ColorIndex = 4
        End If
        
        Next
        
        End With
        
  End Sub
  
