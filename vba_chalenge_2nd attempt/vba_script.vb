Sub RunCode()
    
    'Defining the variables
    Dim Summary_table As Integer
    Summary_table = 2
    Dim Ticker As String
    Dim Yearly_change As Double
    Dim Percentage_change As Double
    Dim Total_stock As Double
    Dim Year_open As Double
    Dim Year_close As Double
    Dim lastrow As Long
    
    'headers
    Range("L1").Value = "Ticker"
    Range("M1").Value = "Yearly Change"
    Range("N1").Value = "Percentage Change"
    Range("O1").Value = "Total Stock Volume"
    
    Range("T1").Value = "Ticker"
    Range("U1").Value = "Value"
    Range("S2").Value = "Greatest % Increase"
    Range("S3").Value = "Greatest % Decrease"
    Range("S4").Value = "Greatest Total Volume"
    
    'Fomula for the last row

    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Year_open = Cells(2, 3).Value
    

    'Loop through the data
For i = 2 To lastrow
    
    'Check if the ticker name matches
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    
    'Set ticker
    Ticker = Cells(i, 1).Value
    
    'fomula for yearly change
    Year_close = Cells(i, 6).Value
    Yearly_change = Year_close - Year_open
    
    'fomula for the colour change
    If Yearly_change > 0 Then
        Range("M" & Summary_table).Interior.ColorIndex = 4
    
        ElseIf Yearly_change < 0 Then
        Range("M" & Summary_table).Interior.ColorIndex = 3
    
        End If
    
    'Calculation of percentage change
    If Year_open <> 0 Then
        Percentage_change = (Yearly_change / Year_open) * 100
        
    Else
        Percentage_change = 0
    
    End If
    
    'Total stock value
    Total_stock = Total_stock + Cells(i, 7).Value
    
    'Print summary table
    Range("L" & Summary_table).Value = Ticker
    Range("O" & Summary_table).Value = Total_stock
    Range("M" & Summary_table).Value = Yearly_change
    Range("N" & Summary_table).Value = (CStr(Percentage_change) & "%")
    
    
    Summary_table = Summary_table + 1
    
    Year_open = Cells(i + 1, 3).Value

    'Reset stock total
    Total_stock = 0
   
   
    Else
    
        Total_stock = Total_stock + Cells(i, 7).Value
    
    End If
    
  Next i
  
  'Fomula for maximum values
  Range("U2").Value = WorksheetFunction.Max(Range("N:N"))
  Range("U2").NumberFormat = "0.00%"
  Range("U3").Value = WorksheetFunction.Min(Range("N:N"))
  Range("U3").NumberFormat = "0.00%"
  Range("U4").Value = WorksheetFunction.Max(Range("O:O"))
  
  
  'Assigning the ticker to the values
  For i = 2 To lastrow
  If Range("U2").Value = Cells(i, 14).Value Then
  Range("T2").Value = Cells(i, 12).Value
  
  End If
  
  If Range("U3").Value = Cells(i, 14).Value Then
  Range("T3").Value = Cells(i, 12).Value
  
  End If
  
  
  If Range("U4").Value = Cells(i, 15).Value Then
  Range("T4").Value = Cells(i, 12).Value
  
  End If
  
  Next i
 
End Sub
