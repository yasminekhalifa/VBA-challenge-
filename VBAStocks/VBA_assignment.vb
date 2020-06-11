Sub Ticker2()

    Dim Ticker_symbol As String
    Dim Volume_Total As Variant
    Dim Yearly_change As Variant
    Dim Percentage_change As Variant
    Dim Close_price As Variant
    Dim Open_price As Variant
    Dim Percentage As Variant
    Dim Summary_Table_Row As Integer
    Dim interOpen As Variant
    Dim interClose As Variant
    
    interOpen = 0
    interClose = 0
    Volume_Total = 0
    Yearly_change = 0
    Percentage_change = 0
    Summary_Table_Row = 2
  
    Cells(1, 12).Value = "Ticker Symbol"
    Cells(1, 13).Value = "Total Volume"
    Cells(1, 14).Value = "Yearly Change"
    Cells(1, 15).Value = "Percentage Change"
  
i = 2
Do While Not IsEmpty(Cells(i + 1, 1))

        Close_price = Cells(i, 6).Value
        Open_price = Cells(i, 3).Value
        
        interOpen = interOpen + Open_price
        interClose = interClose + Close_price
        Volume_Total = Volume_Total + Cells(i, 7).Value
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
          Ticker_symbol = Cells(i, 1).Value
          Yearly_change = interClose - interOpen
          
          If interOpen <> 0 Then
            Percentage_change = (Yearly_change / interOpen)
          Else
            Percentage_change = 0
            End If
          
          Range("L" & Summary_Table_Row).Value = Ticker_symbol
         
    
          Range("M" & Summary_Table_Row).Value = Volume_Total
          Range("N" & Summary_Table_Row).Value = Yearly_change
          Range("O" & Summary_Table_Row).Value = Percentage_change
    
          Summary_Table_Row = Summary_Table_Row + 1
          
    
          Volume_Total = 0
          Yearly_change = 0
          Percentage_change = 0
          interClose = 0
          interOpen = 0
        End If
    i = i + 1
   Loop
   
 i = 2
Do While Not IsEmpty(Cells(i, 12))
        If Cells(i, 14).Value >= 0 Then
            Cells(i, 14).Interior.ColorIndex = 4
        Else
            Cells(i, 14).Interior.ColorIndex = 3
        End If
    
        Cells(i, 15).NumberFormat = "0.00%"
    i = i + 1
    Loop
End Sub
