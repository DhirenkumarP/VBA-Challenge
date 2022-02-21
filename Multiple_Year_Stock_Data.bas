Attribute VB_Name = "Module1"
Sub Multiple_Year_Stock_data():

'------------------------------------------------------------------------------------------------------------------------------------------------
' 1.The ticker symbol.
' 2.Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
' 3.The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
' 4.The total stock volume of the stock.
'------------------------------------------------------------------------------------------------------------------------------------------------

    Dim Ticker As String
    Dim OpenStock As Double
    Dim CloseStock As Double
    Dim Vol As Double
    Vol = 0
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    
' Find Last Row of column

    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim Summary_Table As Integer
    Summary_Table = 2
    
    
' Assign Column and Column Header

    Range("I1").EntireColumn.Insert
    Cells(1, 9).Value = "Ticker"
    Range("J1").EntireColumn.Insert
    Cells(1, 10).Value = "Yearly Change"
    Range("K1").EntireColumn.Insert
    Cells(1, 11).Value = "Percent Change"
    Range("L1").EntireColumn.Insert
    Cells(1, 12).Value = "Total Stock Volume"
    
    For i = 2 To LastRow
        
 ' Get the required Value of each Rows and columns
 
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
               Ticker = Cells(i, 1).Value
               Total_Stock_Volume = Vol + Cells(i, 7).Value
            
               OpenStock = Cells(i, 3).Value
               CloseStock = Cells(i, 6).Value
               
               Yearly_Change = CloseStock - OpenStock
                  
               Percent_Change = (CloseStock - OpenStock) / CloseStock
               
               Range("I" & Summary_Table).Value = Ticker
               Range("L" & Summary_Table).Value = Total_Stock_Volume
               Range("J" & Summary_Table).Value = Yearly_Change
               Range("K" & Summary_Table).Value = Percent_Change
               
               Columns("K").NumberFormat = "0.00%"
               
               Summary_Table = Summary_Table + 1
               
               Vol = 0
               
            Else
            
               Vol = Vol + Cells(i, 7).Value
               
            End If
            
        
    Next i
            
' Get the interior color of cells as per Value

    LastRow = Cells(Rows.Count, 10).End(xlUp).Row
    
          For i = 2 To LastRow
          
             If Cells(i + 1, 10) >= 0 Then
               Cells(i, 10).Interior.ColorIndex = 3
               
            Else
               Cells(i, 10).Interior.ColorIndex = 4
               
            End If
    Next i
    
    
End Sub
