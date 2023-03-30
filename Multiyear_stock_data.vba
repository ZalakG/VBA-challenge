Sub Stock_Data()


For Each ws In Worksheets


'Create Column Name

    ws.Range("I1").Value = "Ticker Name"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Toal Stock Volumn"
    
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volumn"


'Define Variables for claculations

    Dim Ticker_Name As String
    Ticker_Name = ""
    
    Dim Ticker_Row As Long
    Ticker_Row = 2
    
    Dim Stock_Volume As Double
    Stock_Voulme = 0
    
    Dim Opening_Price As Double
    Opening_Price = 0
    
    Dim Closing_Price As Double
    Closing_Price = 0
    
    Dim Yearly_Change As Double
    Yearly_Change = 0
    
    Dim Price_Change_Percent As Double
    Price_Change_Percent = 0
    
    Dim Pct_Increase As Double
    Pct_Increase = 0
    
    Dim Pct_Increase_Ticker As String
    Pct_Increase_Ticker = " "
    
    Dim Pct_Decrease As Double
    Pct_Decrease = 0
    
    Dim Pct_Decrease_Ticker As String
    Pct_Decrease_Ticker = " "
    
    
    Dim GrtTotal_Volumn As Double
    GrtTotal_Volumn = 0
    
    Dim GrtTotal_Volumn_Ticker As String
    GrtTotal_Volumn_Ticker = " "
    
    
    Dim Lastrow As Long
    

'Loop for currentworksheet to Lastrow

    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Opening_Price = ws.Cells(2, 3).Value
   
For i = 2 To Lastrow

'Print Ticker Name

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        Ticker_Name = ws.Cells(i, 1).Value
        ws.Cells(Ticker_Row, "I").Value = Ticker_Name

'Calculate Yearly Price Change

        Closing_Price = ws.Cells(i, 6).Value
        Yearly_Change = Closing_Price - Opening_Price
        ws.Cells(Ticker_Row, "J").Value = Yearly_Change
        
        
' Fill color : Red for negative and Green for Posive change  for yearly price change

    If (ws.Cells(Ticker_Row, "J").Value < 0) Then
        ws.Cells(Ticker_Row, "J").Interior.ColorIndex = 3
    
     ElseIf (ws.Cells(Ticker_Row, "J").Value > 0) Then
             ws.Cells(Ticker_Row, "J").Interior.ColorIndex = 4

    End If
    

'calculate Yearly Change by fixing the opening price = 0 condition

     If Opening_Price <> 0 Then
        
        Price_Change_Percent = ((Closing_Price / Opening_Price) - 1) * 100
        ws.Cells(Ticker_Row, "K").Value = (CStr(Price_Change_Percent) & "%")
        
    End If

'Calculate Total Stock Volumn

        Stock_Volumn = Stock_Volumn + ws.Cells(i, 7).Value
        ws.Cells(Ticker_Row, "L").Value = Stock_Volumn
        
    
'Add 1 to the Ticker_Row count
    Ticker_Row = Ticker_Row + 1
    

'To Get next Opening price
    Opening_Price = ws.Cells(i + 1, 3).Value
    
    
'Calculate Values for Gretest % Increase, Decrease, and Total Volumn

    If (Price_Change_Percent > Pct_Increase) Then
            Pct_Increase = Price_Change_Percent
            Pct_Increase_Ticker = Ticker_Name
    
        ElseIf (Price_Change_Percent < Pct_Decrease) Then
                Pct_Decrease = Price_Change_Percent
                Pct_Decrease_Ticker = Ticker_Name
    
    End If
    
    If (Stock_Volumn > GrtTotal_Volumn) Then
            GrtTotal_Volumn = Stock_Volumn
            GrtTotal_Volumn_Ticker = Ticker_Name
    End If

            'Reset the values of Price Cahnge Percent and Stock Volumn
            
            Price_Change_Percent = 0
            Stock_Volumn = 0
    
    Else
        Stock_Volumn = Stock_Volumn + ws.Cells(i, 7).Value
        
        
End If
   

Next i

    'Print the values in "Ticker" and "Value" columns
        
               
        ws.Range("Q2").Value = (CStr(Pct_Increase) & " %")
        ws.Range("Q3").Value = (CStr(Pct_Decrease) & " %")
        ws.Range("P2").Value = Pct_Increase_Ticker
        ws.Range("P3").Value = Pct_Decrease_Ticker
        ws.Range("Q4").Value = GrtTotal_Volumn
        ws.Range("P4").Value = GrtTotal_Volumn_Ticker
        
'Set columns I To Q width based on contents of cells

    ws.Columns("I:Q").AutoFit
        

Next ws


End Sub
 




