Attribute VB_Name = "Stock_Data"
Sub Stock_Data()

'Basic Part
'Create a script table (column I to column L) with the summary of each year

Dim ws As Worksheet
For Each ws In Sheets

'Declare the variables as needed
    Dim TicketSymbol As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVol As Double
    Dim First_Row As Integer
    Dim i As Long
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    

'Set the initial values of the variables
    OpenPrice = 0
    ClosePrice = 0
    YearlyChange = 0
    PercentChange = 0
    TotalStockVol = 0
    First_Row = 2
    

'Insert headers for the table column
   ws.[I1:L1] = [{"Ticket", "Yearly Change", "Percent Change", "Total Stock Volume"}]
   
'Set the inital opening price for the first row, the rest will be adjusted in the loop
   OpenPrice = ws.Cells(2, 3).Value
'Loop through the rows to the last filled cell
    
   For i = 2 To ws.UsedRange.Rows.Count
    
    'Compare cells with previous cells
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ClosePrice = ws.Cells(i, 6).Value
    'Get the values of the variables for the table
            TicketSymbol = ws.Cells(i, 1).Value
            YearlyChange = ClosePrice - OpenPrice
            
            'Check if the opening price is not equal to zero to get the prcent change
            If OpenPrice <> 0 Then
            PercentChange = (ClosePrice - OpenPrice) / OpenPrice
            End If
            TotalStockVol = TotalStockVol + ws.Cells(i, 7).Value
        
    'Set the location of the values on the table
            ws.Range("I" & First_Row).Value = TicketSymbol
        
        'Format the cells according to its value
            ws.Range("J" & First_Row).Value = YearlyChange
            If ws.Range("J" & First_Row).Value < 0 Then
                ws.Range("J" & First_Row).Interior.ColorIndex = 3
            Else
                ws.Range("J" & First_Row).Interior.ColorIndex = 4
        
            End If
        
        'Get the percent of the change
            ws.Range("K" & First_Row).Value = PercentChange
            ws.Range("K" & First_Row).NumberFormat = "0.00%"
        
            ws.Range("L" & First_Row).Value = TotalStockVol
            First_Row = First_Row + 1
    
        'Reset the value of the variables
            YearlyChange = 0
            PercentChange = 0
            TotalStockVol = 0
            OpenPrice = ws.Cells(i + 1, 3).Value
            ClosePrice = 0
            
        Else
        'Get the value for the next row
        
            YearlyChange = YearlyChange + (ClosePrice - OpenPrice)
            If OpenPrice <> 0 Then
                PercentChange = PercentChange + ((ClosePrice - OpenPrice) / OpenPrice)
            End If
            TotalStockVol = TotalStockVol + ws.Cells(i, 7).Value
    
        End If
    
        
    Next i

'-----------------------------------------------------------------------------------
'Challenge Part
    'Get the greatest values of the stock data and set on table(O1:Q4)
    Dim Last_Row2 As Long
    Dim MaxPercent As Double
    Dim MaxTicker As String
    Dim MinPercent As Double
    Dim MinTicker As String
    Dim MaxStock As Double
    Dim MaxSTicker As String
    
    'Set the inital values of the variables
    Last_Row2 = Cells(Rows.Count, "I").End(xlUp).Row
    MaxPercent = 0
    MaxTicker = ""
    MinPercent = 0
    MinTicker = ""
    MaxStock = 0
    MaxSTicker = ""
    
    'Insert headers for the greatest values table(Challenge part)
    ws.[P1:Q1] = [{"Ticket","Value"}]
   
    'Set the titles for the greatest values
    ws.Range("O2").Formula = "Greatest% Increase"
    ws.Range("O3").Formula = "Greatest % Decrease"
    ws.Range("O4").Formula = "Greatest Total Volume"
        
    'Evaluate the greatest values
    For i = 2 To ws.UsedRange.Rows.Count
        'Get the greatest % increase
        If ws.Cells(i, 11).Value > MaxPercent Then
            MaxPercent = ws.Cells(i, 11).Value
            MaxTicker = ws.Cells(i, 9).Value
        End If
        
        'Get the greatest % decrease
        If ws.Cells(i, 11).Value < MinPercent Then
            MinPercent = ws.Cells(i, 11).Value
            MinTicker = ws.Cells(i, 9).Value
        End If
        
        'Get the greatest stock vol
        If ws.Cells(i, 12).Value > MaxStock Then
            MaxStock = ws.Cells(i, 12).Value
            MaxSTicker = ws.Cells(i, 9).Value
        
        End If
    Next i
    
    'Set the location of the greatest values
    ws.Range("P2").Value = MaxTicker
    ws.Range("Q2").Value = MaxPercent
    ws.Range("P3").Value = MinTicker
    ws.Range("Q3").Value = MinPercent
    ws.Range("P4").Value = MaxSTicker
    ws.Range("Q4").Value = MaxStock
        
      

Next ws

    
End Sub


