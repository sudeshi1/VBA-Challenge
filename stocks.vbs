'Create a script that will loop through all the stocks for one year and output the following information:
'- The ticker symbol.
'- Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'- The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'- The total stock volume of the stock.
'You should also have conditional formatting that will highlight positive change in green and negative change in red.
'BONUS
'Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and
'"Greatest total volume"
'Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet,
'i.e., every year, just by running the VBA script once.

Sub Stocks()

'----Loop through all the worksheets---
For Each ws In Worksheets

'Activate ws
ws.Activate

'Set an initial variable for holding ticker
Dim Ticker_Name As String

'Set an initial variable for holding the total stocks volume per ticker
Dim Total_Volume As Double
Total_Volume = 0

'Set an initial variable for open price
Dim Open_Price As Double
Open_Price = Cells(2, 3).Value

'Set an initial variable for close price
Dim Close_Price As Double
Close_Price = 0

'Set an initial variable for yearly price
Dim Yearly_Price As Double
Yearly_Price = 0

'Set an initial variable for percent change
Dim Percent_Change As Double
Percent_Change = 0

'Set an initial variable for greatest percent increase with respective ticker
Dim Inc_Percent As Double
Inc_Percent = 0
Dim Inc_Percent_Ticker As String

'Set an initial variable for greatest percent decrease with respective ticker
Dim Dec_Percent As Double
Dec_Percent = 0
Dim Dec_Percent_Ticker As String

'Set an initial variable for greatest stock with respective ticker
Dim Max_Vol As Double
Max_Vol = 0
Dim Max_Vol_Ticker As String

'Summary Table
Dim Summary As Integer
Summary = 2

'Set an initial variable to keep track of last row in each worksheet
Dim LastRow As Long
LastRow = Cells(Rows.Count, "A").End(xlUp).Row

'Looping through all the tickers
For i = 2 To LastRow

    'Check if we are within the same ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'Set the ticker name
    Ticker_Name = Cells(i, 1).Value
    'Calculate yearly price and percent change
    Close_Price = Cells(i, 6).Value
    Yearly_Price = Close_Price - Open_Price
    
    If Open_Price <> 0 Then
    Percent_Change = (Yearly_Price / Open_Price)
    End If
    
    'Calculate the total stock volume
    Total_Volume = Total_Volume + Cells(i, 7).Value
    
    'Display the ticker and total stock volume
    Range("I" & Summary).Value = Ticker_Name
    Range("J" & Summary).Value = Yearly_Price
    'Fill "Yearly Change" with Green and Red colors
                If (Yearly_Price > 0) Then
                    'Fill column with GREEN color
                    Range("J" & Summary).Interior.ColorIndex = 4
                    Range("J" & Summary).Font.ColorIndex = 5
                ElseIf (Yearly_Price <= 0) Then
                    'Fill column with RED color
                    Range("J" & Summary).Interior.ColorIndex = 3
                    Range("J" & Summary).Font.ColorIndex = 6
                End If
    Range("K" & Summary).Value = Percent_Change
    Range("L" & Summary).Value = Total_Volume
    
    'Add one to summary table
    Summary = Summary + 1
    
    'Reset total volume
    Total_Volume = 1
    
    'Reset yearly price, close price and open price
    Yearly_Price = 0
    Percent_Change = 0
    Close_Price = 0
    Open_Price = Cells(i + 1, 3).Value
    
'If the cell immediately following is the same brand
Else
    
    'Adding the total stocks volume
    Total_Volume = Total_Volume + Cells(i, 7).Value
        
    End If
Next i

    'Set last row in summary table
    Last_Row = Cells(Rows.Count, "I").End(xlUp).Row

    'Bonus
    For i = 2 To Last_Row
    
        'Setting Percent Increase
        If Cells(i, 11).Value > Inc_Percent Then
        Inc_Percent = Cells(i, 11).Value
        Inc_Percent_Ticker = Cells(i, 9).Value
        End If
        
        'Setting Percent Decrease
        If Cells(i, 11).Value < Dec_Percent Then
        Dec_Percent = Cells(i, 11).Value
        Dec_Percent_Ticker = Cells(i, 9).Value
        End If
    
        'Setting Greatest Total Stocks
        If Cells(i, 12).Value > Max_Vol Then
        Max_Vol = Cells(i, 12).Value
        Max_Vol_Ticker = Cells(i, 9).Value
        End If
 
    'Display Greatest Percent Increase, Greatest Percent Increase Ticker
    Range("P2").Value = Format(Inc_Percent, "Percent")
    Range("O2").Value = Inc_Percent_Ticker
    'Display Greatest Percent Decrease, Greatest Percent Decrease Ticker
    Range("P3").Value = Format(Dec_Percent, "Percent")
    Range("O3").Value = Dec_Percent_Ticker
    'Display Greatest Total Stocks, Greatest Total Stocks Ticker
    Range("P4").Value = Max_Vol
    Range("O4").Value = Max_Vol_Ticker
    
Next i

Next ws

End Sub

