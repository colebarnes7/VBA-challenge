Attribute VB_Name = "Module1"
Sub StockSorter()

    'Sets the Column Titles
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Value"
        
        
    'Declare Variables
    Dim Ticker As String
    Dim Counter As Integer
    Dim OpenPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Single
    Dim GreatestVolume As Single
    Dim GreatestPercentIncrease As Double
    Dim GreatestPercentDecrease As Double
    Dim R As Range
        
    'Determines the last row in the active sheet for use in loops
    Dim EndRow As Long
    EndRow = Cells(Rows.Count, 1).End(xlUp).Row
        
    'Set initial values
    OpenPrice = Range("C2").Value
    YearlyChange = 0
    PercentChange = 0
    TotalVolume = 0
        
    'Variable to store the row to place the sorted data
    Counter = 2
        
    'For loop to iterate through the rows
    For I = 2 To EndRow
            
        'Checks if new ticker is coming up
        If (Cells(I + 1, 1).Value <> Cells(I, 1).Value) Then
                
            'Grabs current ticker number
            Ticker = Cells(I, 1).Value
                
            'Add last number of stocks to volume
            TotalVolume = TotalVolume + Cells(I, 7).Value
                
            'Calculates the yearly change of the ticker
            YearlyChange = Cells(I, 6).Value - OpenPrice
                
            'Calculates the percent change of the ticker
            PercentChange = (Cells(I, 6).Value - OpenPrice) / OpenPrice
                
            'Sets ticker, yearly change, percent change, volume into correct location
            Cells(Counter, 9).Value = Ticker
            Cells(Counter, 10).Value = YearlyChange
            Cells(Counter, 11).Value = PercentChange
            Cells(Counter, 12).Value = TotalVolume
                
            'Adds to counter to place new data in correct row
            Counter = Counter + 1
                
            'Resets volume for the next ticker
            TotalVolume = 0
                
            'Grabs opening price for new ticker
            OpenPrice = Cells(I + 1, 3).Value
                
        'If not a new ticker, add current day to total volume
        Else
            TotalVolume = TotalVolume + Cells(I, 7).Value
            
        End If
    Next I
    
    'Converts the percent change row to percentage format
    EndRow1 = Cells(Rows.Count, 11).End(xlUp).Row
    Range("K2:K" & EndRow1).NumberFormat = "0.00%"
    
    'Formats the yearly change row to green for positive and red for negative
    For I = 2 To EndRow1
        If (Cells(I, 11).Value > 0) Then
            Cells(I, 11).Interior.ColorIndex = 4
                
        ElseIf (Cells(I, 11).Value < 0) Then
            Cells(I, 11).Interior.ColorIndex = 3
            
        End If
    Next I
    
    'Sets Range to look for Greatest % increase and decrease
    Set R = Range("K2:K" & Rows.Count)
        
    'Grabs the Greatest % Increase and outputs it
    GreatestPercentIncrease = Application.WorksheetFunction.Max(R)
    Range("P2").Value = GreatestPercentIncrease
        
    'Grabs the Greatest % Decrease and outputs it
    GreatestPercentDecrease = Application.WorksheetFunction.Min(R)
    Range("P3").Value = GreatestPercentDecrease
        
    'Grabs the Greatest Total Volume and outputs it
    Set R = Range("L2:L" & Rows.Count)
    GreatestTotalVolume = Application.WorksheetFunction.Max(R)
    Range("P4").Value = GreatestTotalVolume
        
    'Formats the Greatest % increase and decrease column to percentages
    Range("P2:P3").NumberFormat = "0.00%"
        
End Sub
