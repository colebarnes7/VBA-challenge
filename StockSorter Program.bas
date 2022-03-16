Sub StockSorter()

    'Sets the Column Titles
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
        
        
    'Declare Variables
    Dim Ticker As String
    Dim Counter As Integer
    Dim OpenPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Single
    Dim GreatestVolume As Single
    Dim GreatestVolumeTicker As String
    Dim GreatestPercentIncrease As Double
    Dim PercentIncreaseTicker As String
    Dim GreatestPercentDecrease As Double
    Dim PercentDecreaseTicker As String
        
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
    
    'Set initial values to check against
    GreatestVolume = Range("L2").Value
    GreatestPercentIncrease = Range("K2").Value
    GreatestPercentDecrease = Range("K2").Value
    
    'For loop to determine the greatest volume, % increase and % decrease
    For I = 2 To EndRow1
        
        'Conditional for greatest volume, grabs ticker as well
        If (Cells(I, 12).Value > GreatestVolume) Then
            GreatestVolume = Cells(I, 12).Value
            GreatestVolumeTicker = Cells(I, 9).Value
            
        'Conditional for greatest % increase, grabs ticker as well
        ElseIf (Cells(I, 11).Value > GreatestPercentIncrease) Then
            GreatestPercentIncrease = Cells(I, 11).Value
            PercentIncreaseTicker = Cells(I, 9).Value
        
        'Conditional for greatest % decrease, grabs ticker as well
        ElseIf (Cells(I, 11).Value < GreatestPercentDecrease) Then
            GreatestPercentDecrease = Cells(I, 11).Value
            PercentDecreaseTicker = Cells(I, 9).Value
        
        End If
     Next I
     
     'Outputs greatest volume, % increase and % decrease
     Range("P2").Value = PercentIncreaseTicker
     Range("P3").Value = PercentDecreaseTicker
     Range("P4").Value = GreatestVolumeTicker
     Range("Q2").Value = GreatestPercentIncrease
     Range("Q3").Value = GreatestPercentDecrease
     Range("Q4").Value = GreatestVolume
     
     'Puts the bonus values into percent format
     Range("Q2:Q3").NumberFormat = "0.00%"
     

End Sub
