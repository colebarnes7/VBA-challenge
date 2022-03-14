Attribute VB_Name = "Module1"
Sub StockSorter()
    'Sets the Column Titles
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    'Declare Variables
    Dim Ticker As String
    Dim Counter As Integer
    Dim OpenPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Single
    
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
    For i = 2 To EndRow
        
        'Checks if new ticker is coming up
        If (Cells(i + 1, 1).Value <> Cells(i, 1).Value) Then
            
            'Grabs current ticker number
            Ticker = Cells(i, 1).Value
            
            'Add last number of stocks to volume
            TotalVolume = TotalVolume + Cells(i, 7).Value
            
            'Calculates the yearly change of the ticker
            YearlyChange = Cells(i, 6).Value - OpenPrice
            
            'Calculates the percent change of the ticker
            PercentChange = (Cells(i, 6).Value - OpenPrice) / OpenPrice
            
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
            OpenPrice = Cells(i + 1, 3).Value
            
        'If not a new ticker, add current day to total volume
        Else
            TotalVolume = TotalVolume + Cells(i, 7).Value
        
        End If
    Next i

     'Converts the percent change row to percentage format
    EndRow1 = Cells(Rows.Count, 11).End(xlUp).Row
    Range("K2:K" & EndRow1).NumberFormat = "0.00%"

    'Formats the yearly change row to green for positive and red for negative
    EndRow2 = Cells(Rows.Count, 11).End(xlUp).Row
    For i = 2 To EndRow2
        If (Cells(i, 11).Value > 0) Then
            Cells(i, 11).Interior.ColorIndex = 4
            
        ElseIf (Cells(i, 11).Value < 0) Then
            Cells(i, 11).Interior.ColorIndex = 3
        
        End If
    Next i
End Sub
