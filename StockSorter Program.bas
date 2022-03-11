Attribute VB_Name = "Module1"
Sub StockSorter()
    'Sets the Column Titles
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    'Declare Variables
    Dim Ticker As String
    Dim PlaceHolder As Integer
    
    'Determines the last row in the active sheet for use in loops
    Dim EndRow As Long
    EndRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Variable to store the row to place the sorted data
    PlaceHolder = 2
    
    'Sets initial ticker to check against
    Range("I2").Value = Range("A2").Value
    
    'For loop to iterate through the rows
    For i = 2 To EndRow
        'Sets Ticker equal to row i in column 1
        Ticker = Cells(i, 1).Value
        
        'Conditional to check if a new ticker has been found
        If (Ticker <> Cells(PlaceHolder, 9).Value) Then
            'Places new ticker in the correct row
            PlaceHolder = PlaceHolder + 1
            Cells(PlaceHolder, 9).Value = Ticker
        End If
    Next i 
End Sub
