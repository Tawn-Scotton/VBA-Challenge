Attribute VB_Name = "Module1"

'This is the public function that can be used as a macro
Public Sub Stock_Analysis()
    SetHeaders
    Yearly_Stock_Values
End Sub
'This function sets the new headers
Private Sub SetHeaders()
    Dim sheet1Rng As Range
    Set sheet1Rng = Sheet1.Range("A1").CurrentRegion
    sheet1Rng.Cells(1, 9).Value = "Ticker"
    sheet1Rng.Cells(1, 10).Value = "Yearly Change"
    sheet1Rng.Cells(1, 11).Value = "Percent Change"
    sheet1Rng.Cells(1, 12).Value = "Total Stock Volume"
End Sub
'This function generates the new column data for the yearly stock values
Private Sub Yearly_Stock_Values()
    'This will hold the row count
    Dim x As Long
    'This the daily ticker name
    Dim dailyTickerName As String
    'This is the range for the active sheet
    Dim rng As Range
    'This is the daily change value
    Dim dailyChangeValue As Double
    'This is the total yearly change value
    Dim yearlyChangeValue As Double
    'This is the daily percent change value
    Dim dailyPercentChange As Double
    'This is the total yearly change percent
    Dim yearlyPercentChange As Double
    'This is the daily stock volume - using Double because Long caused overflow
    Dim dailyStockVolume As Double
    'This is the yearly stock volume - using Double because Long caused overflow
    Dim yearlyStockVolume As Double
    'This is the cell row number like the X coordinate
    Dim cellRowValue As Integer
    'This will hold the range for Sheet 1
    Dim sheet1Rng As Range
    'The row starts at 2 because the headers are at row 1
    cellRowValue = 2
    'Set the range for sheet 1 to add the header because all the new headers will be on sheet 1
    Set sheet1Rng = Sheet1.Range("A1").CurrentRegion
    'The first row, 10th column (J) is given the text Yearly Change
    sheet1Rng.Cells(1, 10).Value = "Yearly Change"
    
    ' Loop through all of the worksheets in the active workbook.
    For Each Sheet In Worksheets
        'Set the yearly values to 0 for each NEW sheet
        yearlyChangeValue = 0
        yearlyPercentChange = 0
        yearlyStockVolume = 0
        'Get the range for the active sheet found in the loop above
        Set rng = Sheet.Range("A1").CurrentRegion
        'Loop through the rows using the range of the active sheet
        For x = 2 To rng.Rows.Count
            'Get the daily ticker name from the current row and the first column(A)
            dailyTickerName = rng.Cells(x, 1).Value
            'Get the daily change value by subtracting the daily opening value from the close value
            dailyChangeValue = rng.Cells(x, 6).Value - rng.Cells(x, 3).Value
            'Add the daily change value to the yearly change for the current ticker symbol
            yearlyChangeValue = yearlyChangeValue + dailyChangeValue
            'Get the daily percentage change value by subtracting the daily opening value from the close value and then dividing by the opening value
            dailyPercentChange = (rng.Cells(x, 6).Value - rng.Cells(x, 3).Value) / rng.Cells(x, 3).Value
            'Add the daily percent change value to the yearly percent change for the current ticker symbol
            yearlyPercentChange = yearlyPercentChange + dailyPercentChange
            'Get the daily stock volum from the current row and column 7 (G)
            dailyStockVolume = rng.Cells(x, 7).Value
            'Add the daily stock volume to the yearly stock volume for the current ticker symbol
            yearlyStockVolume = yearlyStockVolume + dailyStockVolume
            'Check if the ticker symbol will change on the next line.  If it is changing, set the ticker symbol,
            'the yearly change value, the yearly percentage change, and the yearly stock volume in their respective cells on Sheet 1
            If dailyTickerName <> rng.Cells(x + 1, 1).Value Then
                'Set yearly values
                Sheet1.Range("A1").Cells(cellRowValue, 9).Value = dailyTickerName
                Sheet1.Range("A1").Cells(cellRowValue, 10).Value = yearlyChangeValue
                Sheet1.Range("A1").Cells(cellRowValue, 11).Value = CStr(yearlyPercentChange * 100) + "%"
                Sheet1.Range("A1").Cells(cellRowValue, 12).Value = yearlyStockVolume
                'Set yearly change value colors
                If Sheet1.Range("A1").Cells(cellRowValue, 10).Value >= 0 Then
                    Sheet1.Range("A1").Cells(cellRowValue, 10).Interior.ColorIndex = 4
                Else
                    Sheet1.Range("A1").Cells(cellRowValue, 10).Interior.ColorIndex = 3
                End If
                'Increment the cell row by 1 for the next ticker symbol and yearly values
                cellRowValue = cellRowValue + 1
                'Reset the yearly change values for each different Ticker code
                yearlyChangeValue = 0
                yearlyPercentChange = 0
                yearlyStockVolume = 0
            End If
        Next x
    Next
End Sub


