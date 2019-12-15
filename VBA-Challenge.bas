Attribute VB_Name = "Module1"
Sub test():

' Looping through worksheets
For Each ws In Worksheets
    
    ' declaring the ticker to a string
    Dim Ticker As String
    
    ' declaring the yearly change
    Dim Opening_Price
    Opening_Price = Range("C2").Value
    
    Dim Yearly_Change
    
    ' variable for holding the total stock volume
    Dim Stock_Volume As Double
    Stock_Volume = 0
    
    ' keeping track of the new location of the new columns
    Dim Summary_Stock As Integer
    Summary_Stock = 2
    
    MsgBox (ws.Name)
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    
    ' add the word ticker to the new column
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 11).Value = "Percent Change"
    
    ' looping through the columns
    For i = 2 To LastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        ' Assigning the ticker values to the Ticker variable
        Ticker = ws.Cells(i, 1).Value
        
        ' printing the tickers to the new column
        ws.Range("I" & Summary_Stock).Value = Ticker
        
        
        ' Assigning the opening price to the Opening_Price variable
        ws.Cells(i + 1, 3).Value = Opening_Price
        
        ' subtracting the difference of open and close for yearly change
        Yearly_Change = ws.Cells(i, 6).Value - Opening_Price
        
        'printing the yearly change to the new column
        ws.Range("J" & Summary_Stock).Value = Yearly_Change
        
        
            ' Conditional formatting If statement for the yearly change
            If ws.Range("J" & Summary_Stock).Value < 0 Then
        
            ws.Range("J" & Summary_Stock).Interior.ColorIndex = 3
        
            ElseIf ws.Range("J" & Summary_Stock).Value > 0 Then
        
            ws.Range("J" & Summary_Stock).Interior.ColorIndex = 4
        
            End If
        
        
        ' Get the percentage of the yearly change
        Percent_Change = Yearly_Change / Opening_Price
        
        ws.Range("K" & Summary_Stock).NumberFormat = "0.00%"
        
        ws.Range("K" & Summary_Stock).Value = Percent_Change
        
        
        ' Adding to the total stock volume and assigning to variable
        Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
        
        ' printing the total stock volume to the new column
        ws.Range("L" & Summary_Stock).Value = Stock_Volume
        
        ' Adding the next row in the summary column
        Summary_Stock = Summary_Stock + 1
        
        ' reseting total stock volume
        Stock_Volume = 0
    
        Else
        Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
        
        End If
        
        
    Next i
    
Next ws

End Sub
