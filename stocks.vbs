Attribute VB_Name = "Module1"
Sub Stock_Data()

'Define variables
Dim Ticker As String
Dim YearOpen As Double
Dim YearClose As Double
Dim Yearly_Change As Double
Dim Total_Stock_Volume As Double
Dim Percent_Change As Double
Dim NextRow As Integer
Dim ws As Worksheet

' Apply to all worksheets

For Each ws In Worksheets

    'Output headers

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'Starting loop
    NextRow = 2
    NewTicker = 1
    Total_Stock_Volume = 0

    ' Last row for loop
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Loop rows for info

        For i = 2 To LastRow

            ' Identify ticker symbols
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            'Ticker symbol
            Ticker = ws.Cells(i, 1).Value

            ' Starts loop at first ticker value
            NewTicker = NewTicker + 1

            ' First value for stock
            YearOpen = ws.Cells(NewTicker, 3).Value
            ' Last value for stock
            YearClose = ws.Cells(i, 6).Value

            ' loop for total stock volume

            For j = NewTicker To i

                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(j, 7).Value

            Next j

            ' Get Yearly and Percent Change

            If YearOpen = 0 Then

                Percent_Change = YearClose

            Else
                Yearly_Change = YearClose - YearOpen

                Percent_Change = Yearly_Change / YearOpen

            End If
         '--------------------------------------------------

            ' Output values

            ws.Cells(NextRow, 9).Value = Ticker
            ws.Cells(NextRow, 10).Value = Yearly_Change
            ws.Cells(NextRow, 11).Value = Percent_Change

            ' Format as percent

            ws.Cells(NextRow, 11).NumberFormat = "0.00%"
            ws.Cells(NextRow, 12).Value = Total_Stock_Volume

            ' Fill out next row

            NextRow = NextRow + 1

            ' Start at 0

            Total_Stock_Volume = 0
            Yearly_Change = 0
            Percent_Change = 0

            ' Continue to new ticker
            NewTicker = i

        End If

    'Done the loop

    Next i
    
    ' Get last row for yearly change
    YCLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row


        For i = 2 To YCLastRow

            'if greater than or less than zero
            If ws.Cells(i, 10) > 0 Then

                ws.Cells(i, 10).Interior.ColorIndex = 4

            Else

                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If

        Next i

'Excute to next worksheet
Next ws

End Sub
