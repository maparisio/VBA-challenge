'-----------------------------------------------------------------------------------
' CREATE A SCRIPT TO LOOP THROUGH ALL STOCKS FOR ONE YEAR AND OUTPUTS THE FOLLOWING:
        ' TICKER SYMBOL
        ' YEARLY CHANGE FROM OPEING PRICE TO CLOSING PRICE FOR EACH YEAR
        ' PERCENTAGE CHANGE FROM THE OPEING PRICE TO THE CLOSING PRICE FOR EACH YEAR
        ' TOTAL STOCK VOLUME FOR EACH TICKER SYMBOL
'-----------------------------------------------------------------------------------
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub multi_quarter_stock_analysis():

    ' Set dimensions
    
    Dim i As Long
    Dim start As Long
    Dim RowCount As Long
    Dim j As Integer
    Dim days As Integer
    Dim total As Double
    Dim PercentChange As Double
    Dim change As Double
    Dim DailyChange As Double
    Dim AverageChange As Double
    Dim ws As Worksheet

    For Each ws In Worksheets
        ' Set values for each worksheet
        j = 0
        total = 0
        change = 0
        start = 2
        DailyChange = 0

        ' Set title row
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "PercentChange"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        ' get the row number of the last row with data
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row

        For i = 2 To RowCount

            ' If ticker changes then print results
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Stores results in variables
                total = total + ws.Cells(i, 7).Value

                ' Handle zero total volume
                If total = 0 Then
                    ' print the results
                    ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = 0
                    ws.Range("K" & 2 + j).Value = "%" & 0
                    ws.Range("L" & 2 + j).Value = 0

                Else
                    ' Find First non zero starting value
                    If ws.Cells(start, 3) = 0 Then
                        For find_value = start To i
                            If ws.Cells(find_value, 3).Value <> 0 Then
                                start = find_value
                                Exit For
                            End If
                        Next find_value
                    End If

                    ' Calculate Change
                    change = (ws.Cells(i, 6) - ws.Cells(start, 3))
                    PercentChange = change / ws.Cells(start, 3)

                    ' start of the next stock ticker
                    start = i + 1

                    ' print the results
                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = change
                    ws.Range("J" & 2 + j).NumberFormat = "0.00"
                    ws.Range("K" & 2 + j).Value = PercentChange
                    ws.Range("K" & 2 + j).NumberFormat = "0.00%"
                    ws.Range("L" & 2 + j).Value = total

                    ' colors positives green and negatives red
                    Select Case change
                        Case Is > 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                        Case Else
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                    End Select
     Select Case change
                        Case Is > 0
                            ws.Range("K" & 2 + j).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("K" & 2 + j).Interior.ColorIndex = 3
                        Case Else
                            ws.Range("K" & 2 + j).Interior.ColorIndex = 0
                    End Select
                End If

                ' reset variables for new stock ticker
                total = 0
                change = 0
                j = j + 1
                days = 0
                DailyChange = 0

            ' If ticker is still the same add results
            Else
                total = total + ws.Cells(i, 7).Value

            End If

        Next i

        ' take the max and min and place them in a separate part in the worksheet
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & RowCount)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & RowCount)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & RowCount))

        ' returns one less because header row not a factor
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & RowCount)), ws.Range("L2:L" & RowCount), 0)

        ' final ticker symbol for  total, greatest % of increase and decrease, and average
        ws.Range("P2") = ws.Cells(increase_number + 1, 9)
        ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
        ws.Range("P4") = ws.Cells(volume_number + 1, 9)

    Next ws

End Sub

