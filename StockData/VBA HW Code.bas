Attribute VB_Name = "Module1"
Sub Wall_Street():
For Each ws In Worksheets
    ' establishing variables
    
    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Double

    ' Setting header titles
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"

    ' Set initial variable values
    j = 0
    total = 0
    change = 0
    start = 2

    ' getting last row with data
    rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row

    For i = 2 To rowCount

        ' If ticker changes then print results
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ' Storing results
            total = total + ws.Cells(i, 7).Value

            ' in the event there is a 0
            If total = 0 Then
                ' results printed
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = 0
                ws.Range("K" & 2 + j).Value = "%" & 0
                ws.Range("L" & 2 + j).Value = 0

            Else
                ' Find the first non zero value
                If ws.Cells(start, 3) = 0 Then
                    For find_value = start To i
                        If ws.Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                     Next find_value
                End If

                ' Change calculation
                change = (ws.Cells(i, 6) - ws.Cells(start, 3))
                percentChange = Round((change / ws.Cells(start, 3) * 100), 2)

                ' start of the next ticker
                start = i + 1

                ' results printed
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = Round(change, 2)
                ws.Range("K" & 2 + j).Value = "%" & percentChange
                ws.Range("L" & 2 + j).Value = total

                ' color coding based on yearly change
                Select Case change
                    Case Is > 0
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Is < 0
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                End Select

            End If

            ' resetting for new ticker
            total = 0
            change = 0
            j = j + 1
            Days = 0

        ' If ticker the same add to total
        Else
            total = total + ws.Cells(i, 7).Value

        End If

    Next i

    ' max and min separate from summary table
    ws.Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & rowCount)) * 100
    ws.Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & rowCount)) * 100
    ws.Range("Q4") = WorksheetFunction.Max(Range("L2:L" & rowCount))

    ' returns one less to avoid header row
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & rowCount)), Range("L2:L" & rowCount), 0)

    ' ticker for greatest total volume, greatest % of increase and decrease
    ws.Range("P2") = ws.Cells(increase_number + 1, 9)
    ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
    ws.Range("P4") = ws.Cells(volume_number + 1, 9)

Next ws

End Sub
