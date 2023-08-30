Sub Stockanalysis()

 ' Set the dimensions'
 
    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim start As Long
    Dim RowCount As Long
    Dim percent_Change As Double
    Dim days As Integer
    Dim daily_Change As Double
    Dim average_Change As Double
    Dim wk As Worksheet
    
   For Each wk In Worksheets
   
        ' Print out values for each worksheet'
        j = 0
        total = 0
        change = 0
        start = 2
        dailyChange = 0
        'make title now'
        wk.Range("I1").Value = "Ticker"
        wk.Range("J1").Value = "Yearly_Change"
        wk.Range("K1").Value = "Percentage_Change"
        wk.Range("L1").Value = "Total_Stock_Volume"
        wk.Range("P1").Value = "Ticker"
        wk.Range("Q1").Value = "Value"
        wk.Range("O2").Value = "Greatest % Increase"
        wk.Range("O3").Value = "Greatest % Decrease"
        wk.Range("O4").Value = "Greatest Total Volume"
         ' Print rownumber of the last row with data'
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        For i = 2 To RowCount
            ' If Ticker changes then print results'
            If wk.Cells(i + 1, 1).Value <> wk.Cells(i, 1).Value Then
                  ' Stores results as variables'
                total = total + wk.Cells(i, 7).Value
                ' Print zero total volume'
                If total = 0 Then
                    ' Print out the results as'
                    wk.Range("I" & 2 + j).Value = Cells(i, 1).Value
                    wk.Range("J" & 2 + j).Value = 0
                    wk.Range("K" & 2 + j).Value = "%" & 0
                    wk.Range("L" & 2 + j).Value = 0
                Else
                    ' Check non-zero starting value'
                    If wk.Cells(start, 3) = 0 Then
                        For find_value = start To i
                            If wk.Cells(find_value, 3).Value <> 0 Then
                                start = find_value
                                Exit For
                            End If
                        Next find_value
                    End If
          ' Calculate Change'
                    change = (wk.Cells(i, 6) - wk.Cells(start, 3))
                    percent_Change = change / wk.Cells(start, 3)
                    ' start WITH THE next stock Ticker'
                    start = i + 1
                    ' Print results AS'
                    wk.Range("I" & 2 + j).Value = wk.Cells(i, 1).Value
                    wk.Range("J" & 2 + j).Value = change
                    wk.Range("J" & 2 + j).NumberFormat = "0.00"
                    wk.Range("K" & 2 + j).Value = percent_Change
                    wk.Range("K" & 2 + j).NumberFormat = "0.00%"
                    wk.Range("L" & 2 + j).Value = total
                     ' Color OUT  positive green and negative red'
                    Select Case change
                        Case Is > 0
                            wk.Range("J" & 2 + j).Interior.ColorIndex = 4
                        Case Is < 0
                            wk.Range("J" & 2 + j).Interior.ColorIndex = 3
                        Case Else
                            wk.Range("J" & 2 + j).Interior.ColorIndex = 0
                    End Select
                     End If
                ' Reset variables for new stock Ticker'
                total = 0
                change = 0
                j = j + 1
                days = 0
                daily_Change = 0
            ' If Ticker is still the same add results'
            Else
                total = total + wk.Cells(i, 7).Value
            End If
        Next i
        'Take the Max and Min & put then in separate cells in worksheet'
        wk.Range("Q2") = "%" & WorksheetFunction.Max(wk.Range("K2:K" & RowCount)) * 100
        wk.Range("Q3") = "%" & WorksheetFunction.Min(wk.Range("K2:K" & RowCount)) * 100
         wk.Range("Q4") = WorksheetFunction.Max(wk.Range("L2:L" & RowCount))
           ' Return one less As header row not a counted'
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(wk.Range("K2:K" & RowCount)), wk.Range("K2:K" & RowCount), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(wk.Range("K2:K" & RowCount)), wk.Range("K2:K" & RowCount), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(wk.Range("L2:L" & RowCount)), wk.Range("L2:L" & RowCount), 0)
        ' final Ticker symbol for  total, greatest % of increase and decrease, and average'
        wk.Range("P2") = wk.Cells(increase_number + 1, 9)
        wk.Range("P3") = wk.Cells(decrease_number + 1, 9)
        wk.Range("P4") = wk.Cells(volume_number + 1, 9)
        
        Next wk
        
        End Sub
