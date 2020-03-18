Attribute VB_Name = "Module1"
Sub WallStreetData()

For Each ws In Worksheets

  ' Set variables
  Dim TickerSymbol As String
  Dim TotalVolume As Double
  Dim TickerRowTotal As Integer
  Dim OpenBegPrice As Double
  Dim CloseEndPrice As Double
  Dim YearlyChg As Double
  
  'Variable Assignments
  TickerRowTotal = 2
  TotalVolume = 0
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  OpenBegPrice = ws.Cells(2, 3).Value
  
  'Input summary headers
  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Yearly Change"
  ws.Range("K1").Value = "Percent Change"
  ws.Range("L1").Value = "Total Stock Volume"

    'Sort ticker symbols and years to group them
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add2 Key:=ws.Range("A2:A" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ws.Sort.SortFields.Add2 Key:=ws.Range("B2:B" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange Range("A1:G" & LastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Loop through all the rows in a sheet
    For RowNumber = 2 To LastRow

        ' Check if the next ticker and year are the same
        If ws.Cells(RowNumber + 1, 1).Value <> ws.Cells(RowNumber, 1).Value Then
    
                ' Set the Ticker Symbol
                TickerSymbol = ws.Cells(RowNumber, 1).Value
                Debug.Print TickerSymbol
                
                'Add Closing Stock Price for End of Year
                CloseEndPrice = ws.Cells(RowNumber, 6).Value
                Debug.Print CloseEndPrice
                
                'Subtract the beginning and ending prices
                YearlyChg = CloseEndPrice - OpenBegPrice
                Debug.Print YearlyChg
                
                'Calculate the percentage (Ran into error with the numbers being zeroes)
                If YearlyChg <> 0 And OpenBegPrice <> 0 Then
                    YearlyChgPct = Format((YearlyChg / OpenBegPrice), "0.00%")
                    Debug.Print YearlyChgPct
                Else
                    YearlyChgPct = 0
                End If
                
                ' Add to the Total Stock Volume
                TotalVolume = TotalVolume + ws.Cells(RowNumber, 7).Value
                Debug.Print TotalVolume
                
                ' Put the ticker symbol in the summary column
                ws.Range("I" & TickerRowTotal).Value = TickerSymbol
                
                ' Put the Yearly Change in the summary column
                ws.Range("J" & TickerRowTotal).Value = YearlyChg
                
                'Add conditional formatting to the change
                If YearlyChg > 0 Then
                    ws.Range("J" & TickerRowTotal).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & TickerRowTotal).Interior.ColorIndex = 3
                End If
                
                ' Put the Yearly Change Percentage in the summary column
                ws.Range("K" & TickerRowTotal).Value = YearlyChgPct
                
                ' Put the Volume in the summary column
                ws.Range("L" & TickerRowTotal).Value = TotalVolume
                
                ' Add one to the summary table row
                TickerRowTotal = TickerRowTotal + 1
                
                ' Reset the Brand Total
                TotalVolume = 0
                
                'Set new Open price for beginning of year on next ticker
                OpenBegPrice = ws.Cells(RowNumber + 1, 3).Value
                Debug.Print OpenBegPrice

        ' If the cell immediately following a row is the same, then continue totaling the volume
        Else
    
          ' Add to the Brand Total
          TotalVolume = TotalVolume + ws.Cells(RowNumber, 7).Value
    
        End If
    
      Next RowNumber

    'Autofit Cells
    ws.Columns("A:L").AutoFit
    
    'Bold headers
    ws.Range("A1:L1").Font.Bold = True
    
Next ws

MsgBox "The macro has finished running.", , "Done"

End Sub

