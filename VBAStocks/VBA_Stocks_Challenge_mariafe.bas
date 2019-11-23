Attribute VB_Name = "Module1"
Sub Stock_Challenge()

    'Run on every worksheet, just by running the VBA script once
    Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
        
        'Set the Variables to its data types and initial values
        Dim openprice As Double
        Dim closeprice As Double
        Dim Ticker_Name As String
        Dim yearlychange As Double
        Dim Percent_Change As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Integer
        Row = 2
        
        'Set the heading to the summary
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        ' Determine the Last Row
        lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row
         'Assign value to the variable openprice
                openprice = Cells(2, 3).Value
        ' Loop on the all the ticker
        For i = 2 To lastrow
         ' Check if we are still within the same ticker symbol, if it is not...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ' Set Ticker name
                Ticker_Name = Cells(i, 1).Value
                Cells(Row, 9).Value = Ticker_Name
                ' Assign value to the variable closeprice
                closeprice = Cells(i, 6).Value
                ' Add Yearly Change
                yearlychange = closeprice - openprice
                'Indicate the position of the result
                Cells(Row, 10).Value = yearlychange
                ' Percent Change
                If (openprice = 0 And closeprice = 0) Then
                    Percent_Change = 0
                ElseIf (openprice = 0 And closeprice <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = yearlychange / openprice
                    Cells(Row, 11).Value = Percent_Change
                    Cells(Row, 11).NumberFormat = "0.00%"
                End If
                ' Add total volume
                Volume = Volume + Cells(i, 7).Value
                Cells(Row, 12).Value = Volume
                ' Add one to the summary table row
                Row = Row + 1
                ' reset the open price
                openprice = Cells(i + 1, 3)
                ' reset the total volume
                Volume = 0
            'if cells are the same ticker
            Else
                Volume = Volume + Cells(i, 7).Value
            End If
        Next i
        
        ' Set the Last Row of Yearly Change per WS
        lastrowyc = WS.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Indicate the color of the cells with the condition
        For j = 2 To lastrowyc
            If (Cells(j, 10).Value >= 0) Then
                Cells(j, 10).Interior.ColorIndex = 4
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
        
        ' For the Greatest % Decrease,% Increase, and Total Volume
        Cells(1, 15).Value = "Ticker"
        Cells(1, 15).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
        ' Find the greatest value through looping in each row
        For p = 2 To lastrowyc
            If Cells(p, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & lastrowyc)) Then
                Cells(2, 16).Value = Cells(p, 9).Value
                Cells(2, 17).Value = Cells(p, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"
            ElseIf Cells(p, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & lastrowyc)) Then
                Cells(3, 16).Value = Cells(p, 9).Value
                Cells(3, 17).Value = Cells(p, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            ElseIf Cells(p, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & lastrowyc)) Then
                Cells(4, 16).Value = Cells(p, 9).Value
                Cells(4, 17).Value = Cells(p, 12).Value
            End If
        Next p
        
    Next WS
        
End Sub
