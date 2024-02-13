Attribute VB_Name = "Module6"
Sub resetData()
    Range("i2:i760000").Value = ""
    Range("j2:j760000").Value = ""
    Range("k2:k760000").Value = ""
    Range("l2:l760000").Value = ""
    
    Range("p2:p4").Value = ""
    Range("q2:q4").Value = ""
End Sub


Sub GenerateSummaryReport()

    'layout
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percentage Change"
    Cells(1, 12).Value = "Total Sock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest % Total Volume"
    
    Dim Volume_total As Double
    Dim ticker_record As Double
    Dim Summary_Table_row As Integer
    Dim year_change As Double
    Dim percentage_change As Double
    Dim earliestDateOpen As Double
    Dim latestDateClose As Double
    Dim lastRow As Long
    Dim brand_name As String

    Summary_Table_row = 2
    Volume_total = 0
    ticker_record = 0

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow - 1 ' Updated loop condition
        If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
            ticker_record = ticker_record + 1
        End If

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            brand_name = Cells(i, 1).Value
            Volume_total = Volume_total + Cells(i, 7).Value

            'last date
            latestDateClose = Cells(i, 6).Value
            'first date
            earliestDateOpen = Cells(i - ticker_record, 3).Value

            ' Calculate year_change
            year_change = latestDateClose - earliestDateOpen
            ' Calculate percentage_change
            percentage_change = year_change / earliestDateOpen

            Range("I" & Summary_Table_row).Value = brand_name
            Range("J" & Summary_Table_row).Value = Format(year_change, "0.00")
            Range("K" & Summary_Table_row).Value = percentage_change
            Range("L" & Summary_Table_row).Value = Volume_total
            Summary_Table_row = Summary_Table_row + 1
            Volume_total = 0
            ticker_record = 0
        Else
            Volume_total = Volume_total + Cells(i, 7).Value
        End If
    Next i
    
    ' Find the highest and lowest percentage change
    Dim maxPercentageChange As Double
    Dim minPercentageChange As Double
    Dim maxPercentageChangeTicker As String
    Dim minPercentageChangeTicker As String
    
    maxPercentageChange = Application.WorksheetFunction.Max(Range("K2:K" & Summary_Table_row - 1))
    minPercentageChange = Application.WorksheetFunction.Min(Range("K2:K" & Summary_Table_row - 1))
    
    maxPercentageChangeTicker = Cells(Application.WorksheetFunction.Match(maxPercentageChange, Range("K2:K" & Summary_Table_row - 1), 0) + 1, 9).Value
    minPercentageChangeTicker = Cells(Application.WorksheetFunction.Match(minPercentageChange, Range("K2:K" & Summary_Table_row - 1), 0) + 1, 9).Value
    
    ' Output to cells
    Range("P2").Value = maxPercentageChangeTicker
    Range("Q2").Value = maxPercentageChange
    Range("P3").Value = minPercentageChangeTicker
    Range("Q3").Value = minPercentageChange
    
    ' Find the highest volume
    Dim maxVolume As Double
    Dim maxVolumeTicker As String
    
    maxVolume = Application.WorksheetFunction.Max(Range("L2:L" & Summary_Table_row - 1))
    maxVolumeTicker = Cells(Application.WorksheetFunction.Match(maxVolume, Range("L2:L" & Summary_Table_row - 1), 0) + 1, 9).Value
    
    ' Output to cells
    Range("P4").Value = maxVolumeTicker
    Range("Q4").Value = maxVolume
   
End Sub
