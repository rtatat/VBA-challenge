Sub Ticker()
   'Making sure the code will work in all worksheets
   For Each ws In Worksheets
   'Getting the Ticker column to return unique entries
    Dim Ticker As Range
    Set Ticker = ws.Range(ws.Range("A1"), ws.Range("A1").End(xlDown))
    Ticker.AdvancedFilter xlFilterCopy, , ws.Range("J1"), True
    'Titling the column
    ws.Cells(1, 10).Value = "Ticker"
   Next ws
End Sub

Sub YearlyChange()
    Defining variables
   Dim UniTicker As Range
   Set UniTicker = Range(Range("J2"), Range("J2").End(xlDown))
   Dim i As Object
   Dim k As Integer
   Dim FirstRow As Long
   Dim LastRow As Long
   Dim StartValue As Double
   Dim EndValue As Double
   'Making sure the code will work in all worksheets
   For Each ws In Worksheets
      For Each i In UniTicker
            'Reset all values to 0 at the start of each loop
            k = 0
            FirstRow = 0
            LastRow = 0
            StartValue = 0
            EndValue = 0
            'Find the first and last instance of a ticker, ie what row they appear in
            FirstRow = ws.Range("A:A").Find(what:=i, after:=ws.Range("A1"), LookAt:=xlWhole).Row
            LastRow = ws.Range("A:A").Find(what:=i, after:=ws.Range("A1"), LookAt:=xlWhole, searchdirection:=xlPrevious).Row
                'Getting the opening and closing value of the current ticker
                StartValue = ws.Cells(FirstRow, 3).Value
                EndValue = ws.Cells(LastRow, 6).Value
                'Setting k as the row that the current ticker is in for UniTicker
                k = ws.Range("J:J").Find(what:=i, after:=ws.Range("J1")).Row
                'Finding the yearly change and placing it into the cell
                ws.Cells(k, 11) = EndValue - StartValue
                'Cell formatting
                If ws.Cells(k, 11) >= 0 Then
                    ws.Cells(k, 11).Interior.ColorIndex = 4
                ElseIf ws.Cells(k, 11) < 0 Then
                    ws.Cells(k, 11).Interior.ColorIndex = 3
                End If
                'Titling the column
                ws.Cells(1, 11).Value = "Yearly Change"
        Next i
     Next ws
End Sub

Sub PercentChange()
    'Defining variables
    Dim UniTicker As Range
    Set UniTicker = Range(Range("J2"), Range("J2").End(xlDown))
    Dim i As Object
    Dim YearlyChange As Double
    Dim YearlyRow As Integer
    'Making sure the code will work in all worksheets
    For Each ws In Worksheets
        For Each i In UniTicker
            'Reset all values at the start of the loop
            StartValue = 0
            YearlyChange = 0
            FirstRow = 0
            'Find the opening value of a ticker
            FirstRow = ws.Range("A:A").Find(what:=i, after:=ws.Range("A1"), LookAt:=xlWhole).Row
            StartValue = ws.Cells(FirstRow, 3).Value
            'Find the yearly change
            YearlyRow = ws.Range("J:J").Find(what:=i, after:=ws.Range("J1"), LookAt:=xlWhole).Row
            YearlyChange = ws.Cells(YearlyRow, 11).Value
            'Finding the percent change and placing it into the cell
            ws.Cells(YearlyRow, 12) = YearlyChange / StartValue
            'Convert the returned values into a percentage
            ws.Range("L:L").NumberFormat = "0.00%"
            'Titling the column
            ws.Cells(1, 12).Value = "Percent Change"
        Next i
     Next ws
End Sub

Sub TotalVolume()
    'Defining variables
    Dim UniTicker As Range
    Set UniTicker = Range(Range("J2"), Range("J2").End(xlDown))
    Dim i As Object
    Dim FirstRow As Long
    Dim LastRow As Long
    Dim TotalRow As Long
    Dim VolRange As String
    Dim total As Long
    'Making sure the code will work in all worksheets
    For Each ws In Worksheets
        For Each i In UniTicker
            FirstRow = 0
            LastRow = 0
            TotalRow = 0
            TotVol = 0
            'Identifying the overall range of a ticker
            FirstRow = ws.Range("A:A").Find(what:=i, after:=ws.Range("A1"), LookAt:=xlWhole).Row
            LastRow = ws.Range("A:A").Find(what:=i, after:=ws.Range("A1"), LookAt:=xlWhole, searchdirection:=xlPrevious).Row
            VolRange = "ws.Cells(FirstRow, 7):ws.Cells(LastRow, 7)"
            'Finding the total volume for each ticker and placing it into the cell
            TotalRow = ws.Range("J:J").Find(what:=i, after:=ws.Range("J1"), LookAt:=xlWhole).Row
            ws.Cells(TotalRow, 13).Value = Application.Sum(Range(ws.Cells(FirstRow, 7), ws.Cells(LastRow, 7)))
            'Titling the column
            ws.Cells(1, 13) = "Total Stock Volume"
        Next i
    Next ws
End Sub

Sub Greatest()
    'Making sure the code will work in all worksheets
    For Each ws In Worksheets
        'Cell names
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        'Searching through the Percent Change column for the first two codes and the Total Stock Volume column for the last code
        ws.Cells(2, 17).Value = Application.WorksheetFunction.Max(ws.Range("L:L"))
        ws.Cells(3, 17).Value = Application.WorksheetFunction.Min(ws.Range("L:L"))
        ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(ws.Range("M:M"))
        'Using Xlookup to place the ticker names for the respected desired values
        ws.Range("P2") = Application.WorksheetFunction.XLookup(ws.Range("Q2"), ws.Range("L:L"), ws.Range("J:J"))
        ws.Range("P3") = Application.WorksheetFunction.XLookup(ws.Range("Q3"), ws.Range("L:L"), ws.Range("J:J"))
        ws.Range("P4") = Application.WorksheetFunction.XLookup(ws.Range("Q4"), ws.Range("M:M"), ws.Range("J:J"))
        'Formatting the cells to effectively display information
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "##0.00E+0"
    Next ws
End Sub
