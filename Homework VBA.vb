'EASY, MODERATE, HARD and CHALLENGE exercises from the homework are contained in this code
'For the CHALLENGE section, the code calculates the totals required in a new sheet called Totals that inserts at the end
Sub total_calc()

'Loop to execute the script in all sheets of the workbook
For Each WS In Worksheets

If WS.Name <> "Totals" Then

'Variable declarations
    Dim i As Integer
    Dim count_ticker As Integer
    Dim vol As LongLong
    Dim opener, closing, year_ch, percent_ch As Double
    Dim rg As Range
    Dim cond1, cond2 As FormatCondition

'Variable initializations
    vol = 0
    count_ticker = 1

' Determine the Last Row of data of each sheet
    lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row

'Adjusting headers
    WS.Range("I1").Value = "Ticker"
    WS.Range("J1").Value = "Yearly Change"
    WS.Range("K1").Value = "Percent Change"
    WS.Range("L1").Value = "Total Stock Volume"

'Loop for calculations
    For m = 2 To lastrow
        Ticker = WS.Cells(m, 1).Value

    'Obtain the opener value of the year
        fecha = Right(WS.Cells(m, 2).Value, 4)
        If fecha = "0101" Then
            opener = WS.Cells(m, 3).Value
        End If
        
        If m > 2 Then
'Checking, if Ticker is the same, add volume
            If Ticker = WS.Cells(m + 1, 1).Value Then
                vol = vol + WS.Cells(m, 7).Value
            Else
'If not, calculate totals
                vol = vol + WS.Cells(m, 7).Value
                count_ticker = count_ticker + 1
                WS.Cells(count_ticker, 9) = Ticker      'Ticker
                WS.Cells(count_ticker, 12) = vol        'Total Stock Volume
                vol = 0
'Obtain the closing value of the year
                fecha = Right(WS.Cells(m, 2).Value, 4)
                If fecha = "1230" Or fecha = "1231" Then
                    closing = WS.Cells(m, 6).Value
                End If
'Calculate the Yearly Change
                year_ch = closing - opener
                WS.Cells(count_ticker, 10).Value = year_ch
'Calculate the Percent Change
                If opener <> 0 Then
                    percent_ch = year_ch / opener
                    WS.Cells(count_ticker, 11).Value = percent_ch
                Else
                    percent_ch = 0
                End If
            End If

        Else
        vol = vol + WS.Cells(m, 7).Value
        End If
        
    Next m

'Percentage format
WS.Range("K2:K" & count_ticker).NumberFormat = "0.00%"

'Cells Conditional Format of Yearly Change
Set rg = WS.Range("J2", WS.Range("J2").End(xlDown))
 
'Clear any existing conditional formatting
rg.FormatConditions.Delete
 
'Define the rule for each conditional format
Set cond1 = rg.FormatConditions.Add(xlCellValue, xlGreater, "0")
Set cond2 = rg.FormatConditions.Add(xlCellValue, xlLess, "0")
 
'Define the format applied for each conditional format
With cond1
.Interior.Color = vbGreen
End With
 
With cond2
.Interior.Color = vbRed
End With
 
End If

Next WS

'Create a sheet for the totals of the HARD exercise
For Each WS In Worksheets
    If WS.Name = "Totals" Then
        Application.DisplayAlerts = False
        Sheets("Totals").Delete
        Application.DisplayAlerts = True
    End If
Next WS

Sheets.Add After:=Worksheets(Worksheets.Count)
Sheets(Worksheets.Count).Name = "Totals"

'Copy all the results to Total Worksheet
Sheets("Totals").Activate

For Each WS In Worksheets
lastrow_f = WS.Cells(Rows.Count, 10).End(xlUp).Row
If WS.Name <> "Totals" Then
    WS.Range("i2:L" & lastrow_f).Copy
    ActiveSheet.Paste Range("A1048576").End(xlUp).Offset(1, 0)
End If
Next WS

'Create the format for final results
Range("G3").Value = "Ticker"
Range("H3").Value = "Value"
Range("F4").Value = "Greatest % Increase"
Range("F5").Value = "Greatest % Decrease"
Range("F6").Value = "Greatest Total Volume"

'Locate results max, min and volume
lastrow_f = Cells(Rows.Count, 1).End(xlUp).Row

Max = Application.WorksheetFunction.Max(Columns("C"))
Min = Application.WorksheetFunction.Min(Columns("C"))
Max_V = Application.WorksheetFunction.Max(Columns("D"))

For i = 1 To lastrow_f

    If Cells(i, 3).Value = Max Then

        Range("g4").Value = Cells(i, 1).Value
        Range("h4").Value = Cells(i, 3).Value

    ElseIf Cells(i, 3).Value = Min Then

        Range("g5").Value = Cells(i, 1).Value
        Range("h5").Value = Cells(i, 3).Value

    ElseIf Cells(i, 4).Value = Max_V Then

        Range("g6").Value = Cells(i, 1).Value
        Range("h6").Value = Cells(i, 4).Value
    
    End If

Next i

'Adjust format
Columns("F:F").EntireColumn.AutoFit
Columns("G:G").EntireColumn.AutoFit
Columns("H:H").EntireColumn.AutoFit
Range("G3").Font.Bold = True
Range("H3").Font.Bold = True
Range("H4").NumberFormat = "0.00%"
Range("H5").NumberFormat = "0.00%"

'Delete the info on the totals worksheet
Range("A1:D1048576").ClearContents

End Sub
