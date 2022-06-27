Attribute VB_Name = "Module1"

Sub multiple():

Dim ws As Worksheet

For Each ws In Worksheets

'create cells:

    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Stock Volume"
'set up a variable of type Worksheet
Dim Sheet1, Sheet2, Sheet3 As Worksheet


'set the Sheet1 reference to the Sheet1 Worksheet
Set Sheet1 = Worksheets("2018")
Set Sheet2 = Worksheets("2019")
Set Sheet3 = Worksheets("2020")
'create cells in other sheets
Dim string1 As String
Dim string2 As String
Dim string3 As String
Dim string4 As String

string1 = "Ticker"
string2 = "Yearly Change"
string3 = "Percent Change"
string4 = "Total Stock Volume"

Sheet2.Range("J1").Value = string1
Sheet3.Range("J1").Value = string1

Sheet2.Range("K1").Value = string2
Sheet3.Range("K1").Value = string2

Sheet2.Range("L1").Value = string3
Sheet3.Range("L1").Value = string3

Sheet2.Range("M1").Value = string4
Sheet3.Range("M1").Value = string4

'autofit columns
Worksheets("2018").Range("J1:M1").Columns.AutoFit
Worksheets("2019").Range("J1:M1").Columns.AutoFit
Worksheets("2020").Range("J1:M1").Columns.AutoFit


'variables
    Ticker = ""
    totalStocks = 0
    summaryTableRow = 2
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    openPrice = ws.Cells(2, 3).Value

'functions
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastRow

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            Ticker = ws.Cells(i, 1).Value

            closePrice = ws.Cells(i, 6).Value

            yearlyChange = closePrice - openPrice

            percentChange = (closePrice - openPrice) / openPrice

            totalStocks = totalStocks + ws.Cells(i, 7).Value


            ws.Cells(summaryTableRow, 10).Value = Ticker

            ws.Cells(summaryTableRow, 13).Value = totalStocks

            ws.Cells(summaryTableRow, 11).Value = yearlyChange
            ws.Cells(summaryTableRow, 12).Value = percentChange


            summaryTableRow = summaryTableRow + 1

            totalStocks = 0

            openPrice = ws.Cells(i + 1, 3).Value

        Else
            totalStocks = totalStocks + ws.Cells(i, 7).Value
        End If

            If ws.Cells(summaryTableRow, 11).Value < 0 Then
               ws.Cells(summaryTableRow, 11).Interior.ColorIndex = 3

            ElseIf ws.Cells(summaryTableRow, 11).Value > 0 Then
                 ws.Cells(summaryTableRow, 11).Interior.ColorIndex = 4

        End If
    
    Next i

Next ws

End Sub
