VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Wallstreeticker()

Dim ws As Worksheet
'set an initial variable for The ticker symbol
Dim Ticker As String
'set an initial variable for the Opening Price
Dim Openingprice As Double
'set an initial variable for the Closing Price
Dim Closingprice As Double
'Set an initial variable for the Yearly Price Change
Dim YearlyChange As Double
'Set an initial variable for Total Volume
Dim TotalStockVolume As Double
'set an initial variable for the Percent Change at YE
Dim PercentChange As Double
Dim lastrow As Long
Dim J As Long
Dim i As Long
Dim First As Double

For Each ws In Worksheets
J = 1
First = 2
YearlyChange = 0
TotalStockVolume = 0

'Set Column Titles
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"


lastrow = ws.Cells(Row.Count, 1).End(xlUp).Row


'Loop through all stocks for one year
For i = 2 To lastrow


    If i = First Then
    
    J = J + 1

'set the ticker
Ticker = ws.Cells(i, 1).Value
'set the Opening Price
Openingprice = Cells(i, 3).Value
'Set the volume
TotalStockVolume = Cells(i, 7).Value
'set end ticker value
ws.Cells(J, 9).Value = Ticker


ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

    First = i + 1

TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
Closingprice = ws.Cells(i, 6).Value
YearlyChange = (Closingprice - Openingprice)

If Openingprice = 0 Then
    PercentChange = 0
    
Else
    PercentChange = (YearlyChange / Openingprice) * 100
    PercentChange = Round(PercentChange, 2)

End If

If PercentChange > 0 Then
    ws.Cells(J, 10).Interior.ColorIndex = 4
    ws.Cells(J, 11).Interior.ColorIndex = 4
Else
    ws.Cells(J, 10).Interior.ColorIndex = 3
    ws.Cells(J, 11).Interior.ColorIndex = 3



End Sub



















