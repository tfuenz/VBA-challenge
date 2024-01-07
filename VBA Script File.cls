VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub CommandButton6_Click()
For Each ws In Worksheets
lastrow = ws.Cells(Rows.count, 9).End(xlUp).Row
For i = 2 To lastrow
For j = 9 To 12
ws.Cells(i, j).Clear
Next
Next
ws.Range("P2:Q4").Clear
Next
End Sub

Private Sub CommandButton7_Click()
For Each ws In Worksheets
'Declaring variables
lastrow1 = ws.Cells(Rows.count, 1).End(xlUp).Row
Dim count As Integer
Dim BOYOpen As Double
Dim EOYClose As Double
Dim YearlyChange As Double
'Total Stock Volume (TSV)
Dim TSV As LongLong
Dim DaysOfYear As Integer
Dim Day As Integer
'Greatest Percent Increase (GPI) & Greatest Percent Decrease (GPD)
Dim GPI As Double
Dim GPITicker As String
Dim GPD As Double
Dim GPDTicker As String
Dim GTSV As LongLong
Dim GTSVTicker As String
'We are going to initialize the first BOYOpen/Volume for consistency with code
BOYOpen = ws.Cells(2, 3).Value
count = 2
GPI = 0
GPITicker = ""
GPD = 0
GPDTicker = ""
GTSV = 0
GTSVTicker = ""
'Initialize certain fixed cells
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change ($)"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

'(1) Unique Ticker Symbols
For i = 2 To lastrow1
'Checking if the prior ticker is the same to determine new ticker
If ws.Cells(count - 1, 9).Value <> ws.Cells(i, 1).Value Then
ws.Cells(count, 9).Value = ws.Cells(i, 1).Value
count = count + 1
End If
'The dates are given as long #'s, so the first date will be the smallest
'and the last date will be the largest #
'(2) Yearly Change
If i > 2 And ws.Cells(i, 2).Value < ws.Cells(i - 1, 2).Value Then
'This signals i is the row with the first date
BOYOpen = ws.Cells(i, 3).Value
ElseIf ws.Cells(i, 2).Value > ws.Cells(i + 1, 2).Value Then
'This signals i is the row with the last date
EOYClose = ws.Cells(i, 6).Value
End If
If BOYOpen > 0 And EOYClose > 0 Then
YearlyChange = EOYClose - BOYOpen
ws.Cells(count - 1, 10).Value = YearlyChange
'(3) Percent Change
ws.Cells(count - 1, 11).Value = YearlyChange / BOYOpen
BOYOpen = 0
EOYClose = 0
End If
Next

'Counting unique tickers and formatting
lastrow2 = ws.Cells(Rows.count, 10).End(xlUp).Row
For i = 2 To lastrow2
ws.Cells(i, 11) = Format(ws.Cells(i, 11), "Percent")
If ws.Cells(i, 10) < 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 3
ElseIf ws.Cells(i, 10) > 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 4
End If
If ws.Cells(i, 11) < 0 Then
ws.Cells(i, 11).Interior.ColorIndex = 3
ElseIf ws.Cells(i, 11) > 0 Then
ws.Cells(i, 11).Interior.ColorIndex = 4
End If
Next

'(4) Total Stock Volume
DaysOfYear = (lastrow1 - 1) / (lastrow2 - 1)
count = 2
For i = 2 To lastrow1
TSV = TSV + ws.Cells(i, 7).Value
Day = Day + 1
If Day = DaysOfYear Then
ws.Cells(count, 12).Value = TSV
TSV = 0
Day = 0
count = count + 1
End If
Next

'(5) Script functionality
For i = 2 To lastrow2
If ws.Cells(i, 11).Value > GPI Then
GPI = ws.Cells(i, 11).Value
GPITicker = ws.Cells(i, 9).Value
ElseIf ws.Cells(i, 11).Value < GPD Then
GPD = ws.Cells(i, 11).Value
GPDTicker = ws.Cells(i, 9).Value
End If
If ws.Cells(i, 12).Value > GTSV Then
GTSV = ws.Cells(i, 12).Value
GTSVTicker = ws.Cells(i, 9).Value
End If
Next

'Assigning values to the final table
ws.Range("Q2").Value = GPI
ws.Range("Q3").Value = GPD
ws.Range("Q4").Value = GTSV
ws.Range("P2").Value = GPITicker
ws.Range("P3").Value = GPDTicker
ws.Range("P4").Value = GTSVTicker
ws.Range("Q2").Value = Format(ws.Range("Q2").Value, "Percent")
ws.Range("Q3").Value = Format(ws.Range("Q3").Value, "Percent")

Next
End Sub


