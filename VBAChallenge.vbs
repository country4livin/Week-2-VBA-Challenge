VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_challenge()
Dim Ticker As String
Dim QChange As Double
Dim Total As Variant
Dim PChange As Double
Dim LastRow As Long
Dim ws As Worksheet
Dim Opn As Double
Dim Cloz As Double
Dim i As Long
Dim SumTable As Integer
SumTable = 2
Set ws = ThisWorkbook.Sheets("Q1")

For Each ws In ThisWorkbook.Sheets
    With ws

Total = 0
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Opn = ws.Cells(2, 3).Value
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Quarterly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Volume"

For i = 2 To LastRow
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    Cloz = ws.Cells(i, 6).Value
    Ticker = ws.Cells(i, 1).Value
    QChange = Cloz - Opn
    PChange = QChange / Opn
    Total = Total + ws.Cells(i, 7).Value
    ws.Range("I" & SumTable).Value = Ticker
    ws.Range("J" & SumTable).Value = QChange
    ws.Range("K" & SumTable).Value = FormatPercent(PChange, 2)
    ws.Range("L" & SumTable).Value = Total
    SumTable = SumTable + 1
    Opn = ws.Cells(i + 1, 3).Value
    Total = 0

Else
   ws.Range("K" & SumTable).Value = Null
    Total = Total + ws.Cells(i, 7).Value
   If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    Total = 0
 End If
End If
Next i

For i = 2 To LastRow

If ws.Cells(i, 10) > 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 4
ElseIf ws.Cells(i, 10).Value < 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 3

End If


Next i
SumTable = 2

End With

Next ws

End Sub
