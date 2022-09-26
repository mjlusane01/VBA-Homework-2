Attribute VB_Name = "Module1"
Option Explicit
Sub tickerData()
Dim ws As Worksheet
Dim i As Double
Dim j As Double
Dim ticker As String
Dim open1 As Double
Dim volume As Double
Dim close1 As Double
Dim yearlychange As Double
Dim percentchange As Double
Dim lastrow As Double
Dim lastrow2 As Double

For Each ws In Worksheets
ws.Activate
i = 2

lastrow = Cells(Rows.Count, "A").End(xlUp).Row

Cells(1, "H").Value = "Ticker"
Cells(1, "I").Value = "Yearly Change"
Cells(1, "J").Value = "Percent Change"
Cells(1, "K").Value = "Total Stock Volume"

For i = 2 To lastrow

'First instance of Ticker
If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
open1 = Format(Cells(i, 3).Value, "#.00")
startrow = Cells(i, 1).Row

ticker = Cells(i, 1).Value
Cells(Rows.Count, "H").End(xlUp).Offset(1, 0).Activate
ActiveCell.Value = ticker
End If

'Last instance of Ticker
If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
endrow = Cells(i, 1).Row

close1 = Format(Cells(i, 6).Value, "#.00")
volume = Application.WorksheetFunction.Sum(Range(Cells(startrow, 7), Cells(endrow, 7)))

yearlychange = close1 - open1
percentchange = (close1 - open1) / open1

Cells(Rows.Count, "I").End(xlUp).Offset(1, 0).Activate
ActiveCell.Value = Format(yearlychange, "#.00")


Cells(Rows.Count, "J").End(xlUp).Offset(1, 0).Activate
ActiveCell.Value = Format(percentchange, "0.00%")


Cells(Rows.Count, "K").End(xlUp).Offset(1, 0).Activate
ActiveCell.Value = volume

End If

Next i

Range("H:K").EntireColumn.AutoFit

lastrow2 = Cells(Rows.Count, "I").End(xlUp).Row

For j = 2 To lastrow2
If Cells(j, "I").Value < 0 Then
Range("I" & j).Interior.Color = vbRed
Else
Range("I" & j).Interior.Color = vbGreen
End If
Next j

Next ws


End Sub
