VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub market()
Dim ticker As String
Dim numticker As Long
Dim ychange As Long

Cells(1, 9) = "Ticker"
Cells(1, 10) = "Year change"
Cells(1, 11) = "percent change"
Cells(1, 12) = "Total Volume"


lastrow = Cells(Rows.Count, 1).End(xlUp).Row
ticker = 2
numticker = 2
    For i = 2 To lastrow
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker = Cells(i, 1).Value
        Range("I" & numticker).Value = ticker
        Cells(numticker, 10).Value = Cells(i, 6) - Cells(numticker, 3)
        Cells(numticker, 11).Value = Cells(numticker, 10).Value / Cells(numticker, 3)
        Cells(numticker, 12).Value = WorksheetFunction.Sum(Range(Cells(i, 7), Cells(numticker, 7)))
        If Cells(numticker, 10).Value < 0 Then
        Cells(numticker, 10).Interior.ColorIndex = 3
        Else
        Cells(numticker, 10).Interior.ColorIndex = 4
        End If
        numticker = numticker + 1
    End If
    Next i
End Sub

