Attribute VB_Name = "Module1"
Sub VBA_challenge()

'Set initial values

Dim Ticker_Name As String
Dim Trade_Volume As Double
Trade_Volume = 0
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
Dim Open_value As Double
Dim Close_value As Double
Dim Report_date As Date
Dim Percentage_change As Double
Dim Percentage As String
Dim Annual_change As Double

Cells(1, 10).Value = ("Ticker")
Cells(1, 11).Value = ("Annual Change")
Cells(1, 12).Value = ("Percentage Change")
Cells(1, 13).Value = ("Trade Volume")

'Start Loop

For i = 2 To 800000
If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
Open_value = Cells(i, 3).Value

'Check Ticker Value. If different add new value. Add Trade volume to total.
ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
Ticker_Name = Cells(i, 1).Value

Close_value = Cells(i, 6).Value
Trade_Volume = Trade_Volume + Cells(i, 7).Value

'Calculate difference & Percentage

Total_Change = (Close_value - Open_value)
Percentage_change = ((Total_Change / Open_value) * 100)
'Percentage = FormatPercent(Percentage_change, , , vbTrue)

'Print in summary table
Range("J" & Summary_Table_Row).Value = Ticker_Name

Range("K" & Summary_Table_Row).Value = Total_Change
If Total_Change >= 0 Then Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
If Total_Change < 0 Then Range("K" & Summary_Table_Row).Interior.ColorIndex = 3

Range("L" & Summary_Table_Row).Value = (Percentage_change)

Range("M" & Summary_Table_Row).Value = Trade_Volume

'Next Row

Summary_Table_Row = Summary_Table_Row + 1
        
 ' Reset
Trade_Volume = 0

'Add to Total
Else
Trade_Volume = Trade_Volume + Cells(i, 7).Value

End If

Next i

End Sub

