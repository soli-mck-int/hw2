Sub hw2_try()

For Each ws In Worksheets
Dim ticker As String
Dim vol As Double
Dim V_open As Double
Dim V_close As Double

j = 2
lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow
If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
V_open = ws.Cells(i + 1, 3).Value
ws.Cells(j + 1, 14).Value = V_open
V_close = ws.Cells(i, 6).Value
ws.Cells(j, 15).Value = V_close
ticker = ws.Cells(i, 1)
ws.Cells(j, 9) = ticker
vol = vol + ws.Cells(i, 7).Value
ws.Cells(j, 12).Value = vol
j = j + 1
vol = 0

Else
vol = vol + ws.Cells(i, 7).Value
End If
Next i
Next ws
End Sub

Sub part2()
For Each ws In Worksheets

new_lastrow = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
For c = 2 To new_lastrow
ws.Cells(c, 10).Value = ws.Cells(c, 15).Value - ws.Cells(c, 14).Value
ws.Cells(c, 11).Value = (ws.Cells(c, 10).Value / ws.Cells(c, 14).Value) * 100
If ws.Cells(c, 10).Value >= 0 Then
ws.Cells(c, 10).Interior.ColorIndex = 4
Else
ws.Cells(c, 10).Interior.ColorIndex = 3
End If
Next c
Next ws
End Sub
