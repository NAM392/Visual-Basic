Sub principal()
Dim n As Integer
Dim p As Integer

n = InputBox("Deme un numero")
p = InputBox("Deme un numero")
MsgBox ("resto =" & n - p * di(n, p) & " - división = " & di(n, p))

End Sub
Sub mas()

For i = 1 To 100
    Cells(i, 3).Interior.Color = RGB(24, i, 75)
Next
ActiveCell.Font.Color = vbMagenta
End Sub
Sub forma()
Cells(1, 1).Select
For i = 0 To 56
    ActiveCell.Offset(i, 0).Interior.ColorIndex = i
Next
End Sub
Function fac(a As Integer) As Long
If a = 1 Then
    fac = 1
Else
    fac = a * fac(a - 1)
End If
End Function
Function di(a As Integer, b As Integer) As Integer
If a < b Then
    di = 0
Else
    di = 1 + di(a - b, b)
End If
End Function
