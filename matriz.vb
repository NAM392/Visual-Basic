Sub calculando()
MsgBox ("El cursor se posicionara en C4")
Cells(4, 3).Select
num = InputBox("Introduce un número")
If num Mod 2 Then
MsgBox ("Es Impar")
Cells(3, 4) = "IMPAR"
Else
MsgBox ("Es Par")
Cells(3, 4) = "PAR"
End If
ActiveCell.Offset(-2, 1) = num
End Sub
Sub nombres_ito()
nombre = InputBox("Carga un nombre")
tam = Len(nombre)

End Sub

Sub matriz()
Dim mat As Integer
Dim fil As Integer
Dim col As Integer

mat = InputBox("Orden de la matriz cuadratica")
fil = InputBox("Introduce Número de fila")
col = InputBox("Introduce Numero de columma")
Cells(fil, col).Select
    For i = 1 To mat
        For j = 1 To mat
        If i = j Then
            ActiveCell.Offset(i - 1, j - 1) = 1
        Else
            ActiveCell.Offset(i - 1, j - 1) = 0
        End If
    Next
    Next

End Sub
