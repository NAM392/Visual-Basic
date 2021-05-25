Sub potencia_recursiva()

Dim NUM As Integer
Dim NUM2 As Integer



NUM = InputBox("ingrese numero")

NUM2 = InputBox("ingrese potencia")

MsgBox (potencia(NUM, NUM2))



End Sub

Function potencia(a As Integer, b As Integer) As Long


    If b = 1 Then
        potencia = a
    Else
        potencia = a * potencia(a, b - 1)
    End If



End Function
