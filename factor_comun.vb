Sub recursivo_2()

Dim NUM As Integer

Dim NUM2 As Integer

    NUM = InputBox("ingrese primer numero")

NUM2 = InputBox("ingrese segundo numero")

MsgBox ("el factor comun es " & MCD(NUM, NUM2))

End Sub

Function MCD(a As Integer, b As Integer) As Long

    Dim i As Integer
Dim C As Integer
Dim D As Integer
Dim cont As Integer

If a <> 0 & b <> 0 Then
    If a > b Then
        If b = 1 Then
            MCD = C
        Else
            
            C = b
            D = a Mod b
            If D <> 0 Then
                MCD = MCD(C, D)
            Else
                MCD = C
            End If
        End If
    Else
        If b > a Then
        If a = 1 Then
            MCD = C
        Else
            
            C = a
            D = b Mod a
            If D <> 0 Then
                MCD = MCD(C, D)
            Else
                MCD = C
            End If
        End If
    End If
End If

End If

End Function
