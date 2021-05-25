'fibonacci

Sub punto_1()


 Dim ingreso As Integer
 
ingreso = InputBox("indique posicion de numero fibonacci")

MsgBox ("en posicion " & ingreso & "el numero fibonacci es " & fibonacci(ingreso))


End Sub

Function fibonacci(NUM As Integer) As Integer
If NUM = 0 Then
    FIBO = 0
Else
    If NUM = 1 Then
        fibonacci = 1
    Else
        fibonacci = fibonacci(n - 1) + fibonacci(n - 2)
    End If
End If





End Function


Sub binario()



    Dim ingreso As String
    Dim NUM As Long
    NUM = 0
    Dim j As Integer
    j = 1
    Dim BIN As Long
    
    ingreso = InputBox("ingrese un numero binario")
    For i = 0 To Len(ingreso) - 1

    NUM = (Val(Mid(ingreso, Len(ingreso) - i, 1)))
    BIN = BIN + NUM * j
    j = j + j
    
    Next


   MsgBox ("en decimales es : " & BIN)
        






































End Sub
