Sub primera()

    'MsgBox ("Bienvenidos al Sistema")          'COMENTARIO

    ActiveCell.Value = "Nicolás"                'ESCRIBE EN CELDA ACTIVA
    ActiveCell.Clear                            'BORRA CONTENIDO DE CELDA ACTIVA

End Sub

Sub segunda()

    Dim var As Integer                          'DECLARO VARIABLE DEL TIPO ENTERO
    Dim triple As Integer                       'DECLARO VARIABLE DEL TIPO ENTERO
    
    var = InputBox("Deme un número entero")     'TOMO VALOR DESDE TECLADO
    triple = var * 3                            'INICIALIZO VARIABLE
    ActiveCell = triple                         'VEO VALOR EN CELDA ACTIVA
    Cells(3, 2) = var                           'VEO VALOR EN CELDA 3B (B3)
    Cells(5, 2) = triple                        'VEO VALOR EN CELDA 5B (B5)
        
    ActiveCell.Offset(3, 2) = var               'ESCRIBO 3 CELDAS ABAJO Y 2 A LA DERECHA
    ActiveCell.Offset(-1, -1) = triple          'ESCRIBO 1 CELDA ARRIBA Y 1 A LA IZQUIERDA
    ActiveCell.Offset(0, 3) = var + triple      'PONGO VALOR 3 CELDAS A LA DERECHA
    ActiveCell.Offset(0, 0) = "LLEGUE"          'ESCRIBO "LLEGUE" EN CELDA ACTIVA
    
    Cells(9, 7).Select                          'ACTIVO CELDA 9G (G9)
               
    
End Sub

Sub ejercicio()

'Dar un número y escribir en la celda activa dicho número, en la celda
'que está a la derecha de ésta escribir una leyenda que indiqué si el número es par o impar.

Dim numero As Integer

numero = InputBox("Ingresar número")

ActiveCell = numero

If numero Mod 2 Then
    ActiveCell.Offset(0, 1) = "IMPAR"
    Else
    ActiveCell.Offset(0, 1) = "PAR"
End If



End Sub


Sub ejercicio_2()

'Dado un nombre si  el mismo termina con "a" escribirlo sin ella (sin la "a") pero concatenándole
'la terminación -ita. Si termina en "o" concaternarlo en -ito y sino no decir "que lindo nombre".

Dim nombre As String
Dim largo As Integer


nombre = InputBox("Ingresar nombre")
largo = Len(nombre)

If Right(nombre, 1) = "a" Then
    MsgBox (Left(nombre, largo - 1) + "ita")
Else
    If Right(nombre, 1) = "o" Then
    MsgBox (Left(nombre, largo - 1) + "ito")
    Else
        MsgBox ("Que lindo nombre")
    End If

End If
    
End Sub

    
Sub ejercicio_3()

'Armar una matriz identidad de tantas filas como indique el usuario teniendo en cuenta que ésta tiene que
'tener "1" en la principal y "0" en el resto, y es una matriz cuadrada.
'Mostrar esta matriz a partir de la celda que indique el usuario.


Dim orden As Integer
Dim fila As Integer
Dim columna As Integer
Dim f As Integer
Dim c As Integer

orden = InputBox("Indicar el orden de la matriz cuadrada")
fila = InputBox("Indicar fila de inicio de matriz")
columna = InputBox("Indicar columna de inicio de la matriz")

Cells(fila, columna).Select

For f = 0 To orden - 1
    For c = 0 To orden - 1
        If f = c Then
            ActiveCell.Offset(f, c) = 1
        Else
            ActiveCell.Offset(f, c) = 0
        End If
    Next
    
Next


End Sub


Sub ejercicio_4()

    'Pedirle al usuario un nombre y escribirlo 10 veces a partir de la celda A1 en diagonal.


    Dim nombre As String                            'DIM _ AS para declarar variable. String es el tipo
    Dim i As Integer                                'Integer es el tipo de variable.
Dim j As Integer


nombre = InputBox("Indicar nombre")             'InputBox inserto ventana

Cells(1, 1).Select                              'Me posiciono en la celda A1
    
For i = 0 To 9                                  'De 0 a 9 tengo las 10 veces
    For j = 0 To 9
        If i = j Then                           'Despues de la candición va Then
            ActiveCell.Offset(i, j) = nombre
        End If                                  'Siempre cierro el If con End If

    Next
Next

End Sub


