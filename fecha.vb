Sub tabla()

    Dim i As Integer
    Dim num As Integer
    
    num = InputBox("ingrese un numero")
    
    For i = 1 To 10
    
        Cells(i, 1) = Str(num) & " x " & Str(i) & " = " & Str(num * i)
        
    Next

End Sub


Sub mostrar_fecha()
    
    Dim meses(1 To 12) As String
    Dim dia As Date

    meses(1) = "enero"
    meses(2) = "febrero"
    meses(3) = "marzo"
    meses(4) = "abril"
    meses(5) = "mayo"
    meses(6) = "junio"
    meses(7) = "julio"
    meses(8) = "agosto"
    meses(9) = "septiembre"
    meses(10) = "octubre"
    meses(11) = "noviembre"
    meses(12) = "diciembre"
    
    dia = InputBox("ingrese una fecha")
    
    MsgBox (Str(Day(dia)) & " de " & meses(Month(dia)) & " de " & Str(Year(dia)))
    

End Sub

Sub mostrar_recorrido()

    Range("a1:c4").Name = "conjunto"
    
    For Each celda In Range("conjunto")
        
        MsgBox (celda.Address)
        
    Next

End Sub

Sub escribir_nombre()

    Dim nombre As String
    Dim i As Integer
    
    nombre = InputBox("ingrese un nombre")
    
    Cells(1, 1).Select
    
    Do Until IsEmpty(ActiveCell)
    
        ActiveCell.Offset(1, 0).Select 'me muevo a la fila de abajo

    Loop
    
    ActiveCell = nombre
    
End Sub

Sub nombre_hojas()

    Dim nombre As String
    Dim cantidad As Integer
    Dim i As Integer
    
    nombre = InputBox("que nombre quiere poner a las hojas")
    
    cantidad = Sheets.Count
    
    For i = 1 To cantidad
    
        Sheets(i).Name = nombre & Str(i)
    
    Next
    
End Sub


