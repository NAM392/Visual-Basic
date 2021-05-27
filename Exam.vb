Sub parcial_1A()


Dim texto As String

Sheets("TEXTOS").Select
texto = Cells(1, 1)
Sheets("PARRAFO").Select
Cells(2, 1) = UCase(Left(texto, 1)) & Mid(texto, 2, Len(texto))   'copia la primera letra en mayiscula


For i = 2 To 10
    Sheets("TEXTOS").Select
    texto = Cells(i, 1)
    Sheets("PARRAFO").Select
    Cells(2, i) = texto

Next


'Cells(3, 3) = UCase(Left(texto, 1))  'copia la primera letra en mayiscula


End Sub

Sub parcial_1B()

Dim nota As Integer
Dim NOMBRE As String
Dim i As Integer
Dim SUMA As Integer


i = 4
Sheets("ALUMNOS").Select
NOMBRE = Cells(i, 2)

Do While (Cells(i, 2) <> "")
    
    For j = 3 To 5
        nota = InputBox("ingresa nota de" & " " & NOMBRE)
        Cells(i, j) = nota
        SUMA = SUMA + nota
    Next
    If SUMA / 3 > 4 Then
        Cells(i, 7) = SUMA / 3
    Else:
        Cells(i, 7) = 2
    End If
    i = i + 1
    NOMBRE = Cells(i, 2)
    SUMA = 0
    

Loop





End Sub

Sub parcial_1c()





    '    Dim i As Integer
    '
    '    Dim Conta As Integer
    '    Dim CARRERA As String
    '    i = 0
    '
    '    Cells(4, 1).Select
    '
    '    Do Until IsEmpty(ActiveCell)
    '
    '        CARRERA = Cells(4 + i, 1)
    '        Cells(4, 1).Select
    '        Do Until IsEmpty(ActiveCell)
    '            If (CARRERA = Cells(4 + Conta, 1)) Then
    '                cantidad = cantidad + 1
    '
    '            End If
    '            Sheets("ESTADISTICA").Select
    '            Cells(2 + i, 2) = CARRERA
    '            Cells(2 + i, 2 + Conta) = cantidad
    '            Conta = Conta + 1
    '            MsgBox (CARRERA)
    '            ActiveCell.Offset(1, 0).Select
    '
    '         Loop
    '
    '        i = i + 1
    '        ActiveCell.Offset(1, 0).Select
    '
    '
    '    Loop
    '



End Sub


Sub PROBANDO()

    Dim meses(1 To 12) As String
    Dim DIA As Date

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
    
    DIA = InputBox("ingrese una fecha")
    
    MsgBox (Str(Day(DIA)) & " de " & meses(Month(DIA)) & " de " & Str(Year(DIA)))




End Sub

Sub PROB()





MsgBox (Sheets("rangos").Cells(2, 1).Value)

End Sub

Sub convertir_moneda()

'convertir moneda usando formulario


    Dim flag As Integer
    
    Select Case (Me!inicial.Value & Me!Final.Value)
        Case 12, 13, 14, 21, 23, 24, 31, 32, 34, 41, 42, 43
            flag = 0
        Case Else
            flag = 1
    End Select
    
        If flag = 1 Then
            MsgBox ("La elecci?n de monedas es inv?lida")
        
        Else
        
            Select Case Me!inicial.Value
                Case 1
                    Select Case Me!Final.Value
                        Case 2
                            Me!imp_res.Value = Me!cantidad * 0.91
                            Me!mone_res.Value = "Euros"
                        Case 3
                            Me!imp_res.Value = Me!cantidad * 109.25
                            Me!mone_res.Value = "Nuevos soles"
                        Case 4
                            Me!imp_res.Value = Me!cantidad * 3.36
                            Me!mone_res.Value = "Yenes"
                    End Select
                        
                Case 2
                    Select Case Me!Final.Value
                        Case 1
                            Me!imp_res.Value = Me!cantidad / 0.91
                            Me!mone_res.Value = "Dolares"
                        Case 3
                            Me!imp_res.Value = (Me!cantidad / 0.91) * 109.25
                            Me!mone_res.Value = "Nuevos soles"
                        Case 4
                            Me!imp_res.Value = (Me!cantidad / 0.91) * 3.36
                            Me!mone_res.Value = "Yenes"
                        
                    End Select
                                
                Case 3
                    Select Case Me!Final.Value
                        Case 1
                            Me!imp_res.Value = Me!cantidad / 109.25
                            Me!mone_res.Value = "Dolares"
                        Case 2
                            Me!imp_res.Value = (Me!cantidad / 109.25) * 0.91
                            Me!mone_res.Value = "Euros"
                        Case 4
                            Me!imp_res.Value = (Me!cantidad / 109.25) * 3.36
                            Me!mone_res.Value = "Yenes"
                    End Select
                
                Case 4
                    Select Case Me!Final.Value
                        Case 1
                            Me!imp_res.Value = Me!cantidad / 3.36
                            Me!mone_res.Value = "Dolares"
                        Case 2
                            Me!imp_res.Value = (Me!cantidad / 3.36) * 0.91
                            Me!mone_res.Value = "Euros"
                        Case 3
                            Me!imp_res.Value = (Me!cantidad / 3.36) * 109.25
                            Me!mone_res.Value = "Nuevos soles"
                    End Select
                                    
            End Select
            
            
                
        End If
        
            
            
End Sub


Private Sub convertir_BTN_Click()               'algoritmo  cambio de moneda


    If (Me!inicial = Me!a_convertir) Then

        MsgBox("Las monedas seleccionadas deben ser distintas entre si")

    End If

    If (Me!inicial > 4 Or Me!inicial < 1) Then

        MsgBox("El valor ingresado en 'Moneda Inicial' no es valido")

    End If

    If (Me!a_convertir > 4 Or Me!a_convertir < 1) Then

        MsgBox("El valor ingresado en 'Moneda A Convertir' no es valido")

    End If


    '-----------------------------------------------------

    If (Me!inicial = 1) And (Me!a_convertir = 2) Then

        Me!cant_convertida = Me!cantidad * 0.91
        Me!MONEDA = "Euros"
        Me!moneda_inicial = "Dolares N.A."

    End If

    If (Me!inicial = 1) And (Me!a_convertir = 3) Then

        Me!cant_convertida = Me!cantidad * 109.25
        Me!MONEDA = "Yenes"
        Me!moneda_inicial = "Dolares N.A."

    End If

    If (Me!inicial = 1) And (Me!a_convertir = 4) Then

        Me!cant_convertida = Me!cantidad * 3.36
        Me!MONEDA = "Nuevos Soles"
        Me!moneda_inicial = "Dolares N.A."

    End If

    '-----------------------------------------------------

    If (Me!inicial = 2) And (Me!a_convertir = 1) Then

        Me!cant_convertida = Me!cantidad * 1.1
        Me!MONEDA = "D?lares N.A."
        Me!moneda_inicial = "Euros"

    End If

    If (Me!inicial = 2) And (Me!a_convertir = 3) Then

        Me!cant_convertida = Me!cantidad * 120.39
        Me!MONEDA = "Yenes"
        Me!moneda_inicial = "Euros"

    End If

    If (Me!inicial = 2) And (Me!a_convertir = 4) Then

        Me!cant_convertida = Me!cantidad * 3.7
        Me!MONEDA = "Nuevos Soles"
        Me!moneda_inicial = "Euros"

    End If

    '-------------------------------------------------------

    If (Me!inicial = 3) And (Me!a_convertir = 1) Then

        Me!cant_convertida = Me!cantidad * 0.0092
        Me!MONEDA = "Dolares N.A."
        Me!moneda_inicial = "Yenes"

    End If

    If (Me!inicial = 3) And (Me!a_convertir = 2) Then

        Me!cant_convertida = Me!cantidad * 0.0083
        Me!MONEDA = "Euros"
        Me!moneda_inicial = "Yenes"

    End If

    If (Me!inicial = 3) And (Me!a_convertir = 4) Then

        Me!cant_convertida = Me!cantidad * 0.031
        Me!MONEDA = "Nuevos Soles"
        Me!moneda_inicial = "Yenes"

    End If

    '-------------------------------------------------------

    If (Me!inicial = 4) And (Me!a_convertir = 1) Then

        Me!cant_convertida = Me!cantidad * 0.3
        Me!MONEDA = "D?lares N.A."
        Me!moneda_inicial = "Nuevos Soles"

    End If

    If (Me!inicial = 4) And (Me!a_convertir = 2) Then

        Me!cant_convertida = Me!cantidad * 0.27
        Me!MONEDA = "Euros"
        Me!moneda_inicial = "Nuevos Soles"

    End If

    If (Me!inicial = 4) And (Me!a_convertir = 3) Then

        Me!cant_convertida = Me!cantidad * 32.7
        Me!MONEDA = "Yenes"
        Me!moneda_inicial = "Nuevos Soles"

    End If

    '-------------------------------------------------------

End Sub

Sub probando_lala()
Dim CUIT As String
Dim msp As Long
Dim msd As Long
Dim prod As Long
Dim j As Integer
Dim k As Integer
j = 2
k = 5

                                'parte del algoritmo del digito verificador de cuit
CUIT = InputBox("cuit")

    
  For i = 0 To 5
   
    msp = msp + Val(Mid(CUIT, Len(CUIT) - i, 1)) * j
    j = j + 1
  Next
  
  For i = 1 To 4
  
    msd = msd + Val(Mid(CUIT, i, 1)) * k
    k = k - 1
  Next
  
  
  prod = msp + msd
  
   MsgBox (prod Mod 11)
  
    

End Sub


Sub primera_clase()



Dim ingreso As Integer
Dim cantidad As Integer
Dim flag As Byte
Dim NOMBRE As String


cantidad = Sheets.Count

For f = 1 To cantidad
    If Sheets(f).Name = "PRUEBA" Then flag = 1

 Next
 
 If flag = 1 Then
    Sheets("PRUEBA").Select                         'me posiciona en la hoja prueba si no existe la crea y me posiciona
    MsgBox ("me posiciono en PRUEBA")
 Else
     Sheets.Add after:=Sheets(cantidad)
     Sheets(Sheets.Count).Name = "PRUEBA"
     MsgBox ("creo PRUEBA")
 End If
 
    
    
    NOMBRE = InputBox("ingrese un nombre")
    
    Cells(1, 1).Select
    
    For i = 1 To 10
      ActiveCell.Offset(i, i) = NOMBRE       'escribe en diagonal la palabra que se le ingresa
    Next
    
    
    MsgBox ("limpiar")
    Sheets("PRUEBA").Cells.ClearContents        'borra todo el contenido de las celdas de la hoja
    
    Sheets("PRUEBA").Select
    For Each s In Range("d1:f1")               'escribe una palabra ingresada dentro de el rango d1:f1 (tambien se le puede asignar el nombre de algun rango)
       
       ingreso = InputBox("ingrese sueldos")
       
       s.Value = ingreso
    
    Next
       
    
   
End Sub


Sub contraena()

    'ingresa un usuario y contraseña si este no es valido al 3er intento se cierra excel

    Dim ingreso As String
Dim pass As String
Dim ch As Byte
Dim intento As Integer
Dim flag As Byte
flag = 0
ch = 0
intento = 0

Do While (intento < 3)
    ingreso = InputBox("ingrese usuario")
    
    Sheets("USUARIOS").Select
    For Each s In Range("user")
            If (s = ingreso) Then
                flag = 1
                s.Select
            End If
    Next
        If (flag = 1) Then
            ch = 1
        Else:
            intento = intento + 1
            MsgBox ("usuario incorrecto intento : " & intento)
        End If
      
    
    If (ch = 1) Then
        pass = InputBox("ingrese contraseña")
        If (pass = ActiveCell.Offset(0, 1)) Then
            MsgBox ("ingresaste")
            Sheets("PRUEBA").Select
            intento = 5
        Else
            intento = intento + 1
            MsgBox ("contraseña incorrecto intento : " & intento)
        End If
    End If
Loop

If (intento = 3) Then
    
    For Each w In Application.Workbooks
     w.Save
    Next w
    Application.Quit

End If



End Sub

'verificar si el numero ingresado es binario
Sub binario()

Dim ingreso As String
Dim flag As Byte
Dim no_es As Byte                       'cuando reviso entre dos caracteres unicamente el "ELSE" es clave para determinar

no_es = 0
flag = 0
ingreso = InputBox("ingrese un numero binario")


For i = 1 To Len(ingreso)
    
   If ((Mid(ingreso, i, 1)) = 1) Or ((Mid(ingreso, i, 1)) = 0) Then
        flag = 2
   Else
        no_es = 1
   End If
 Next
 
    If (no_es = 0) Then
        MsgBox ("el numero es binario")
    Else
        MsgBox ("el numero no es binario")
    End If
    
    flag = 0




End Sub





Sub forma()
Sheets("Colores").Select

Cells(1, 1).Select
For i = 0 To 56
    ActiveCell.Offset(i, 0).Interior.ColorIndex = i  'recorre del 1 al 56 los colores vba
Next
End Sub

Sub mas()
    Sheets("Colores").Select
    ActiveCell = Cells(3, 6)
    ActiveCell.Font.Color = vbMagenta  'pinta la celda activa de color magenta
End Sub

Sub rango_colores()
Dim i As Integer



For Each c In Range("r_colores")
    
        c.Interior.ColorIndex = i  'recorre del 1 al 56 los colores vba
        i = i + 1
Next

End Sub
Sub creando_nueva_hoja()


Dim ingreso As String
Dim cantidad As Integer
Dim flag As Byte
    ingreso = InputBox("ingrese el nombre de la hoja")
    cantidad = Sheets.Count
    
    For i = 1 To cantidad
        If (Sheets(i).Name = ingreso) Then flag = 1
    Next
        
         If (flag = 1) Then
            Sheets(i).Select
         Else
            Sheets.Add after:=Sheets(cantidad)
            Sheets(Sheets.Count).Name = ingreso
            Sheets(ingreso).Select
            
         End If
    

End Sub

Sub copiar()

Dim librodatos As Workbook

Set librodatos = Workbooks.Open("C:\Users\nam-1\OneDrive\Documentos\nombres1.xlsm")

librodatos.Sheets("nombres").Range("N_1").Copy
librodatos.Close savechanges:=False
Sheets("DATOS").Select
ActiveSheet.Paste
ActiveCell.Select

End Sub
