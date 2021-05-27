


Private Sub cuit_Click()

End Sub

Private Sub GUARDAR_Click()

    Dim num_CUIT As String

    Me!txtcuit.Visible = False 'hago que el cuit se invisible
    Me!cuit.Visible = False

    If (Me!Tape <> " " & Me!Tnom <> " ") Then
        ActiveCell.Offset(0, 1) = Me!Tape
        ActiveCell.Offset(0, 2) = Me!Tnom

        If (Me!FEM = True) Then ActiveCell.Offset(0, 3) = "FEMENINO"
    Else
        If (Me!MASC = True) Then ActiveCell.Offset(0, 3) = "MASCULINO"
    End If

    If (Me!PRIM = True) Then ActiveCell.Offset(0, 4) = "PRIMARIA"
    If (Me!SEC = True) Then ActiveCell.Offset(0, 4) = "SECUNDARIA"
    If (Me!UNI = True) Then ActiveCell.Offset(0, 4) = "UNIVERSITARIO"
    If (Me!PRIM = True) Then ActiveCell.Offset(0, 4) = "POSGRADO"
    If (Me!FAC = True) Then 'cuando marco la opcion factura A
        Me!txtcuit.Visible = True ' el texto de cuit se hace visible
        Me!cuit.Visible = True
        num_CUIT = Me!txtcuit     ' asigno num_CUIT con el contenido del texto cuit

    Else
        MsgBox("CAMPO INCONCLUSO")


    End If

End Sub


Private Sub UserForm_Initialize()

    'TODAS LAS FUNCIONES QUE OCURREN ANTES DE ENTRAR EN EL FORMULARIO VAN EN INITIALIZE

    For i = 1 To Sheets.Count   'recorro desde la primer hoja hasta la cuenta de las hojas
        If Sheets(i).Name = "Clientes" Then b = 1   'si una de las hojas se llama cliente entonces b=1
    Next
    If b <> 1 Then   'si b es 0 agrego una hoja y le pongo el nombre de clientes
     Sheets.Add
     Sheets(Sheets.Count).Name = "Clientes"
    End If
    Sheets("Clientes").Activate ' esto es lo mismo que el select
    Cells(2, 1).Select
    Do While IsEmpty(ActiveCell) = False  'mientras la celda activa este llena
        ActiveCell.Offset(1, 0).Select
    Loop
    cont = ActiveCell.Offset(-1, 0) + 1  'agrego un uno a cada cliente nuevo
    Me!NRO = Str(cont)                   'la variable cont esta declarada como publica en el MODULO

End Sub
