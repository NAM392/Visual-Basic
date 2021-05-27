Private Sub CANCEL_Click()
For Each w In Application.Workbooks
 w.Save
Next w
Application.Quit

End Sub

Private Sub GRABAR_Click()

    Dim cantidad As Integer

    cantidad = Sheets.Count
    Sheets.Add after:=Sheets(cantidad)
    Sheets(Sheets.Count).Name = "REGISTRO N-" & Sheets.Count
    Sheets("REGISTRO N-" & Sheets.Count).Select

    Cells(1, 1).Select
   ActiveCell.Offset(0, 0) = "NOMBRE DEL PROYECTO : "
   ActiveCell.Offset(0, 1) = Me!NOM_PROY
   ActiveCell.Offset(2, 0) = "CLIENTE : "
   ActiveCell.Offset(2, 1) = Me!CLIENTE
   ActiveCell.Offset(3, 0) = "RESPONSABLE : "
   ActiveCell.Offset(3, 1) = Me!RESPONSABLE
   ActiveCell.Offset(4, 0) = "FECHA DE FIRMA : "
   ActiveCell.Offset(4, 1) = (Me!DIA) & " / " & (Me!MES) & " / " & (Me!ANIO)
   ActiveCell.Offset(5, 0) = "MONTO CONTRACTUAL : "
   ActiveCell.Offset(5, 1) = "$" & Val(Me!MONTO)
   ActiveCell.Offset(5, 2) = "TIPO DE MONEDA : "
   ActiveCell.Offset(5, 3) = Me!MONEDA
   ActiveCell.Offset(6, 0) = "PESO O AREA : "
   ActiveCell.Offset(6, 1) = Val(Me!PESO)
   ActiveCell.Offset(6, 2) = "UNIDAD : "
    ActiveCell.Offset(6, 3) = Me!UNIDAD


End Sub


Private Sub UserForm_Initialize()
    Me!Lugar = "15"
    Me!hoja = Sheets.Count


    For Each celda In Range("PROYECTO")
        TIPO.AddItem celda
Next
    TIPO.ListIndex = 0


    For Each celda In Range("MONEDA")
        MONEDA.AddItem celda
Next
    MONEDA.ListIndex = 0



    For Each celda In Range("UNIDAD")
        UNIDAD.AddItem celda
Next
    UNIDAD.ListIndex = 0

End Sub


