Sub primera()
    MsgBox ("Bienvenidos al Sistema")
    ActiveCell.Value = "NNNN"

End Sub
Sub segunda()
    ActiveCell.Clear
End Sub

Sub tercera()
'Dim variable As Integer
pepe = InputBox("Deme un valor entero")
pepo = pepe * 3
ActiveCell = pepo
Cells(1, 4) = pepe
Cells(2, 4) = ActiveCell
ActiveCell.Offset(2, 2) = pepo
ActiveCell.Offset(-1, -1) = pepo + pepe
    ActiveCell.Offset(3, 3) = "END"
    Cells(5, 5).Select
End Sub

Sub cuarta()
    If condicion Then instruccion única
 pepe = InputBox("deme un numero")
    pepe = pepe + 1
    MsgBox(pepe)
End Sub
