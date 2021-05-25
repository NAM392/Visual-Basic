
Private Sub FACTURA_Click()
    If Me!FACTURA Then
        Me!C_.Visible = True
        Me!CUIT.Visible = True
   
    Else
        Me!C_.Visible = False
        Me!CUIT.Visible = False
    End If
End Sub

Private Sub GUARDAR_Click()

'ELEGIR EL LIBRO DONDE QUIERO QUE SE GUARDE EL FORMULARIO

    
    