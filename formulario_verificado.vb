

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

    
    

   If (Me!APELLIDO = "") Or (Me!NOMBRE = "") Then
        MsgBox ("FALTA DATOS")
   Else:
        'ThisWorkbook.Select
        cantidad = Sheets.Count
        Sheets.Add after:=Sheets(cantidad)
        Sheets(Sheets.Count).Name = "CLIENTE N-" & Sheets.Count
        Sheets("CLIENTE N-" & Sheets.Count).Select
        Cells(1, 3) = "------EMPRESA ABC S.R.L---------"
        Cells(2, 1).Select
        ActiveCell.Offset(0, 0) = "NOMBRE  : "
        ActiveCell.Offset(0, 1) = Me!NOMBRE
        ActiveCell.Offset(2, 0) = "APELLIDO : "
        ActiveCell.Offset(2, 1) = Me!APELLIDO
        ActiveCell.Offset(3, 0) = "DIRECCION : "
        ActiveCell.Offset(3, 1) = Me!DIRECCION
        ActiveCell.Offset(4, 0) = "SEXO : "
        If (Me!FEMENINO = True) Then ActiveCell.Offset(4, 1) = "FEMENINO"
        If (Me!MASCULINO = True) Then ActiveCell.Offset(4, 1) = "MASCULINO"
        ActiveCell.Offset(5, 0) = "Nivel Educativo : "
        If (Me!PRIMARIO = True) Then ActiveCell.Offset(5, 1) = "PRIMARIO"
        If (Me!SECUNDARIO = True) Then ActiveCell.Offset(5, 1) = "SECUNDARIO"
        If (Me!UNIVERSITARIO = True) Then ActiveCell.Offset(5, 1) = "UNIVERSITARIO"
        If (Me!POSGRADO = True) Then ActiveCell.Offset(5, 1) = "POSGRADO"
        If (verificador(Me!CUIT) = 1) Then
           ActiveCell.Offset(6, 0) = "CUIT : "
           ActiveCell.Offset(6, 1) = Me!CUIT
        Else
            MsgBox ("CUIT ERRONEO")
        End If
        
  End If
    
    
    
    
     
    
End Sub

Function verificador(CUIT As String) As Byte
Dim msp As Long
Dim msd As Long
Dim prod As Long
Dim modulo As Integer
Dim j As Integer
Dim k As Integer
j = 2
k = 5


If ((Mid(CUIT, 1, 2)) = 20) Or ((Mid(CUIT, 1, 2)) = 23) Or ((Mid(CUIT, 1, 2)) = 24) Or ((Mid(CUIT, 1, 2)) = 27) Or ((Mid(CUIT, 1, 2)) = 30) Or ((Mid(CUIT, 1, 2)) = 33) Or ((Mid(CUIT, 1, 2)) = 34) Then
    
     
  For i = 0 To 5
   
    msp = msp + Val(Mid(CUIT, Len(CUIT) - 1 - i, 1)) * j
    j = j + 1
  Next
  
  For i = 1 To 4
  
    msd = msd + Val(Mid(CUIT, i, 1)) * k
    k = k - 1
  Next
  
  
  prod = msp + msd
  
  modulo = prod Mod 11
  

    If ((11 - modulo) = (Val(Mid(CUIT, Len(CUIT), 1)))) Then
        verificador = 1
    Else: verificador = 0
    
    End If
    
    
End If
    





End Function

Private Sub LIMPIAR_Click()

'Dim t As Object
'For Each t In PARCIAL_CLASE.Controls
'If TypeName(t) = True Then
't.Value = ""
'End If
'Next

'For Each c In Me.Controls
'
'On Error Resume Next
'
'c.Value = ""
'
'Next


Me!NOMBRE = ""
Me!APELLIDO = ""
Me!DIRECCION = ""
Me!CUIT = ""





End Sub

Private Sub UserForm_Initialize()
Me!C_.Visible = False
Me!CUIT.Visible = False

Me!REG = Sheets.Count




End Sub
