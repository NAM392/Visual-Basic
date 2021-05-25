Private Sub login_Click()

Dim user As String
Dim pass As String
Dim i As Integer
Dim existe As Integer


intentos = intentos + 1
user = Me!usuario
pass = Me!contrasena
i = 2
existe = 0

Do While (IsEmpty(Sheets("USUARIOS").Cells(i, 1)) = False) And (intentos <= 3) And (existe = 0)
    If (user = Sheets("USUARIOS").Cells(i, 1).Value) And (pass = Sheets("USUARIOS").Cells(i, 2).Value) Then existe = 1
    i = i + 1
Loop
 
If intentos = 3 Then
MsgBox ("No mas intentos restantes, intente mas tarde")
Application.Quit
End If

If existe = 1 Then
    Unload Me
    ingreso.Show
Else
    MsgBox ("Usuario o Contraseña no valida")
    Me!usuario = ""
    Me!contrasena = ""
End If
End Sub

Private Sub salir_Click()
Application.Quit
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Terminate()

End Sub
