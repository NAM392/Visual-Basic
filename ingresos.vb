Private Sub G_DATOS_Click()
                            'ingreso uno debajo del otro guardo en hoja nueva y si quiero guardo librpo
Dim cantidad As Integer
Dim flag As Byte
Dim NOMBRE As String


    cantidad = Sheets.Count
    
    For f = 1 To cantidad
        If Sheets(f).Name = "DATOS" Then flag = 1
    
     Next
     
     If flag = 1 Then
        Sheets("DATOS").Select
       
     Else
         Sheets.Add after:=Sheets(cantidad)
         Sheets(Sheets.Count).Name = "DATOS"
        
     End If
 
    
     
    Do Until IsEmpty(ActiveCell)
    
        ActiveCell.Offset(1, 0).Select 'me muevo a la fila de abajo

    Loop
    
    ActiveCell = Me!DATOS
   
   
    Me!DATOS = ""





End Sub

Private Sub G_LIBRO_Click()
For Each w In Application.Workbooks
 w.Save
Next w
End Sub

Private Sub i__Click()

End Sub

Private Sub UserForm_Click()

End Sub