Sub Botón2_Haga_clic_en()






    Dim NOMBRE As String
    Dim i As Integer
    Dim MAX As Integer
    Dim s As CellFormat
    
    
    
  
   
    Cells(2, 2).Select
    
    Do Until IsEmpty(ActiveCell)
               
        ActiveCell.Offset(1, 0).Select
        If (MAX < ActiveCell) Then
            MAX = ActiveCell
        End If
   
    Loop
    
       Cells(2, 3).Select
       Do Until IsEmpty(ActiveCell)
               
        ActiveCell.Offset(1, 0).Select
        If (MAX < ActiveCell) Then
            MAX = ActiveCell
        End If
   
    Loop
        Cells(2, 4).Select
        Do Until IsEmpty(ActiveCell)
               
        ActiveCell.Offset(1, 0).Select
        If (MAX < ActiveCell) Then
            MAX = ActiveCell
            ActiveCell.Font.Color = vbRed
           
        End If
   
    Loop
    
 
   
    MsgBox (MAX)
     





End Sub
