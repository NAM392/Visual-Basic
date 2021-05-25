Sub primera_macro()
'
' primera_macro Macro
'

'
'   Range("A1:A4").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.Font.Size = 16
    Selection.Font.Italic = True
End Sub
