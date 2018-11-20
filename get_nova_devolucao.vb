Sub get_nova_devolucao()
Dim a As Double
Dim CriterioBusca As Variant
Dim planilhademandas As String

planilhademandas = Sheets("config.ini").Range("B2")

a = 3
Do While Not Sheets(planilhademandas).Range("A" & a) = ""
    CriterioBusca = Sheets(planilhademandas).Range("P" & a)
    If IsError(CriterioBusca) And Not (Sheets(planilhademandas).Range("O" & a) = "Devolvida" Or Sheets(planilhademandas).Range("O" & a) = "Implantado") Then
            MsgBox Sheets(planilhademandas).Range("A" & a)
            Range("A" & a).Select
            Sheets(planilhademandas).Activate
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
    End If
    
  a = a + 1
 Loop

End Sub

