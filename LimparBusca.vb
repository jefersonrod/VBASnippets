Sub LimparBusca()
Dim planilhaBusca  As String
planilhaBusca = "Buscar"
'limpa area de movimentação ** Valor do RANGE Minimo "A3:AX1000"
    Worksheets(planilhaBusca).Range("A2:ZZ1000").Interior.ColorIndex = 0
    Worksheets(planilhaBusca).Range("A2:ZZ1000").Value = ""
    Worksheets(planilhaBusca).Range("A2:ZZ1000").Borders.LineStyle = xlNone
End Sub
