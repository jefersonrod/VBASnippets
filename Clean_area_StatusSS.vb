Sub Clean_area_StatusSS()
Dim planilhaDestino  As String

planilhaDestino = Worksheets("config.ini").Cells(3, 2)
'limpa area de movimentação ** Valor do RANGE Minimo "A3:AX1000"
    Worksheets(planilhaDestino).Range("A3:AX1000").Interior.ColorIndex = 0
    Worksheets(planilhaDestino).Range("A3:AX1000").Value = ""
End Sub
