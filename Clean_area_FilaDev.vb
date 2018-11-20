Sub Clean_area_FilaDev()
Dim planilhaFilaDev  As String

planilhaFilaDev = Worksheets("config.ini").Cells(5, 2)
'limpa area de movimentação ** Valor do RANGE Minimo "A3:AX1000"
    Worksheets(planilhaFilaDev).Range("B2:ZZ1000").Interior.ColorIndex = 0
    Worksheets(planilhaFilaDev).Range("B2:ZZ1000").Value = ""
    Worksheets(planilhaFilaDev).Range("A3:ZZ1000").Interior.ColorIndex = 0
    Worksheets(planilhaFilaDev).Range("A3:ZZ1000").Value = ""
End Sub
