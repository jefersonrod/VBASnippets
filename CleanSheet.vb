Sub CleanSheet()
Dim planilha  As String

planilha = "Migradas"
'limpa area de movimentação ** Valor do RANGE Minimo "A3:AX1000"
Worksheets(planilha).Range("A6:CC1000").ClearContents
Worksheets(planilha).Range("A6:CC1000").Clear
    'Worksheets(planilha).Range("A6:CC1000").Interior.ColorIndex = 0
    'Worksheets(planilha).Range("A6:CC1000").Value = ""
    'Worksheets(planilha).Range("A3:ZZ1000").Interior.ColorIndex = 0
    'Worksheets(planilha).Range("A3:ZZ1000").Value = ""
End Sub