Sub Search4All()
'set general variables
Dim plan As String
Dim linhaExecute As Integer
Dim linhaAtual As Integer
Dim linhaFinal As Integer
Dim qtdeLinhas As Integer
Dim estTime As Integer
Dim colLoja As Integer
Dim colRelatFotog As Integer
Dim colSegs As Integer
'feed variables position
colLoja = 1
colRelatFotog = 2
colSegs = 4
'feed general variables
plan = "Buscar"
linhaExecute = FunctionsBuscaRelatFotog.linha_Atual
linhaAtual = linhaExecute
qtdeLinhas = FunctionsBuscaRelatFotog.qtde_linhas
linhaFinal = (linhaAtual + qtdeLinhas) + 1

Do Until (Worksheets(plan).Cells(linhaAtual, colLoja) = "")
    'Debug.Print Application.ActiveCell.Address
    ActiveSheet.Cells(linhaAtual, colRelatFotog).Select
    Do Until (Worksheets(plan).Cells(linhaAtual, colRelatFotog) <> "")
        Search4MailOutlook.Search4MailOutlook
        estTime = (Worksheets(plan).Cells(linhaAtual, colSegs) * qtdeLinhas)
        Worksheets(plan).Cells(linhaFinal, colLoja) = "Tempo estimado para conclusão: " + CStr(estTime) + " segundos"
    Loop
    linhaAtual = linhaAtual + 1
Loop



End Sub
