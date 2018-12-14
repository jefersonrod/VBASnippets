Sub BuscaLojaTrello()
Dim loja As String
Dim url As String
Dim urlTrello As String
Dim linhaAtual As Integer
Dim plan As String

urlTrello = "https://trello.com/search?q="
plan = FunctionsTimeModelX.ActualSheetName
linhaAtual = linha_Atual.linha_Atual
loja = Worksheets(plan).Cells(linhaAtual, 3)
url = urlTrello + loja

If (loja = "" Or loja = " ") Then
    MsgBox ("numero da loja esta vazio verifique!")
Else
    ThisWorkbook.FollowHyperlink (url)
End If


End Sub
