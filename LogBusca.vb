Sub LogBusca(Status As Boolean, Item As Integer, Buscado As String, Encontrado As String)
'set general variables
Dim planLog As String
Dim linhaAtualLog As Integer
Dim statusLog As String
'set positions

Dim colDataLog As Integer
Dim colHoraLog As Integer
Dim colStatusLog As Integer
Dim colItem As Integer
Dim colBuscadoLog As Integer
Dim colEncontradoLog As Integer

'feed general variables
planLog = "Log"
'feed positions
colDataLog = 1
colHoraLog = 2
colStatusLog = 3
colItem = 4
colBuscadoLog = 5
colEncontradoLog = 6

linhaAtualLog = FunctionsBuscaRelatFotog.busca_ultima_linha_vazia_log 'search for last line empty

If (Status) Then
    statusLog = "OK"
Else
    statusLog = "NO"
End If
'fill log tab
Worksheets(planLog).Cells(linhaAtualLog, colDataLog) = Date
Worksheets(planLog).Cells(linhaAtualLog, colHoraLog) = Time
Worksheets(planLog).Cells(linhaAtualLog, colStatusLog) = statusLog
Worksheets(planLog).Cells(linhaAtualLog, colItem) = Item
Worksheets(planLog).Cells(linhaAtualLog, colBuscadoLog) = Buscado
Worksheets(planLog).Cells(linhaAtualLog, colEncontradoLog) = Encontrado

End Sub
