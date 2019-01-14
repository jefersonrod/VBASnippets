Option Explicit
Public Function CheckLogOUT(anlst As String, loja As String, aPrimesysDeuasOrientacoesRESP As String, casoHajaumProblemaOcorridoRESP As String, solicitarUmaAvaliacaodoTecnicoRESP As String, seHouverMaisItensRESP As String, registrado As String)
'get info from sheet and generate card in Trello
On Error Resume Next
On Error GoTo 0

'set general vars
Dim plan As String
Dim tipo As String
Dim linha As Integer


'columns
Dim colLoja As Integer
Dim colData As Integer
Dim colHora As Integer
Dim colTipo As Integer
Dim colCOUTTecL8colP As Integer
Dim colCOUTResL12colQ As Integer
Dim colCOUTResL14colR As Integer
Dim colCOUTResL16colS As Integer
Dim colAnalistacolT As Integer
Dim colRegistradocolU As Integer

'positions
colLoja = 1
colData = 2
colHora = 3
colTipo = 4
colCOUTTecL8colP = 16
colCOUTResL12colQ = 17
colCOUTResL14colR = 18
colCOUTResL16colS = 19
colAnalistacolT = 20
colRegistradocolU = 21
'set general vars
plan = "CheckLog"
tipo = "OUT"
linha = ultimaLinhaLivreChkLog.ultimaLinhaLivreChkLog

'feed log
Worksheets(plan).Cells(linha, colLoja) = loja
Worksheets(plan).Cells(linha, colData) = Date
Worksheets(plan).Cells(linha, colHora) = Time
Worksheets(plan).Cells(linha, colTipo) = tipo
Worksheets(plan).Cells(linha, colCOUTTecL8colP) = aPrimesysDeuasOrientacoesRESP
Worksheets(plan).Cells(linha, colCOUTResL12colQ) = casoHajaumProblemaOcorridoRESP
Worksheets(plan).Cells(linha, colCOUTResL14colR) = solicitarUmaAvaliacaodoTecnicoRESP
Worksheets(plan).Cells(linha, colCOUTResL16colS) = seHouverMaisItensRESP
Worksheets(plan).Cells(linha, colAnalistacolT) = anlst
Worksheets(plan).Cells(linha, colRegistradocolU) = registrado
End Function

