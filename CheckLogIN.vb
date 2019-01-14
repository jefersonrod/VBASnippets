Option Explicit
Public Function CheckLogIN(anlst As String, loja As String, recebeuContatodaPrimesysRESP As String, recebeuOrientacõesSobreoManualdeMigracaoRESP As String, jaRealizouMigracaoRESP As String, possuiWhatsappQualRESP As String, informarSobreoLinkqueEstaSendoInstaladoRESP As String, envioFotosRackRetaguardaBalcaoRESP As String, InformarSobreAcompanhamentoRESP As String, temAlgumChamadoAbertoRESP As String, estaComAlgumProblemaSistemicoRESP As String, orientarAssinarOSSomenteAposRESP As String, confirmaroNumerodoTelefoneRESP As String, registrado As String)
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
Dim colCINTecL8colE As Integer
Dim colCINTecL10colF As Integer
Dim colCINTecL12colG As Integer
Dim colCINTecL14colH As Integer
Dim colCINTecL16colI As Integer
Dim colCINTecL18colJ As Integer
Dim colCINResL22colK As Integer
Dim colCINResL24colL As Integer
Dim colCINResL26colM As Integer
Dim colCINResL28colN As Integer
Dim colCINResL30colO As Integer
Dim colAnalistacolT As Integer
Dim colRegistradocolU As Integer

'positions
colLoja = 1
colData = 2
colHora = 3
colTipo = 4
colCINTecL8colE = 5
colCINTecL10colF = 6
colCINTecL12colG = 7
colCINTecL14colH = 8
colCINTecL16colI = 9
colCINTecL18colJ = 10
colCINResL22colK = 11
colCINResL24colL = 12
colCINResL26colM = 13
colCINResL28colN = 14
colCINResL30colO = 15
colAnalistacolT = 20
colRegistradocolU = 21

'set general vars
plan = "CheckLog"
tipo = "IN"
linha = ultimaLinhaLivreChkLog.ultimaLinhaLivreChkLog

'feed log
Worksheets(plan).Cells(linha, colLoja) = loja
Worksheets(plan).Cells(linha, colData) = Date
Worksheets(plan).Cells(linha, colHora) = Time
Worksheets(plan).Cells(linha, colTipo) = tipo
Worksheets(plan).Cells(linha, colCINTecL8colE) = recebeuContatodaPrimesysRESP
Worksheets(plan).Cells(linha, colCINTecL10colF) = recebeuOrientacõesSobreoManualdeMigracaoRESP
Worksheets(plan).Cells(linha, colCINTecL12colG) = jaRealizouMigracaoRESP
Worksheets(plan).Cells(linha, colCINTecL14colH) = possuiWhatsappQualRESP
Worksheets(plan).Cells(linha, colCINTecL16colI) = informarSobreoLinkqueEstaSendoInstaladoRESP
Worksheets(plan).Cells(linha, colCINTecL18colJ) = envioFotosRackRetaguardaBalcaoRESP
Worksheets(plan).Cells(linha, colCINResL22colK) = InformarSobreAcompanhamentoRESP
Worksheets(plan).Cells(linha, colCINResL24colL) = temAlgumChamadoAbertoRESP
Worksheets(plan).Cells(linha, colCINResL26colM) = estaComAlgumProblemaSistemicoRESP
Worksheets(plan).Cells(linha, colCINResL28colN) = orientarAssinarOSSomenteAposRESP
Worksheets(plan).Cells(linha, colCINResL30colO) = confirmaroNumerodoTelefoneRESP
Worksheets(plan).Cells(linha, colAnalistacolT) = anlst
Worksheets(plan).Cells(linha, colRegistradocolU) = registrado

End Function
