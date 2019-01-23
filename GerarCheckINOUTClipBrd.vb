Sub GerarCheckINOUTClipBrd()
'get info from sheet and generate card in Trello
On Error Resume Next
On Error GoTo 0

'set general vars
Dim plan As String
Dim planIN As String
Dim planOUT As String
Dim tipo As String
Dim linha As Integer
Dim datastr As String
Dim horastr As String
Dim corpo As String
Dim loja As String
Dim lojaCodigo As String
Dim checkin As String
Dim checkout As String
Dim nl As String
Dim line As String
'columns
Dim colLoja As Integer
Dim collojaCodigo As Integer
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
Dim colCOUTTecL8colP As Integer
Dim colCOUTResL12colQ As Integer
Dim colCOUTResL14colR As Integer
Dim colCOUTResL16colS As Integer
Dim colAnalistacolT As Integer
Dim colRegistradocolU As Integer


'set form vars checkIN
'Técnico
Dim tecnico As String
Dim recebeuContatodaPrimesys As String
Dim recebeuContatodaPrimesysRESP As String
Dim recebeuOrientacõesSobreoManualdeMigracao As String
Dim recebeuOrientacõesSobreoManualdeMigracaoRESP As String
Dim jaRealizouMigracao As String
Dim jaRealizouMigracaoRESP As String
Dim possuiWhatsappQual As String
Dim possuiWhatsappQualRESP As String
Dim informarSobreoLinkqueEstaSendoInstalado As String
Dim informarSobreoLinkqueEstaSendoInstaladoRESP As String
Dim envioFotosRackRetaguardaBalcao As String
Dim envioFotosRackRetaguardaBalcaoRESP As String

'Responsavel
Dim resposavel As String
Dim InformarSobreAcompanhamento As String
Dim InformarSobreAcompanhamentoRESP As String
Dim temAlgumChamadoAberto As String
Dim temAlgumChamadoAbertoRESP As String
Dim estaComAlgumProblemaSistemico As String
Dim estaComAlgumProblemaSistemicoRESP As String
Dim orientarAssinarOSSomenteApos As String
Dim orientarAssinarOSSomenteAposRESP As String
Dim confirmaroNumerodoTelefone As String
Dim confirmaroNumerodoTelefoneRESP As String

'set form vars checkOUT
'Técnico
Dim aPrimesysDeuasOrientacoes As String
Dim aPrimesysDeuasOrientacoesRESP As String
'Responsavel
Dim casoHajaumProblemaOcorrido As String
Dim casoHajaumProblemaOcorridoRESP As String
Dim solicitarUmaAvaliacaodoTecnico As String
Dim solicitarUmaAvaliacaodoTecnicoRESP As String
Dim seHouverMaisItens As String
Dim seHouverMaisItensRESP As String

'positions
collojaCodigo = 1
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
colCOUTTecL8colP = 16
colCOUTResL12colQ = 17
colCOUTResL14colR = 18
colCOUTResL16colS = 19
colAnalistacolT = 20
colRegistradocolU = 21
colAnalistacolT = 20
colRegistradocolU = 21

'set general vars
plan = "CheckLog"
planIN = "Checkin"
planOUT = "Checkout"
checkin = "*Check-In*"
checkout = "*Check-Out*"
linha = linha_Atual.linha_Atual
tipo = Worksheets(plan).Cells(linha, colTipo)
nl = vbCrLf 'new line
line = "---------------------------------"

If (tipo = "IN") Then

    'feed vars
    lojaCodigo = Worksheets(plan).Cells(linha, collojaCodigo)
    datastr = Worksheets(plan).Cells(linha, colData)
    horastr = Worksheets(plan).Cells(linha, colHora)
    
    recebeuContatodaPrimesysRESP = Worksheets(plan).Cells(linha, colCINTecL8colE)
    recebeuOrientacõesSobreoManualdeMigracaoRESP = Worksheets(plan).Cells(linha, colCINTecL10colF)
    jaRealizouMigracaoRESP = Worksheets(plan).Cells(linha, colCINTecL12colG)
    possuiWhatsappQualRESP = Worksheets(plan).Cells(linha, colCINTecL14colH)
    informarSobreoLinkqueEstaSendoInstaladoRESP = Worksheets(plan).Cells(linha, colCINTecL16colI)
    envioFotosRackRetaguardaBalcaoRESP = Worksheets(plan).Cells(linha, colCINTecL18colJ)
    InformarSobreAcompanhamentoRESP = Worksheets(plan).Cells(linha, colCINResL22colK)
    temAlgumChamadoAbertoRESP = Worksheets(plan).Cells(linha, colCINResL24colL)
    estaComAlgumProblemaSistemicoRESP = Worksheets(plan).Cells(linha, colCINResL26colM)
    orientarAssinarOSSomenteAposRESP = Worksheets(plan).Cells(linha, colCINResL28colN)
    confirmaroNumerodoTelefoneRESP = Worksheets(plan).Cells(linha, colCINResL30colO)
    anlst = Worksheets(plan).Cells(linha, colAnalistacolT)
    registrado = Worksheets(plan).Cells(linha, colRegistradocolU)
    
    'get text from form
    'tecnico
    loja = Worksheets(planIN).Cells(4, 3)
    tecnico = Worksheets(planIN).Cells(7, 3)
    recebeuContatodaPrimesys = Worksheets(planIN).Cells(8, 3)
    recebeuOrientacõesSobreoManualdeMigracao = Worksheets(planIN).Cells(10, 3)
    jaRealizouMigracao = Worksheets(planIN).Cells(12, 3)
    possuiWhatsappQual = Worksheets(planIN).Cells(14, 3)
    informarSobreoLinkqueEstaSendoInstalado = Worksheets(planIN).Cells(16, 3)
    envioFotosRackRetaguardaBalcao = Worksheets(planIN).Cells(18, 3)
        
    'Responsavel
    responsavel = Worksheets(planIN).Cells(21, 3)
    InformarSobreAcompanhamento = Worksheets(planIN).Cells(22, 3)
    temAlgumChamadoAberto = Worksheets(planIN).Cells(24, 3)
    estaComAlgumProblemaSistemico = Worksheets(planIN).Cells(26, 3)
    orientarAssinarOSSomenteApos = Worksheets(planIN).Cells(28, 3)
    confirmaroNumerodoTelefone = Worksheets(planIN).Cells(30, 3)
    
    'generate content text
    
    corpo = checkin + nl + _
            loja + lojaCodigo + nl + nl + _
            tecnico + nl + nl + _
            recebeuContatodaPrimesys + nl + _
            recebeuContatodaPrimesysRESP + nl + _
            recebeuOrientacõesSobreoManualdeMigracao + nl + _
            recebeuOrientacõesSobreoManualdeMigracaoRESP + nl + _
            jaRealizouMigracao + nl + _
            jaRealizouMigracaoRESP + nl + _
            possuiWhatsappQual + nl + _
            possuiWhatsappQualRESP + nl + _
            informarSobreoLinkqueEstaSendoInstalado + nl + _
            informarSobreoLinkqueEstaSendoInstaladoRESP + nl + _
            envioFotosRackRetaguardaBalcao + nl + _
            envioFotosRackRetaguardaBalcaoRESP + nl + nl + _
            line + nl + _
            responsavel + nl + nl + _
            InformarSobreAcompanhamento + nl + _
            InformarSobreAcompanhamentoRESP + nl + _
            temAlgumChamadoAberto + nl + _
            temAlgumChamadoAbertoRESP + nl + _
            estaComAlgumProblemaSistemico + nl + _
            estaComAlgumProblemaSistemicoRESP + nl + _
            orientarAssinarOSSomenteApos + nl + _
            orientarAssinarOSSomenteAposRESP + nl + confirmaroNumerodoTelefone + nl + confirmaroNumerodoTelefoneRESP

ElseIf (tipo = "OUT") Then
    'feed vars
    lojaCodigo = Worksheets(plan).Cells(linha, collojaCodigo)
    datastr = Worksheets(plan).Cells(linha, colData)
    horastr = Worksheets(plan).Cells(linha, colHora)
    
    tipo = Worksheets(plan).Cells(linha, colTipo)
    aPrimesysDeuasOrientacoesRESP = Worksheets(plan).Cells(linha, colCOUTTecL8colP)
    casoHajaumProblemaOcorridoRESP = Worksheets(plan).Cells(linha, colCOUTResL12colQ)
    solicitarUmaAvaliacaodoTecnicoRESP = Worksheets(plan).Cells(linha, colCOUTResL14colR)
    seHouverMaisItensRESP = Worksheets(plan).Cells(linha, colCOUTResL16colS)
    anlst = Worksheets(plan).Cells(linha, colAnalistacolT)
    registrado = Worksheets(plan).Cells(linha, colRegistradocolU)
    
    'get text from form
    'Técnico
        loja = Worksheets(planOUT).Cells(4, 3)
        tecnico = Worksheets(planOUT).Cells(7, 3)
        aPrimesysDeuasOrientacoes = Worksheets(planOUT).Cells(8, 3)
        'Responsavel
        responsavel = Worksheets(planOUT).Cells(11, 3)
        casoHajaumProblemaOcorrido = Worksheets(planOUT).Cells(12, 3)
        solicitarUmaAvaliacaodoTecnico = Worksheets(planOUT).Cells(14, 3)
        seHouverMaisItens = Worksheets(planOUT).Cells(16, 3)
        
    
    corpo = checkout + nl + _
            loja + lojaCodigo + nl + nl + _
            tecnico + nl + nl + _
            aPrimesysDeuasOrientacoes + nl + _
            aPrimesysDeuasOrientacoesRESP + nl + nl + _
            line + nl + _
            responsavel + nl + nl + _
            casoHajaumProblemaOcorrido + nl + _
            casoHajaumProblemaOcorridoRESP + nl + _
            solicitarUmaAvaliacaodoTecnico + nl + _
            solicitarUmaAvaliacaodoTecnicoRESP + nl + _
            seHouverMaisItens + nl + _
            seHouverMaisItensRESP + nl
            
Else
    MsgBox ("Tipo invalido " + tipo + nl + "Verifique a linha " + linha)
End If

'generate content
    CopyText corpo 'call copy to clipboard function

End Sub
