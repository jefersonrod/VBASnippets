Sub formCheckin()
'get info from sheet and generate card in Trello
On Error Resume Next
On Error GoTo 0

'set general vars
Dim plan As String
Dim analista As String
Dim registrado As String
Dim corpo As String
Dim colNum As Integer
Dim colSelect As String
Dim loja As String
Dim lojaCodigo As String
Dim line As String
Dim checkin As String
Dim checkout As String
Dim formCheckin As String
Dim usuarioAtual As String
Dim ivan As String
Dim jeferson As String
Dim luiz As String
Dim rener As String
Dim thiago As String
Dim datahoraAtual As Date
Dim tipo As String
'Técnico
Dim tecnico As String
Dim nomeTecnico As String
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
Dim nR As String
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

'set general vars
ivan = "Ivan Erison Gambarra da Silva"
jeferson = "Jeferson Rodrigues"
luiz = "Luiz Victor Lomba de Oliveira"
rener = "Renervaldo Wizenffat"
thiago = "Thiago Hiroshi Da Silva Endo"
plan = FunctionsTimeModelX.ActualSheetName
colSelect = coluna_Atual.coluna_Atual
registrado = FunctionsTimeModelX.Username
usuarioAtual = FunctionsTimeModelX.Username
nl = vbCrLf 'new line
line = "---------------------------------"
checkin = "*Check-In*"
tipo = "CheckIN"

If (colSelect = "C" And usuarioAtual = ivan) Then
    analista = "Ivan"
    colNum = 3
ElseIf (colSelect = "G" And usuarioAtual = jeferson) Then
    analista = "Jeferson"
    colNum = 7
ElseIf (colSelect = "K" And usuarioAtual = luiz) Then
    analista = "Luiz"
    colNum = 11
ElseIf (colSelect = "O" And usuarioAtual = rener) Then
    analista = "Rener"
    colNum = 15
ElseIf (colSelect = "S" And usuarioAtual = thiago) Then
    analista = "Thiago"
    colNum = 19
    
Else
    MsgBox ("Ola " + usuarioAtual + " coluna selecionada " + colSelect + " é invalida, verifique")
    Exit Sub
End If
    
    
    'Check loja field
    If (Worksheets(plan).Cells(5, colNum) = "") Then
        MsgBox ("Favor preencher o numero da loja na linha 5")
        Worksheets(plan).Cells(5, colNum).Select
        Exit Sub
        
    Else
        'check tecnico answer
        For i = 7 To 19 Step 2
            If (Worksheets(plan).Cells(i, colNum) = "") Then
            MsgBox ("Favor preencher a linha " + CStr(i))
            Worksheets(plan).Cells(i, colNum).Select
            Exit Sub
            End If
            
        Next i
        
        'check responsavel field
        For i = 21 To 31 Step 2
            If (Worksheets(plan).Cells(i, colNum) = "") Then
            MsgBox ("Favor preencher a linha " + CStr(i))
            Worksheets(plan).Cells(i, colNum).Select
            Exit Sub
            End If
            
        Next i
        
        loja = Worksheets(plan).Cells(4, colNum)
        lojaCodigo = Worksheets(plan).Cells(5, colNum)
        'Técnico
        tecnico = Worksheets(plan).Cells(6, colNum)
        nomeTecnico = Worksheets(plan).Cells(7, colNum)
        recebeuContatodaPrimesys = Worksheets(plan).Cells(8, colNum)
        recebeuContatodaPrimesysRESP = Worksheets(plan).Cells(9, colNum)
        recebeuOrientacõesSobreoManualdeMigracao = Worksheets(plan).Cells(10, colNum)
        recebeuOrientacõesSobreoManualdeMigracaoRESP = Worksheets(plan).Cells(11, colNum)
        jaRealizouMigracao = Worksheets(plan).Cells(12, colNum)
        jaRealizouMigracaoRESP = Worksheets(plan).Cells(13, colNum)
        possuiWhatsappQual = Worksheets(plan).Cells(14, colNum)
        possuiWhatsappQualRESP = Worksheets(plan).Cells(15, colNum)
        informarSobreoLinkqueEstaSendoInstalado = Worksheets(plan).Cells(16, colNum)
        informarSobreoLinkqueEstaSendoInstaladoRESP = Worksheets(plan).Cells(17, colNum)
        envioFotosRackRetaguardaBalcao = Worksheets(plan).Cells(18, colNum)
        envioFotosRackRetaguardaBalcaoRESP = Worksheets(plan).Cells(19, colNum)
        'Responsavel
        responsavel = Worksheets(plan).Cells(20, colNum)
        nR = Worksheets(plan).Cells(21, colNum)
        InformarSobreAcompanhamento = Worksheets(plan).Cells(22, colNum)
        InformarSobreAcompanhamentoRESP = Worksheets(plan).Cells(23, colNum)
        temAlgumChamadoAberto = Worksheets(plan).Cells(24, colNum)
        temAlgumChamadoAbertoRESP = Worksheets(plan).Cells(25, colNum)
        estaComAlgumProblemaSistemico = Worksheets(plan).Cells(26, colNum)
        estaComAlgumProblemaSistemicoRESP = Worksheets(plan).Cells(27, colNum)
        orientarAssinarOSSomenteApos = Worksheets(plan).Cells(28, colNum)
        orientarAssinarOSSomenteAposRESP = Worksheets(plan).Cells(29, colNum)
        confirmaroNumerodoTelefone = Worksheets(plan).Cells(30, colNum)
        confirmaroNumerodoTelefoneRESP = Worksheets(plan).Cells(31, colNum)
        
    End If
    
    
    
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
            orientarAssinarOSSomenteAposRESP + nl + confirmaroNumerodoTelefone + nl + confirmaroNumerodoTelefoneRESP + nl + CStr(datahoraAtual) + nl + registrado
    
    Call CheckLogIN.CheckLogIN(analista, lojaCodigo, recebeuContatodaPrimesysRESP, recebeuOrientacõesSobreoManualdeMigracaoRESP, jaRealizouMigracaoRESP, possuiWhatsappQualRESP, informarSobreoLinkqueEstaSendoInstaladoRESP, envioFotosRackRetaguardaBalcaoRESP, InformarSobreAcompanhamentoRESP, temAlgumChamadoAbertoRESP, estaComAlgumProblemaSistemicoRESP, orientarAssinarOSSomenteAposRESP, confirmaroNumerodoTelefoneRESP, registrado, nomeTecnico, nR)
    Call TrelloCheckin.TrelloCheckin(tipo, analista, lojaCodigo, recebeuContatodaPrimesysRESP, recebeuOrientacõesSobreoManualdeMigracaoRESP, jaRealizouMigracaoRESP, possuiWhatsappQualRESP, informarSobreoLinkqueEstaSendoInstaladoRESP, envioFotosRackRetaguardaBalcaoRESP, InformarSobreAcompanhamentoRESP, temAlgumChamadoAbertoRESP, estaComAlgumProblemaSistemicoRESP, orientarAssinarOSSomenteAposRESP, confirmaroNumerodoTelefoneRESP, registrado, nomeTecnico, nR)
    'generate content
    CopyText corpo 'call copy to clipboard function
    
    'change color of the loja (colunm 1) to mark as sent
    'Worksheets(plan).Cells(linhaAtual, colLoja).Font.Color = corAzul

End Sub
