Option Explicit
Sub TrelloCheckin(tipo As String, analista As String, lojaCodigo As String, recebeuContatodaPrimesysRESP As String, recebeuOrientacõesSobreoManualdeMigracaoRESP As String, jaRealizouMigracaoRESP As String, possuiWhatsappQualRESP As String, informarSobreoLinkqueEstaSendoInstaladoRESP As String, envioFotosRackRetaguardaBalcaoRESP As String, InformarSobreAcompanhamentoRESP As String, temAlgumChamadoAbertoRESP As String, estaComAlgumProblemaSistemicoRESP As String, orientarAssinarOSSomenteAposRESP As String, confirmaroNumerodoTelefoneRESP As String, registrado As String, nomeTecnico As String, nR As String)
'set error definition
On Error Resume Next
On Error GoTo 0
'set email variable
Dim OutApp As Object
Dim OutMail As Object
'set general variable
Dim plan As String
Dim corpo As String
Dim email As String
Dim corpoEmail As String
Dim trello As String
Dim assunto As String
Dim emailBoardTrello As String
Dim trelloUser As String
Dim nl As Variant
Dim br As String
Dim datahoraAtual As Date
Dim gerado As String
Dim criado As String

'set config.ini vars
Dim linhaConfig As Integer
'set body vars
Dim checkin As String
Dim loja As String
Dim tecnico As String
Dim recebeuContatodaPrimesys As String
Dim recebeuOrientacõesSobreoManualdeMigracao As String
Dim jaRealizouMigracao As String
Dim possuiWhatsappQual As String
Dim informarSobreoLinkqueEstaSendoInstalado As String
Dim envioFotosRackRetaguardaBalcao As String
Dim line As String
Dim responsavel As String
Dim InformarSobreAcompanhamento As String
Dim temAlgumChamadoAberto As String
Dim estaComAlgumProblemaSistemico As String
Dim orientarAssinarOSSomenteApos As String
Dim confirmaroNumerodoTelefone As String
'feed body vars
checkin = "**Check-IN**"
loja = "Loja:"
tecnico = "Nome do técnico:"
recebeuContatodaPrimesys = "- Recebeu contato da Primesys?"
recebeuOrientacõesSobreoManualdeMigracao = "- Recebeu orientações sobre o manual de migração?"
jaRealizouMigracao = "- Já realizou migração?"
possuiWhatsappQual = "- Possui Whatsapp? Qual?"
informarSobreoLinkqueEstaSendoInstalado = "- Informar sobre o link que está sendo instalado ou migrado para a nova solução."
envioFotosRackRetaguardaBalcao = "- Envio fotos rack, retaguarda, balcão, cabeamentos (balcão e PDV’s)"
line = "---"
responsavel = "Nome do responsável:"
InformarSobreAcompanhamento = "- Informar sobre acompanhamento da equipe boticário."
temAlgumChamadoAberto = "- Tem algum chamado aberto?"
estaComAlgumProblemaSistemico = "- Está com algum problema sistêmico ou em equipamentos?"
orientarAssinarOSSomenteApos = "- Orientar assinar OS somente após todos os testes."
confirmaroNumerodoTelefone = "- Confirmar o numero do telefone (fixo ou celular da loja)"
'feed vars position from config.ini
emailBoardTrello = getAnalistaBoardTrello(analista)
trelloUser = getAnalistaTrelloUser(analista)
'set general var value
plan = "CheckLog"
nl = vbCrLf 'new line
br = "<br>" 'new line HTML
gerado = "Gerado em: "
criado = "Criado por: "
    
    'Check mail address to send
    If (emailBoardTrello = "") Then
        MsgBox ("Preencha o endereço de e-mail do Board do Trello na aba CONFIG.INI do analista " + analista)
        Exit Sub
    End If
    
    'Check Trello address to add as a member
    If (trelloUser = "") Then
        MsgBox ("Preencha o endereço de usuario do Trello na aba CONFIG.INI do analista " + analista)
        Exit Sub
    End If
    

    'Feed email fields
    datahoraAtual = Now()
    assunto = "#" + lojaCodigo + " - " + tipo
    corpo = checkin + br + _
            "##" + loja + lojaCodigo + "##" + br + br + _
            tecnico + " " + nomeTecnico + br + br + _
            recebeuContatodaPrimesys + br + _
            recebeuContatodaPrimesysRESP + br + _
            recebeuOrientacõesSobreoManualdeMigracao + br + _
            recebeuOrientacõesSobreoManualdeMigracaoRESP + br + _
            jaRealizouMigracao + br + _
            jaRealizouMigracaoRESP + br + _
            possuiWhatsappQual + br + _
            possuiWhatsappQualRESP + br + _
            informarSobreoLinkqueEstaSendoInstalado + br + _
            informarSobreoLinkqueEstaSendoInstaladoRESP + br + _
            envioFotosRackRetaguardaBalcao + br + _
            envioFotosRackRetaguardaBalcaoRESP + br + br + _
            line + br + _
            responsavel + " " + nR + br + br + _
            InformarSobreAcompanhamento + br + _
            InformarSobreAcompanhamentoRESP + br + _
            temAlgumChamadoAberto + br + _
            temAlgumChamadoAbertoRESP + br + _
            estaComAlgumProblemaSistemico + br + _
            estaComAlgumProblemaSistemicoRESP + br + _
            orientarAssinarOSSomenteApos + br + _
            orientarAssinarOSSomenteAposRESP + br + confirmaroNumerodoTelefone + br + confirmaroNumerodoTelefoneRESP + br + br + ">" + gerado + CStr(datahoraAtual) + br + ">" + criado + registrado + br + br
       
    corpoEmail = trelloUser + br + corpo
    'setup app definitions
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    
    'Build Email
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    On Error Resume Next
    
    'Send EMail
    With OutMail
        .To = emailBoardTrello
        '.CC = ""
        '.BCC = ""
        .Subject = assunto
        .HTMLBody = corpoEmail
        .Send
    End With

    
    'setup app definitions enable
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With


End Sub
