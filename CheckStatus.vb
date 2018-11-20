Sub CheckStatus()

Call ZerarLineCount

Dim count As Integer
Dim countSheetTarget As Integer
Dim lastPositionFree As Integer
Dim lineCount As Integer

Dim line As String

Dim today As Date
Dim planilhaBase As String
Dim planilhaDestino As String
'Status SSs
Dim ssBase As String
Dim assunto As String
Dim site As String
Dim tipo As String
Dim prio As String
Dim statusInterno As String
Dim dataEstimada As Date
Dim dataEnvio As Date
Dim dataAprov As Date
Dim dataAloc As Date
Dim pendente As String

Dim statusAtual As String
Dim gestor As String
Dim recurso As String
Dim esforco As Date
'var from tabelinha
Dim filaDev As String
Dim implement As String
'cores
Dim corVermelho As Long
Dim corVerde As Long
Dim corBranco As Long
Dim corLaranja As Long
'position from planilhaBase
Dim bssBase As Integer
Dim bassunto As Integer
Dim bsite As Integer
Dim btipo As Integer
Dim bprio As Integer
Dim bstatusInterno As Integer
Dim bstatusAtual  As Integer
Dim bdataAloc  As Integer
Dim bdataInicio  As Integer
Dim bdataEstimada  As Integer
Dim bdataEnvio  As Integer
Dim bdataAprov  As Integer
Dim bgestor  As Integer
Dim brecurso  As Integer
Dim besforco  As Integer
'position from planilhaDestino
Dim dssBase As Integer
Dim dassunto As Integer
Dim dsite As Integer
Dim dtipo As Integer
Dim dprio As Integer
Dim dstatusInterno As Integer
Dim dstatusAtual As Integer
Dim dpendente As Integer
Dim ddataInicio As Integer
Dim desforco As Integer
Dim ddataEstimada  As Integer
Dim dgestor As Integer
Dim drecurso As Integer
Dim ddataAloc As Integer
Dim ddataAprov As Integer
Dim ddataEnvio As Integer
Dim dfarol As Integer
Dim destimativa As Integer
Dim LastRowEmpty As Long

'variaveis de exceção
Dim excep As String
Dim excepMark As String
Dim dtStartExcep As Date
Dim dtEndExcep As Date
Dim commentExcep As String
Dim bcolExcep As Integer
Dim bcolStartExcep As Integer
Dim bcolEndExcep As Integer
Dim bcolComment As Integer
Dim dcolExcep As Integer
Dim dcolStartExcep As Integer
Dim dcolEndExcep As Integer
Dim dcolComment As Integer

'feed that damn vars
planilhaBase = Worksheets("config.ini").Cells(2, 2)
planilhaDestino = Worksheets("config.ini").Cells(3, 2)
lastPositionFree = Worksheets("config.ini").Cells(4, 2)
today = Date
'line = ReadAddress() 'obter posição da linha selecionada
line = 3
count = line
countSheetTarget = lastPositionFree
corVermelho = RGB(255, 0, 0)
corVerde = RGB(112, 173, 71)
corBranco = RGB(255, 255, 255)
corLaranja = RGB(255, 117, 11)
'planilhaBase position vars
bssBase = 1
bassunto = 2
bsite = 3
btipo = 4
bprio = 5
bstatusInterno = 15
bstatusAtual = 16
bdataAloc = 8
bdataInicio = 9
bdataEstimada = 10
bdataEnvio = 11
bdataAprov = 12
bgestor = 18
brecurso = 19
besforco = 20
'planilhaDestino position vars
dssBase = 1
dassunto = 2
dsite = 3
dtipo = 4
dprio = 5
dstatusInterno = 6
dstatusAtual = 7
dpendente = 8
ddataInicio = 10
desforco = 11
ddataEstimada = 13
dgestor = 16
drecurso = 18
ddataAloc = 19
ddataAprov = 21
ddataEnvio = 20
dfarol = 9
destimativa = 22
'variaveis de exceção posição
bcolExcep = 21
bcolStartExcep = 22
bcolEndExcep = 23
bcolComment = 24
dcolExcep = 28
dcolStartExcep = 29
dcolEndExcep = 30
dcolComment = 31
'
filaDev = Worksheets("Principal").Cells(24, 1)
implement = Worksheets("Principal").Cells(19, 1)

'procura celula vazia na planilha destino
While (Worksheets(planilhaDestino).Cells(countSheetTarget, 1) <> "")
    countSheetTarget = countSheetTarget + 1
Wend

'procura ultima linha livre sem numero de SS
With ActiveSheet
        LastRowEmpty = .Cells(.Rows.count, "A").End(xlUp).Row
End With
    
    
    'define cor da celula exibe status para verde
    Worksheets(planilhaBase).Cells(1, 6).Interior.Color = corLaranja
    Worksheets(planilhaBase).Cells(1, 6).Activate
While (Worksheets(planilhaBase).Cells(count, bssBase) <> "")
    'Debug.Print count
    Worksheets(planilhaBase).Cells(1, 6) = "Processando linha: " + CStr(count) + " / " + CStr(LastRowEmpty) 'Exibe status do processamento na linha 1 coluna F
    ssBase = Worksheets(planilhaBase).Cells(count, bssBase) 'col A
    assunto = Worksheets(planilhaBase).Cells(count, bassunto) 'col B
    site = Worksheets(planilhaBase).Cells(count, bsite) 'col C
    tipo = Worksheets(planilhaBase).Cells(count, btipo) 'col D
    prio = Worksheets(planilhaBase).Cells(count, bprio) 'col E
    statusInterno = Worksheets(planilhaBase).Cells(count, bstatusInterno) 'col O
        'Procura status na planilha principal na Guia de Status
        lineCount = 13
        Do While Worksheets("Principal").Cells(lineCount, 1) <> ""
            If statusInterno = Worksheets("Principal").Cells(lineCount, 1) Then
                pendente = Worksheets("Principal").Cells(lineCount, 3)
                lineCount = 13
                Exit Do
            Else
                pendente = "Status não encontrado"
            End If
        lineCount = lineCount + 1
        Loop
        
    'verifica se a coluna P contem um valor de data e obtem o valor
    If IsDate(Worksheets(planilhaBase).Cells(count, bstatusAtual)) Then
        statusAtual = Worksheets(planilhaBase).Cells(count, bstatusAtual) 'col P
    End If
    
    
    dataAloc = Worksheets(planilhaBase).Cells(count, bdataAloc) 'col H
    dataInicio = Worksheets(planilhaBase).Cells(count, bdataInicio) 'col I
    dataEstimada = Worksheets(planilhaBase).Cells(count, bdataEstimada) 'col J
    dataEnvio = Worksheets(planilhaBase).Cells(count, bdataEnvio) 'col K
    dataAprov = Worksheets(planilhaBase).Cells(count, bdataAprov) 'col L
    gestor = Worksheets(planilhaBase).Cells(count, bgestor) 'col R
    recurso = Worksheets(planilhaBase).Cells(count, brecurso) 'col S
    esforco = Worksheets(planilhaBase).Cells(count, besforco) 'col T
    
    'verifica se possui exceção marcada como e na coluna U/21, se true alimenta as datas
    excepMark = Worksheets(planilhaBase).Cells(count, bcolExcep)
    excepMark = UCase(excepMark)
    If (excepMark = "S") Then
        If (Worksheets(planilhaBase).Cells(count, bcolStartExcep) <> "" And Worksheets(planilhaBase).Cells(count, bcolEndExcep) <> "") Then
            excep = Worksheets(planilhaBase).Cells(count, bcolExcep)
            dtStartExcep = Worksheets(planilhaBase).Cells(count, bcolStartExcep)
            dtEndExcep = Worksheets(planilhaBase).Cells(count, bcolEndExcep)
            commentExcep = Worksheets(planilhaBase).Cells(count, bcolComment)
        End If
    End If

    If (Worksheets(planilhaDestino).Cells(countSheetTarget, dssBase) = "") Then
        'fill the sheet
        Worksheets(planilhaDestino).Cells(countSheetTarget, dssBase) = ssBase 'Nº SS - col A
        Worksheets(planilhaDestino).Cells(countSheetTarget, dassunto) = assunto 'ASSUNTO - col B
        Worksheets(planilhaDestino).Cells(countSheetTarget, dsite) = site 'SITE - col C
        Worksheets(planilhaDestino).Cells(countSheetTarget, dtipo) = tipo 'Tipo - col D
        Worksheets(planilhaDestino).Cells(countSheetTarget, dprio) = prio 'Prioridade - col E
        Worksheets(planilhaDestino).Cells(countSheetTarget, dstatusInterno) = statusInterno 'Status (interno) - col F
        Worksheets(planilhaDestino).Cells(countSheetTarget, dstatusAtual) = statusAtual 'Status (SSOL) - col G
        Worksheets(planilhaDestino).Cells(countSheetTarget, dpendente) = pendente 'status da tabelinha - col H
        'checa Data vazio
        If dataInicio <> "00:00:00" Then
            Worksheets(planilhaDestino).Cells(countSheetTarget, ddataInicio) = dataInicio 'Data Inicio - col J
        End If
        'checa Data vazio
        If esforco <> "00:00:00" Then
            Worksheets(planilhaDestino).Cells(countSheetTarget, desforco) = esforco 'Esforço - col K
        End If
        'checa Data vazio
        If dataEstimada <> "00:00:00" Then
            Worksheets(planilhaDestino).Cells(countSheetTarget, ddataEstimada) = dataEstimada 'Data Estimada / Data Fim - col M
        End If
        
        Worksheets(planilhaDestino).Cells(countSheetTarget, dgestor) = gestor 'Gestor - col P
        Worksheets(planilhaDestino).Cells(countSheetTarget, drecurso) = recurso 'Recurso - col R
        
        'checa Data vazio
        If (dataAloc <> "00:00:00") Then
            Worksheets(planilhaDestino).Cells(countSheetTarget, ddataAloc) = dataAloc 'Data alocada Scopus - col S
        End If
        'checa Data vazio
        If dataAprov <> "00:00:00" Then
            Worksheets(planilhaDestino).Cells(countSheetTarget, ddataAprov) = dataAprov 'Data alocada Scopus - Col U
        End If
        
        'movido logica para antes do Farol para funcionar a logica do Farol entendimento escopo vazio = vermelho
        'Entendimento Escopo = coluna j e se existe data = verde se está vazio = vermelho
        If dataEnvio <> "00:00:00" Then
            Worksheets(planilhaDestino).Cells(countSheetTarget, ddataEnvio) = dataEnvio 'col T
            Worksheets(planilhaDestino).Cells(countSheetTarget, ddataEnvio).Interior.Color = corVerde 'col T
        Else
        'Entendimento Escopo = coluna j se está vazio = vermelho
            Worksheets(planilhaDestino).Cells(countSheetTarget, ddataEnvio).Interior.Color = corVermelho 'col T
        End If
        'Farol = se data fim <=hoje -> vermelho
        If dataEstimada <= today Then
            Worksheets(planilhaDestino).Cells(countSheetTarget, dfarol).Interior.Color = corVermelho 'col I
        End If
        'Farol = se data fim >hoje -> verde
        If dataEstimada > today Then
            Worksheets(planilhaDestino).Cells(countSheetTarget, dfarol).Interior.Color = corVerde 'col I
        End If
        'Farol = se data (scopus backlog)+2>hoje && entendimento escopo vazio -> vermelho
        If ((dataAloc + CDate(2)) > today) And (Worksheets(planilhaDestino).Cells(countSheetTarget, 20) <> "") Then
            Worksheets(planilhaDestino).Cells(countSheetTarget, dfarol).Interior.Color = corVermelho 'col I
        End If
        'Farol = se data fim vazio && status interno =(fila dev ou implementação)->  vermelho
        If (dataEstimada = "00:00:00") And ((statusInterno = filaDev) Or (statusInterno = implement)) Then
            Worksheets(planilhaDestino).Cells(countSheetTarget, dfarol).Interior.Color = corVermelho 'col I
        End If
        'Scopus Backlog = coluna h verde se data diferente de vazio
        If Worksheets(planilhaDestino).Cells(countSheetTarget, ddataAloc) <> "" Then
            Worksheets(planilhaDestino).Cells(countSheetTarget, ddataAloc).Interior.Color = corVerde 'col S
        End If
        
        'Aguardando aprovação = coluna k e se existe data = verde
        If Worksheets(planilhaDestino).Cells(countSheetTarget, ddataAprov) <> "" Then
            Worksheets(planilhaDestino).Cells(countSheetTarget, ddataAprov).Interior.Color = corVerde 'col U
        End If
        'Aguardando aprovação = se celula vazia(sem data) e entendimento escopo tem data = vermelho
        If (Worksheets(planilhaDestino).Cells(countSheetTarget, ddataAprov) = "") And (Worksheets(planilhaDestino).Cells(countSheetTarget, ddataEnvio) <> "") Then
            Worksheets(planilhaDestino).Cells(countSheetTarget, ddataAprov).Interior.Color = corVermelho 'col U
        End If
        'Estimativa = se "data fim" preenchida = verde
        If Worksheets(planilhaDestino).Cells(countSheetTarget, ddataEstimada) <> "" Then
            Worksheets(planilhaDestino).Cells(countSheetTarget, destimativa).Interior.Color = corVermelho 'col V
        'se "data fim" em branco && aguardando aprovação em branco = branco
        ElseIf (Worksheets(planilhaDestino).Cells(countSheetTarget, ddataEstimada) = "") And (Worksheets(planilhaDestino).Cells(countSheetTarget, ddataAprov) = "") Then
            Worksheets(planilhaDestino).Cells(countSheetTarget, destimativa).Interior.Color = corBranco 'col V
        'senão = vermelha
        Else
            Worksheets(planilhaDestino).Cells(countSheetTarget, destimativa).Interior.Color = corVermelho 'col V
        End If
        'preenche colunas de exception se existir
        If (excep = "S") Then
            Worksheets(planilhaDestino).Cells(countSheetTarget, dcolExcep) = excep 'col AB/28
            Worksheets(planilhaDestino).Cells(countSheetTarget, dcolStartExcep) = dtStartExcep 'col AC/29
            Worksheets(planilhaDestino).Cells(countSheetTarget, dcolEndExcep) = dtEndExcep 'col AD/30
            Worksheets(planilhaDestino).Cells(countSheetTarget, dcolComment) = commentExcep 'col AE/31
        End If
        'zerar excep
        excep = ""
        'add count line sheet target
        countSheetTarget = countSheetTarget + 1
    Else
        MsgBox ("erro ao inserir")
    End If
    

    count = count + 1
Wend

Worksheets(planilhaBase).Cells(1, 6).Interior.Color = corBranco 'define cor da celula exibe status branco
Worksheets(planilhaBase).Cells(1, 6) = "" 'limpa status
Worksheets("config.ini").Cells(4, 2) = countSheetTarget


End Sub