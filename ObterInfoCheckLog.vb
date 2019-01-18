Sub ObterInfoCheckLog()
Dim strFilename As String: strFilename = "C:\temp\trellotemp.txt"
Dim strFileContent As String
Dim iFile As Integer: iFile = FreeFile
Dim Ar() As String
Dim s As Variant
Dim chin As String
Dim chout As String
'set general vars
Dim plan As String
Dim nl As String
Dim data As String
Dim hora As String
Dim analista As String
Dim registrado As String
Dim colNum As Integer
Dim colSelect As String
Dim loja As String
Dim lojaCodigo As String
Dim line As String
Dim checkin As String
Dim checkout As String
Dim tipo As String
Dim linha As Integer
Dim pos As Integer
Dim Recebeucontato As Boolean
Dim Recebeuorienta As Boolean
Dim realizoumig As Boolean
Dim PossuiWhatsapp As Boolean
Dim Informarsobreolink As Boolean
Dim Enviofotos As Boolean
Dim sobreacompanhamento As Boolean
Dim algumchamado As Boolean
Dim problemasist As Boolean
Dim assinarOS As Boolean
Dim numerodotelefone As Boolean
Dim analNome As Boolean
Dim deuorienta As Boolean
Dim momentodamig As Boolean
Dim avaliacao As Boolean
Dim houvermaisitens As Boolean

'set vars for position
'columns
Dim colLoja As Integer
Dim colData As Integer
Dim colHora As Integer
Dim colTipo As Integer
'feed position vars
colData = 2
colHora = 3
colTipo = 4
colAnalistacolT = 20
colRegistradocolU = 21

'feed vars
plan = "Coletado"
nl = vbCrLf 'new line
chin = "Check-In"
chout = "Check-Out"
linha = ultimaLinhaLivreChkLog.ultimaLinhaLivreChkLog
Recebeucontato = False
analNome = True

Open strFilename For Input As #iFile
strFileContent = Input(LOF(iFile), iFile)
Close #iFile

Ar = Split(strFileContent, vbCrLf)
'read string line by line
For Each s In Ar
' s is the variable to process
     'Debug.Print s
    If (analNome) Then
                
        If (InStr(s, "ivan") > 0) Then
            analista = "Ivan"
        ElseIf ((InStr(s, "jeferson") > 0)) Then
            analista = "Jeferson"
        ElseIf ((InStr(s, "luiz") > 0)) Then
            analista = "Luiz"
        ElseIf ((InStr(s, "rener") > 0)) Then
            analista = "Rener"
        ElseIf ((InStr(s, "thiago") > 0)) Then
            analista = "Thiago"
        Else
            analista = "???"
        End If
        Debug.Print analista
        analNome = False
    End If
    If (InStr(s, "jan") > 0) Then
        pos = InStr(s, "jan") + 1
        data = Mid(LTrim(s), 1, pos)
        data = Replace(data, " de ", "/") + "/2019"
        pos = InStr(s, ":")
        hora = Mid(s, pos - 2, 5)
        Debug.Print data
        Debug.Print hora
    End If
    
    If (s = chin) Then
        GoTo processarCheckIN:
    End If
    If (s = chout) Then
        GoTo processarCheckOUT:
    End If
        
    
     
Next

processarCheckIN:
'Técnico
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


Worksheets(plan).Cells(linha, colData) = data
Worksheets(plan).Cells(linha, colHora) = hora
Worksheets(plan).Cells(linha, colTipo) = "IN"
Worksheets(plan).Cells(linha, colAnalistacolT) = analista
'read string line by line
For Each s In Ar
' s is the variable to process
    'Loja
    If (Left(s, 5) = "Loja:") Then
        Debug.Print "CheckIN" + s
        pos = InStr(s, ":")
        s = Mid(s, pos + 1)
        Worksheets(plan).Cells(linha, colLoja) = s
    End If
    'colE Recebeu contato da Primesys?
    If (Recebeucontato) Then
        Worksheets(plan).Cells(linha, colCINTecL8colE) = LTrim(s)
        Debug.Print s
        Recebeucontato = False
    End If
    If (InStr(s, "Recebeu contato") > 0) Then
        Recebeucontato = True
    End If
    'colF Recebeu orientações sobre o manual de migração?
    If (Recebeuorienta) Then
        Worksheets(plan).Cells(linha, colCINTecL10colF) = LTrim(s)
        Debug.Print s
        Recebeuorienta = False
    End If
    If (InStr(s, "Recebeu orienta") > 0) Then
        Recebeuorienta = True
    End If
    'colG Já realizou migração?
    If (realizoumig) Then
        Worksheets(plan).Cells(linha, colCINTecL12colG) = LTrim(s)
        Debug.Print s
        realizoumig = False
    End If
    If (InStr(s, "realizou migr") > 0) Then
        realizoumig = True
    End If
    'colH Possui Whatsapp? Qual?
    If (PossuiWhatsapp) Then
        Worksheets(plan).Cells(linha, colCINTecL14colH) = LTrim(s)
        Debug.Print s
        PossuiWhatsapp = False
    End If
    If (InStr(s, "Possui Whatsapp") > 0) Then
        PossuiWhatsapp = True
    End If
    'colI Informar sobre o link que está sendo instalado ou migrado para a nova solução.
    If (Informarsobreolink) Then
        Worksheets(plan).Cells(linha, colCINTecL16colI) = LTrim(s)
        Debug.Print s
        Informarsobreolink = False
    End If
    If (InStr(s, "sobre o link") > 0) Then
        Informarsobreolink = True
    End If
    'colJ Envio fotos rack, retaguarda, balcão, cabeamentos (balcão e PDV’s)
    If (Enviofotos) Then
        Worksheets(plan).Cells(linha, colCINTecL18colJ) = LTrim(s)
        Debug.Print s
        Enviofotos = False
    End If
    If (InStr(s, "Envio fotos") > 0) Then
        Enviofotos = True
    End If
    'colK Informar sobre acompanhamento da equipe boticário.
    If (sobreacompanhamento) Then
        Worksheets(plan).Cells(linha, colCINResL22colK) = LTrim(s)
        Debug.Print s
        sobreacompanhamento = False
    End If
    If (InStr(s, "Envio fotos") > 0) Then
        sobreacompanhamento = True
    End If
    'colL Tem algum chamado aberto?
    If (algumchamado) Then
        Worksheets(plan).Cells(linha, colCINResL24colL) = LTrim(s)
        Debug.Print s
        algumchamado = False
    End If
    If (InStr(s, "Envio fotos") > 0) Then
        algumchamado = True
    End If
    'colM Está com algum problema sistêmico ou em equipamentos?
    If (problemasist) Then
        Worksheets(plan).Cells(linha, colCINResL26colM) = fnConverterUTF8.fnConverterUTF8(LTrim(s))
        Debug.Print s
        problemasist = False
    End If
    If (InStr(s, "problema sist") > 0) Then
        problemasist = True
    End If
    'colN Orientar assinar OS somente após todos os testes.
    If (assinarOS) Then
        Worksheets(plan).Cells(linha, colCINResL28colN) = fnConverterUTF8.fnConverterUTF8(LTrim(s))
        Debug.Print s
        assinarOS = False
    End If
    If (InStr(s, "problema sist") > 0) Then
        assinarOS = True
    End If
    'colO Confirmar o numero do telefone (fixo ou celular da loja)
    If (numerodotelefone) Then
        Worksheets(plan).Cells(linha, colCINResL30colO) = fnConverterUTF8.fnConverterUTF8(LTrim(s))
        Debug.Print s
        numerodotelefone = False
    End If
    If (InStr(s, "numero do telefone") > 0) Then
        numerodotelefone = True
    End If
    
Next

Exit Sub
processarCheckOUT:
Debug.Print "CheckOUT found"
'columns

Dim colCOUTTecL8colP As Integer
Dim colCOUTResL12colQ As Integer
Dim colCOUTResL14colR As Integer
Dim colCOUTResL16colS As Integer


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

Worksheets(plan).Cells(linha, colData) = data
Worksheets(plan).Cells(linha, colHora) = hora
Worksheets(plan).Cells(linha, colTipo) = "OUT"
Worksheets(plan).Cells(linha, colAnalistacolT) = analista
'read string line by line
For Each s In Ar
' s is the variable to process
    'Loja
    If (Left(s, 5) = "Loja:") Then
        Debug.Print "CheckOUT" + s
        pos = InStr(s, ":")
        s = Mid(s, pos + 1)
        Worksheets(plan).Cells(linha, colLoja) = s
    End If
    'colP A Primesys deu as orientações (duvidas e problemas) no processo de migração?
    If (deuorienta) Then
        Worksheets(plan).Cells(linha, colCOUTTecL8colP) = LTrim(s)
        Debug.Print s
        deuorienta = False
    End If
    If (InStr(s, "deu as orienta") > 0) Then
        deuorienta = True
    End If
    'colQ Caso haja um problema ocorrido no momento da migração, que não está relacionado com a
    If (momentodamig) Then
        Worksheets(plan).Cells(linha, colCOUTResL12colQ) = LTrim(s)
        Debug.Print s
        momentodamig = False
    End If
    If (InStr(s, "no momento da migra") > 0) Then
        momentodamig = True
    End If
    'colR Solicitar uma avaliação do técnico de 1 a 5
    If (avaliacao) Then
        Worksheets(plan).Cells(linha, colCOUTResL14colR) = LTrim(s)
        Debug.Print s
        avaliacao = False
    End If
    If (InStr(s, "uma avalia") > 0) Then
        avaliacao = True
    End If
    'colS Se houver mais itens, por favor, podem informar.
    If (houvermaisitens) Then
        Worksheets(plan).Cells(linha, colCOUTResL16colS) = LTrim(s)
        Debug.Print s
        houvermaisitens = False
    End If
    If (InStr(s, "houver mais itens") > 0) Then
        houvermaisitens = True
    End If
    
Next
Exit Sub
End Sub
