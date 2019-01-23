Sub FormCheckout()
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
Dim usuarioAtual As String
Dim ivan As String
Dim jeferson As String
Dim luiz As String
Dim rener As String
Dim thiago As String
'Técnico
Dim tecnico As String
Dim aPrimesysDeuasOrientacoes As String
Dim aPrimesysDeuasOrientacoesRESP As String
'Responsavel
Dim resposavel As String
Dim casoHajaumProblemaOcorrido As String
Dim casoHajaumProblemaOcorridoRESP As String
Dim solicitarUmaAvaliacaodoTecnico As String
Dim solicitarUmaAvaliacaodoTecnicoRESP As String
Dim seHouverMaisItens As String
Dim seHouverMaisItensRESP As String

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
checkout = "*Check-Out*"


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
        If (Worksheets(plan).Cells(9, colNum) = "") Then
            MsgBox ("Favor preencher a linha 9")
            Worksheets(plan).Cells(9, colNum).Select
            Exit Sub
        End If
        
        
        'check responsavel field
        For i = 13 To 17 Step 2
            If (Worksheets(plan).Cells(i, colNum) = "") Then
            MsgBox ("Favor preencher a linha " + CStr(i))
            Worksheets(plan).Cells(i, colNum).Select
            Exit Sub
            End If
        Next i
        
        loja = Worksheets(plan).Cells(4, colNum)
        lojaCodigo = Worksheets(plan).Cells(5, colNum)
        'Técnico
        tecnico = Worksheets(plan).Cells(7, colNum)
        aPrimesysDeuasOrientacoes = Worksheets(plan).Cells(8, colNum)
        aPrimesysDeuasOrientacoesRESP = Worksheets(plan).Cells(9, colNum)
        'Responsavel
        responsavel = Worksheets(plan).Cells(11, colNum)
        casoHajaumProblemaOcorrido = Worksheets(plan).Cells(12, colNum)
        casoHajaumProblemaOcorridoRESP = Worksheets(plan).Cells(13, colNum)
        solicitarUmaAvaliacaodoTecnico = Worksheets(plan).Cells(14, colNum)
        solicitarUmaAvaliacaodoTecnicoRESP = Worksheets(plan).Cells(15, colNum)
        seHouverMaisItens = Worksheets(plan).Cells(16, colNum)
        seHouverMaisItensRESP = Worksheets(plan).Cells(17, colNum)
        
    End If
    
    'generate content text
    
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
            
    registrado = FunctionsTimeModelX.Username
    Call CheckLogOUT.CheckLogOUT(analista, lojaCodigo, aPrimesysDeuasOrientacoesRESP, casoHajaumProblemaOcorridoRESP, solicitarUmaAvaliacaodoTecnicoRESP, seHouverMaisItensRESP, registrado)
    'generate content
    CopyText corpo 'call copy to clipboard function
    
    'change color of the loja (colunm 1) to mark as sent
    'Worksheets(plan).Cells(linhaAtual, colLoja).Font.Color = corAzul


End Sub
