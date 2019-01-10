Sub FormCheckout()
'get info from sheet and generate card in Trello
On Error Resume Next
On Error GoTo 0

'set general vars
Dim plan As String
Dim corpo As String

Dim loja As String
Dim lojaCodigo As String
Dim line As String
Dim checkin As String
Dim checkout As String
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
plan = FunctionsTimeModelX.ActualSheetName
nl = vbCrLf 'new line
line = "---------------------------------"
checkin = "*Check-In*"
checkout = "*Check-Out*"

    'Check loja field
    If (Worksheets(plan).Cells(5, 3) = "") Then
        MsgBox ("Favor preencher o numero da loja na linha 5")
        Worksheets(plan).Cells(5, 3).Select
        Exit Sub
        
    Else
        'check tecnico answer
        If (Worksheets(plan).Cells(9, 3) = "") Then
            MsgBox ("Favor preencher a linha 9")
            Worksheets(plan).Cells(9, 3).Select
            Exit Sub
        End If
        
        
        'check responsavel field
        For i = 13 To 17 Step 2
            If (Worksheets(plan).Cells(i, 3) = "") Then
            MsgBox ("Favor preencher a linha " + CStr(i))
            Worksheets(plan).Cells(i, 3).Select
            Exit Sub
            End If
        Next i
        
        loja = Worksheets(plan).Cells(4, 3)
        lojaCodigo = Worksheets(plan).Cells(5, 3)
        'Técnico
        tecnico = Worksheets(plan).Cells(7, 3)
        aPrimesysDeuasOrientacoes = Worksheets(plan).Cells(8, 3)
        aPrimesysDeuasOrientacoesRESP = Worksheets(plan).Cells(9, 3)
        'Responsavel
        responsavel = Worksheets(plan).Cells(11, 3)
        casoHajaumProblemaOcorrido = Worksheets(plan).Cells(12, 3)
        casoHajaumProblemaOcorridoRESP = Worksheets(plan).Cells(13, 3)
        solicitarUmaAvaliacaodoTecnico = Worksheets(plan).Cells(14, 3)
        solicitarUmaAvaliacaodoTecnicoRESP = Worksheets(plan).Cells(15, 3)
        seHouverMaisItens = Worksheets(plan).Cells(16, 3)
        seHouverMaisItensRESP = Worksheets(plan).Cells(17, 3)
        
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
            
    
    'generate content
    CopyText corpo 'call copy to clipboard function
    
    'change color of the loja (colunm 1) to mark as sent
    'Worksheets(plan).Cells(linhaAtual, colLoja).Font.Color = corAzul


End Sub
