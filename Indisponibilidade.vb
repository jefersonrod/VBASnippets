Sub Indisponibilidade()
    Dim stRecurso As String
    Dim dtDataInicioIndisp As Date
    Dim dtDataInicio As Date
    Dim planilhaFilaDev As String
    Dim planilhaIndisp As String
    Dim dias As Integer
    Dim strDias As String
    Dim corVermelho As Long
    Dim corLaranja As Long
    Dim corBranco As Long
    Dim corPreto As Long
    Dim StartDate As Date
    Dim EndDate As Date
    Dim NoDays As Integer
    Dim rng As Range
    Dim cel As Range
    Dim count As Integer
    Dim countSheetTarget As Integer
    Dim varreLinhaFilaDev As Integer
    Dim varreLinhaIndisponivel As Integer
    Dim ultimaLinhaVaziaFilaDev As Integer
    Dim nomeFilaDev As String
    Dim nomeIndisponivel As String
    Dim colDataInicio As Integer
    Dim colDev As Integer
    Dim colSS As Integer
    
    
    'feed vars
    corVermelho = RGB(255, 0, 0)
    corLaranja = RGB(255, 117, 11)
    corBranco = RGB(255, 255, 255)
    corPreto = RGB(0, 0, 0)
    planilhaIndisp = Worksheets("config.ini").Cells(6, 2)
    planilhaFilaDev = Worksheets("config.ini").Cells(5, 2)
    
    colDev = 1
    colSS = 3
    
    'procura ultima celula vazia na Fila Dev para preencher a indisponibilidade
    colDataInicio = 2
    varreLinhaFilaDev = 3
    While (Worksheets(planilhaFilaDev).Cells(varreLinhaFilaDev, colSS) <> "")
        ultimaLinhaVaziaFilaDev = varreLinhaFilaDev
    varreLinhaFilaDev = varreLinhaFilaDev + 1
    Wend
    'MsgBox (ultimaLinhaVaziaFilaDev)
    'procura nome na aba indisponivel e na aba fila dev, verifica o nome e adiciona o periodo de indisponibilidade
    varreLinhaIndisponivel = 2
    varreLinhaFilaDev = 3
    ultimaLinhaVaziaFilaDev = ultimaLinhaVaziaFilaDev + 1
    While (Worksheets(planilhaIndisp).Cells(varreLinhaIndisponivel, 1) <> "")
    varreLinhaFilaDev = 3
        nomeIndisponivel = Worksheets(planilhaIndisp).Cells(varreLinhaIndisponivel, 1)
        Do While (Worksheets(planilhaFilaDev).Cells(varreLinhaFilaDev, 1) <> "")
            nomeFilaDev = Worksheets(planilhaFilaDev).Cells(varreLinhaFilaDev, 1)
            'varreLinhaFilaDev = varreLinhaFilaDev + 1
            If nomeIndisponivel = nomeFilaDev Then
                Worksheets(planilhaFilaDev).Cells(ultimaLinhaVaziaFilaDev, 1) = nomeIndisponivel
                'procura data de inicio da indisponibilidade
                Do While Worksheets(planilhaIndisp).Cells(varreLinhaIndisponivel, colDataInicio) <> "00:00:00"
                    If Worksheets(planilhaIndisp).Cells(varreLinhaIndisponivel, colDataInicio) Then
                        dtDataInicioIndisp = Worksheets(planilhaIndisp).Cells(varreLinhaIndisponivel, colDataInicio)
                        dias = Worksheets(planilhaIndisp).Cells(varreLinhaIndisponivel, colDataInicio + 1)
                        Exit Do
                    End If
                colDataInicio = colDataInicio + 1
                Loop
                colDataInicio = 2 'retorna para numero inicial para proximo loop
                'MsgBox (CStr(dtDataInicioIndisp))
                dtDataInicio = dtDataInicioIndisp
                Set rng = Worksheets(planilhaFilaDev).Range("D2:CN2")
                For Each cel In rng.Cells
                
                    With cel
                        
                        If dtDataInicio = .Value Then
                            'MsgBox ("Found")
                            'Debug.Print .Address & ":" & .Value
                            address_found = .Address
                        End If
                        
                    End With
            
                Next cel
                
                If address_found <> "" Then
                    'coluna da data inicial encontrada
                    address_found = Range(address_found & 1).Column
                    'MsgBox (address_found)
                    intInicio = CInt(address_found) 'Data inicial
                    Worksheets(planilhaFilaDev).Cells(ultimaLinhaVaziaFilaDev, 2) = "Inicio Indisponibilidade: "
                    Worksheets(planilhaFilaDev).Cells(ultimaLinhaVaziaFilaDev, 2).Font.Color = corPreto
                    Worksheets(planilhaFilaDev).Cells(ultimaLinhaVaziaFilaDev, 3).NumberFormat = "@" 'format celula como texto
                    Worksheets(planilhaFilaDev).Cells(ultimaLinhaVaziaFilaDev, 3) = CStr(dtDataInicioIndisp)
                    Worksheets(planilhaFilaDev).Cells(ultimaLinhaVaziaFilaDev, 3).Font.Color = corPreto
                    Worksheets(planilhaFilaDev).Cells(ultimaLinhaVaziaFilaDev, intInicio).NumberFormat = "@" 'format celula como texto
                    Worksheets(planilhaFilaDev).Cells(ultimaLinhaVaziaFilaDev, intInicio) = dias
                    Worksheets(planilhaFilaDev).Cells(ultimaLinhaVaziaFilaDev, intInicio).Interior.Color = corPreto
                    Worksheets(planilhaFilaDev).Cells(ultimaLinhaVaziaFilaDev, intInicio).Font.Color = corBranco
                    For inicial = intInicio + 1 To intInicio + dias Step 1 'intInicio + 1 para pular a qtde de dias
                    'Worksheets(planilhaFilaDev).Cells(linhaFilaDev, inicial) = "X"
                        Worksheets(planilhaFilaDev).Cells(ultimaLinhaVaziaFilaDev, inicial).Interior.Color = corLaranja
                    Next inicial
                    'Worksheets(planilhaFilaDev).Cells(ultimaLinhaVaziaFilaDev, 3) = CStr(dtDataInicioIndisp)
                Else
                    'Data fora do range atual
                    Worksheets(planilhaFilaDev).Cells(ultimaLinhaVaziaFilaDev, 2) = "Data fora da faixa: "
                    Worksheets(planilhaFilaDev).Cells(ultimaLinhaVaziaFilaDev, 3).NumberFormat = "@" 'format celula como texto
                    Worksheets(planilhaFilaDev).Cells(ultimaLinhaVaziaFilaDev, 3) = CStr(dtDataInicioIndisp)
                    Worksheets(planilhaFilaDev).Cells(ultimaLinhaVaziaFilaDev, 2).Font.Color = corBranco
                    Worksheets(planilhaFilaDev).Cells(ultimaLinhaVaziaFilaDev, 2).Interior.Color = corLaranja
                    Worksheets(planilhaFilaDev).Cells(ultimaLinhaVaziaFilaDev, 3).Font.Color = corBranco
                    Worksheets(planilhaFilaDev).Cells(ultimaLinhaVaziaFilaDev, 3).Interior.Color = corLaranja
                End If
                address_found = ""
                ultimaLinhaVaziaFilaDev = ultimaLinhaVaziaFilaDev + 1
                varreLinhaFilaDev = ultimaLinhaVaziaFilaDev 'aumenta varreLinhaFilaDev para sair do loop Do While e recomeçar do inicio
            End If
            varreLinhaFilaDev = varreLinhaFilaDev + 1
        Loop
    varreLinhaIndisponivel = varreLinhaIndisponivel + 1
    Wend
End Sub
