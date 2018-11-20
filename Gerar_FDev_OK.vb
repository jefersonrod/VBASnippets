Sub Gerar_FDev_OK()
'define vars
Dim stRecurso As String
Dim stSS As String
Dim stProjeto As String
Dim dtDataInicio As Date
Dim dtDataFim As Date
Dim linhaAtual As Integer
Dim planilhaStatus As String
Dim planilhaFilaDev As String
Dim planilhaIndisp As String
Dim diasEntreDatas As Long
Dim diasTotal As Long
Dim intDias As Long
Dim strDias As String
Dim strDiasTotal As String
Dim corAzul As Long
Dim corVerde As Long
Dim corCreme As Long
Dim corPreto As Long
Dim corBranco As Long
Dim corVermelho As Long
Dim varreLinha As Integer
Dim mesVarreMenor As Date
Dim mesVarreMaior As Date
Dim StartDate As Date
Dim EndDate As Date
Dim NoDays As Integer
Dim rng As Range
Dim cel As Range
Dim bssBase As Integer
Dim count As Integer
Dim countSheetTarget As Integer
Dim cores(1 To 3) As Long
'variaveis de exceção
Dim excep As String
Dim dtStartExcep As Date
Dim dtEndExcep As Date
Dim commentExcep As String
Dim dcolExcep As Integer
Dim dcolStartExcep As Integer
Dim dcolEndExcep As Integer
Dim dcolComment As Integer
'variaveis da aba indisponivel
Dim linhaRecurso As Integer
Dim colRecurso As Integer
Dim DevName As String
Dim colRecursoInterno As Integer
Dim recursoInterno As String

    
    'check se esta vazia a planilha
    If WorksheetFunction.CountA(Range("A3:A3")) = 0 Then
        MsgBox "Planilha não contem dados para analisar"
        Exit Sub
    End If
    
    
    'feed vars
    corAzul = RGB(217, 225, 242)
    corVerde = RGB(226, 239, 218)
    corCreme = RGB(255, 242, 204)
    corPreto = RGB(0, 0, 0)
    corBranco = RGB(255, 255, 255)
    corVermelho = RGB(255, 0, 0)
    nl = vbCrLf
    planilhaStatus = Worksheets("config.ini").Cells(3, 2)
    planilhaFilaDev = Worksheets("config.ini").Cells(5, 2)
    planilhaIndisp = Worksheets("config.ini").Cells(6, 2)
    
    cores(1) = corAzul
    cores(2) = corVerde
    cores(3) = corCreme
    'var exception
    bcolExcep = 28
    bcolStartExcep = 29
    bcolEndExcep = 30
    bcolComment = 31
    
    'verifica linha atual para montar dados no destino
    linhaAtual = 3
    varreLinha = linhaAtual
    
    'preenche menor mes e maior para evitar erro na primeira leitura de celula vazia
    If Worksheets(planilhaStatus).Cells(varreLinha, 10) = "" Then
        mesVarreMenor = "01/01/2099"
        mesVarreMaior = "01/01/1900"
    Else
        mesVarreMenor = Worksheets(planilhaStatus).Cells(varreLinha, 10)
    End If
    
    'varre celulas para procurar a menor data e a maior data para gerar colunas com datas
    While (Worksheets(planilhaStatus).Cells(varreLinha, 1) <> "")
        'menor data - col 10 - data inicio
        If (Worksheets(planilhaStatus).Cells(varreLinha, 10) <> "") Then
            'mes = CInt(Format(Worksheets(planilhaStatus).Cells(varreLinha, 10), "m"))
            mes = Worksheets(planilhaStatus).Cells(varreLinha, 10)
            If mes < mesVarreMenor Then
                mesVarreMenor = mes
            End If
            
        End If
        'maior data
        If (Worksheets(planilhaStatus).Cells(varreLinha, bcolEndExcep) <> "") Then
            'mes = CInt(Format(Worksheets(planilhaStatus).Cells(varreLinha, 10), "m"))
            mes = Worksheets(planilhaStatus).Cells(varreLinha, bcolEndExcep)
            
            If mes > mesVarreMaior Then
                mesVarreMaior = mes
            End If
        End If

        varreLinha = varreLinha + 1
    Wend
    
            
    
    'MsgBox (CStr(mesVarreMenor))
    'MsgBox (CStr(mesVarreMaior))
    
    'preenche calendario com as datas obtidas
    StartDate = mesVarreMenor
    EndDate = mesVarreMaior
    NoDays = EndDate - StartDate + 1
  
    Worksheets(planilhaFilaDev).Range("B2").Value = StartDate

    'fill lines
    'Range("A3").Resize(NoDays).DataSeries Rowcol:=xlColumns, Type:=xlChronological, Date:=xlDay, Step:=1, Stop:=EndDate, Trend:=False
    'fill columns
    Worksheets(planilhaFilaDev).Range("B2").Resize(NoDays).DataSeries Rowcol:=xlRows, Type:=xlChronological, Date:=xlDay, Step:=1, Stop:=EndDate, Trend:=False
      
    'VARRE TAB INDISPONIVEL PARA BUSCAR DEVS
    linhaRecurso = 2
    colRecurso = 1
    countSheetTarget = 3
    
    While (Worksheets(planilhaIndisp).Cells(linhaRecurso, colRecurso) <> "")
        DevName = Worksheets(planilhaIndisp).Cells(linhaRecurso, colRecurso)
        'MsgBox (DevName)
        'Varre planilha StatusSS para buscar SS
        count = 3
        
        bssBase = 1
        colRecursoInterno = 18
        While (Worksheets(planilhaStatus).Cells(count, bssBase) <> "")
            recursoInterno = Worksheets(planilhaStatus).Cells(count, colRecursoInterno)
            If (DevName = recursoInterno) Then
                
                
                'vars
                stRecurso = Worksheets(planilhaStatus).Cells(count, 18) 'Recurso
                stSS = Worksheets(planilhaStatus).Cells(count, bssBase) 'Numero SS
                'stProjeto = Worksheets(planilhaStatus).Cells(count, 3) 'Projeto
                dtDataInicio = Worksheets(planilhaStatus).Cells(count, 10) 'Data Inicio
                dtDataFim = Worksheets(planilhaStatus).Cells(count, 13) 'Data Fim
                
                diasEntreDatas = DateDiff("d", dtDataInicio, dtDataFim) 'numero de dias entre datas
                strDias = CStr(diasEntreDatas)
                intDias = CLng(strDias)
                diasTotal = DateAdd("d", intDias, dtDataInicio) ' numero total de dias
                strDiasTotal = CStr(intDias + 1)
                'verifica se possui exceção marcada como e na coluna U/21, se true alimenta as datas
                excep = Worksheets(planilhaStatus).Cells(count, bcolExcep)
                excep = UCase(excep)
                If (excep = "S") Then
                    If (Worksheets(planilhaStatus).Cells(count, bcolStartExcep) <> "" And Worksheets(planilhaStatus).Cells(count, bcolEndExcep) <> "") Then
                        dtStartExcep = Worksheets(planilhaStatus).Cells(count, bcolStartExcep)
                        dtEndExcep = Worksheets(planilhaStatus).Cells(count, bcolEndExcep)
                        commentExcep = Worksheets(planilhaStatus).Cells(count, bcolComment)
                    End If
                    
                End If
                
               
                
                
                'Debug.Print recursoInterno & " : " & stSS & " Ini: " & dtDataInicio & " dias: " & strDias & "|" & countSheetTarget
                
                If (dtDataInicio = "00:00:00" Or dtDataFim = "00:00:00") Then
                    'nothing to do
                ElseIf (dtDataInicio <> "00:00:00" And dtDataFim <> "00:00:00") Then
                    
                    Debug.Print recursoInterno & " : " & stSS & " Ini: " & dtDataInicio & " dias: " & strDias & "|" & countSheetTarget & "/" & excep
                    
                    'procura ultima coluna preenchida com data para popular o rangeTotalCalendario
                    With Worksheets(planilhaFilaDev)
                        'LastCol = Worksheets("Fila Dev").Cells(2, .Columns.count).End(xlToLeft).Column
                        LastColAddress = Worksheets(planilhaFilaDev).Cells(2, .Columns.count).End(xlToLeft).Address
                        'LastColValor = Worksheets(planilhaFilaDev).Cells(2, .Columns.count).End(xlToLeft).Value
                        lastAddressColDate = Worksheets(planilhaFilaDev).Range(LastColAddress).Address
                    End With
            
                    rangeTotalCalendario = "B2:" + lastAddressColDate
                    'MsgBox (rangeTotalCalendario)
                    
                    Set rng = Worksheets(planilhaFilaDev).Range(rangeTotalCalendario)
                    For Each cel In rng.Cells
                    
                        With cel
                            
                            If dtDataInicio = .Value Then
                                'MsgBox ("Found")
                                'Debug.Print .Address & ":" & .Value
                                address_found = .Address
                            End If
                            
                        End With
                
                    Next cel
            
                    'coluna da data inicial encontrada
                    address_found = Range(address_found & 1).Column
                    'MsgBox (address_found)
                    
                    intInicio = CInt(address_found) 'Data inicial
                    final = intInicio + CInt(intDias) 'Data de inicio + Dias para desenv
                    'linhaFilaDev = 3
                    
                    'preenche dados iniciais
                    Worksheets(planilhaFilaDev).Cells(countSheetTarget, 1) = stRecurso & " | " & stSS
                    'Worksheets(planilhaFilaDev).Cells(countSheetTarget, 2) = stProjeto
                    'Worksheets(planilhaFilaDev).Cells(countSheetTarget, 1) = stSS
                    'Worksheets(planilhaFilaDev).Cells(countSheetTarget, 3).Interior.Color = corBranco
                    Worksheets(planilhaFilaDev).Cells(countSheetTarget, 1).Font.Color = corPreto
                    
                    
                    
                    'preenche inicio do projeto e quantidade de dias com cor diferente
                    Worksheets(planilhaFilaDev).Cells(countSheetTarget, intInicio).NumberFormat = "0" 'format celula como texto
                    Worksheets(planilhaFilaDev).Cells(countSheetTarget, intInicio) = strDiasTotal 'qtde de dias
                    Worksheets(planilhaFilaDev).Cells(countSheetTarget, intInicio).Interior.Color = corPreto
                    Worksheets(planilhaFilaDev).Cells(countSheetTarget, intInicio).Font.Color = corBranco
                    Worksheets(planilhaFilaDev).Cells(countSheetTarget, intInicio + 1).NumberFormat = "0" 'format celula como texto
                    Worksheets(planilhaFilaDev).Cells(countSheetTarget, intInicio + 1) = stSS ' numero da SS
                    Randomize
                    cor = cores(Int((3 - 1 + 1) * Rnd + 1))
                    For inicial = intInicio + 1 To final Step 1 'intInicio + 1 para pular a qtde de dias
                        Randomize
                        'Worksheets(planilhaFilaDev).Cells(linhaFilaDev, inicial) = "X"
                        Worksheets(planilhaFilaDev).Cells(countSheetTarget, inicial).Interior.Color = cor
                        If inicial = final Then
                            Worksheets(planilhaFilaDev).Cells(countSheetTarget, inicial).NumberFormat = "0" 'format celula como texto
                            Worksheets(planilhaFilaDev).Cells(countSheetTarget, inicial) = stSS ' numero da SS
                        End If
                        
                    Next inicial
                    
                    '************ EXCEPTION
                    'se for excep preenche barra de calendario
                    If (excep = "S") Then
                        dtDataInicio = dtStartExcep
                        dtDataFim = dtEndExcep
                        diasEntreDatas = DateDiff("d", dtDataInicio, dtDataFim) 'numero de dias entre datas
                        strDias = CStr(diasEntreDatas)
                        intDias = CLng(strDias)
                        diasTotal = DateAdd("d", intDias, dtDataInicio) ' numero total de dias
                        strDiasTotal = CStr(intDias + 1)
                        Debug.Print "S" & recursoInterno & " : " & stSS & " Ini: " & dtDataInicio & " dias: " & strDias & "|" & countSheetTarget & "*"
                        
                        'procura ultima coluna preenchida com data para popular o rangeTotalCalendario
                        With Worksheets(planilhaFilaDev)
                            LastColAddress = Worksheets(planilhaFilaDev).Cells(2, .Columns.count).End(xlToLeft).Address
                            lastAddressColDate = Worksheets(planilhaFilaDev).Range(LastColAddress).Address
                        End With
                
                        rangeTotalCalendario = "B2:" + lastAddressColDate
                        'MsgBox (rangeTotalCalendario)
                        
                        Set rng = Worksheets(planilhaFilaDev).Range(rangeTotalCalendario)
                        For Each cel In rng.Cells
                        
                            With cel
                                
                                If dtDataInicio = .Value Then
                                    'MsgBox ("Found")
                                    'Debug.Print .Address & ":" & .Value
                                    address_found = .Address
                                End If
                                
                            End With
                    
                        Next cel
                
                        'coluna da data inicial encontrada
                        address_found = Range(address_found & 1).Column
                        'MsgBox (address_found)
                        
                        intInicio = CInt(address_found) 'Data inicial
                        final = intInicio + CInt(intDias) 'Data de inicio + Dias para desenv
                        'linhaFilaDev = 3
                        
                        'preenche dados iniciais
                        'Worksheets(planilhaFilaDev).Cells(countSheetTarget, 1) = stRecurso & " | " & stSS
                        'Worksheets(planilhaFilaDev).Cells(countSheetTarget, 2) = stProjeto
                        'Worksheets(planilhaFilaDev).Cells(countSheetTarget, 1) = stSS
                        'Worksheets(planilhaFilaDev).Cells(countSheetTarget, 3).Interior.Color = corBranco
                        'Worksheets(planilhaFilaDev).Cells(countSheetTarget, 1).Font.Color = corPreto
                        
                        
                        
                        'preenche inicio do projeto e quantidade de dias com cor diferente
                        Worksheets(planilhaFilaDev).Cells(countSheetTarget, intInicio).NumberFormat = "0" 'format celula como texto
                        Worksheets(planilhaFilaDev).Cells(countSheetTarget, intInicio) = strDiasTotal 'qtde de dias
                        Worksheets(planilhaFilaDev).Cells(countSheetTarget, intInicio).Interior.Color = corVermelho
                        Worksheets(planilhaFilaDev).Cells(countSheetTarget, intInicio).Font.Color = corBranco
                        Worksheets(planilhaFilaDev).Cells(countSheetTarget, intInicio + 1).NumberFormat = "0" 'format celula como texto
                        Worksheets(planilhaFilaDev).Cells(countSheetTarget, intInicio + 1) = stSS ' numero da SS
                        Randomize
                        cor = cores(Int((3 - 1 + 1) * Rnd + 1))
                        For inicial = intInicio + 1 To final Step 1 'intInicio + 1 para pular a qtde de dias
                            Randomize
                            'Worksheets(planilhaFilaDev).Cells(linhaFilaDev, inicial) = "X"
                            Worksheets(planilhaFilaDev).Cells(countSheetTarget, inicial).Interior.Color = cor
                            If inicial = final Then
                                Worksheets(planilhaFilaDev).Cells(countSheetTarget, inicial).NumberFormat = "0" 'format celula como texto
                                Worksheets(planilhaFilaDev).Cells(countSheetTarget, inicial) = stSS ' numero da SS
                            End If
                            
                        Next inicial
                    
                    
                    End If
                    
                    
                    
                    
                    
                    'incrementa linha da planilha destino
                    countSheetTarget = countSheetTarget + 1
                
                End If
                    
                
                
            End If
            
        
        'ultima linha do loop/incremento
        count = count + 1
        Wend
    'ultima linha do loop/incremento
    linhaRecurso = linhaRecurso + 1
    Wend
     
    
    
    
    
    
    
    'verifica se tem data inicial e data final
'    If (dtDataInicio = "00:00:00" Or dtDataFim = "00:00:00") Then
'        MsgBox ("Verifique a data inicial e a data final" + nl + "Valor vazio encontrado" + nl + "Inicial: " + CStr(dtDataInicio) + nl + "Final: " + CStr(dtDataFim))
'        Exit Sub
'    End If
    
    
        
    


End Sub
