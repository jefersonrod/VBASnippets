Sub Ping_eterno()

    ' Efetua ping continuo no ip informado
    '
    '
    Dim str_num_ip As String
    Dim plan As String
    Dim linhaAtual As Integer
    
    ' Seta variavel com posição atual do cursor
    
    linhaAtual = linha_Atual.linha_Atual 'chama função para obter a linha atual
    plan = ActualSheetName 'obtem o nome da planilha atual
    ' Lê posição onde esta o numero do IP
    str_num_ip = Worksheets(plan).Cells(linhaAtual, 11)
    
    ' Verifica a linha onde esta o cursor
    ' verifica se o campo de end ip foi preenchido
    If (Worksheets(plan).Cells(linhaAtual, 11) = "") Then
        MsgBox ("Preencher o campo IP antes de fazer acesso remoto")
    Else
       ' abre o comando ping -t
        x = Shell("ping" & " " & str_num_ip & " " & "-t", vbNormalNoFocus)
    End If
    
End Sub
