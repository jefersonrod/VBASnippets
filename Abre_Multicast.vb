Sub Abre_Multicast()
    ' Chama share remoto para o ip informado
    '
    '
    Dim str_num_ip As String
    Dim plan As String
    Dim linhaAtual As Integer

    ' Verifica a linha onde esta o cursor
    linhaAtual = linha_Atual.linha_Atual 'chama função para obter a linha atual
    plan = ActualSheetName 'obtem o nome da planilha atual
    
    'verifica se o campo do ip esta vazio
    If (Worksheets(plan).Cells(linhaAtual, 11) = "") Then
        MsgBox ("Preencher o campo IP antes de fazer acesso ao share")
    Else
       ' armazena o numero ip
        str_num_ip = Worksheets(plan).Cells(linhaAtual, 11)
        'executa o comando via cmd
        x = Shell("cmd /c start \\" & str_num_ip & "\E\Multicast", vbNormalFocus)
    End If
     
End Sub


