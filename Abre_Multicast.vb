Sub Abre_Multicast()
    ' Chama share remoto para o ip informado
    '
    '
    Dim str_num_ip As String
    Dim plan As String
	Dim linhaAtual as Integer

    ' Verifica a linha onde esta o cursor
    linhaAtual = linha_Atual.linha_Atual 'chama função para obter a linha atual
    plan = ActualSheetName 'obtem o nome da planilha atual
    
    If (Worksheets(plan).Cells(linhaAtual, 11) = "") Then
        MsgBox ("Preencher o campo IP antes de fazer acesso ao share")
    Else
       ' Lê posição onde esta o numero do IP
        str_num_ip = Worksheets(plan).Cells(addr_lin, 11)
        x = Shell("cmd /c start \\" & str_num_ip & "\E\Multicast", vbNormalFocus)
    End If
     
End Sub


