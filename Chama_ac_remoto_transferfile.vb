Sub Chama_ac_remoto_transferfile()

    ' Chama acesso remoto para o ip informado
    '
    ' isscntr /a<address> /c<command> /l /s<core server>
    '
    Dim str_num_ip As String
    Dim plan As String
    Dim linhaAtual As Integer
    
    ' Verifica a linha onde esta o cursor
    linhaAtual = linha_Atual.linha_Atual 'chama função para obter a linha atual
    plan = ActualSheetName 'obtem o nome da planilha atual
    
    ' Lê posição onde esta o numero do IP
    str_num_ip = Worksheets(plan).Cells(linhaAtual, 11)
    
   ' verifica se o campo de end ip foi preenchido
    
    If str_num_ip = "" Then MsgBox ("Preencher o campo IP antes de fazer acesso remoto"): End
   
   
    ' desabilita o forefront
    ' Z = Shell("C:\Program Files\Forefront TMG Client\fwctool.exe disable", vbHide)
    
    ' abre o bloco de notas com o texto criado
    'x = Shell("C:\Program Files\LANDesk\ManagementSuite\isscntr.exe" & " " & "/a" & str_num_ip & " " & "/crunas /u:suporte cmd", vbNormalNoFocus)
    x = Shell("C:\Program Files\LANDesk\ManagementSuite\isscntr.exe" & " " & "/a" & str_num_ip & " " & "/c""file transfer""", vbNormalNoFocus)
End Sub
