Sub Abre_share_job_install_tefcontingencia()


    ' Chama share remoto para o ip informado
    '
    '
    Dim str_num_ip As String
    Dim pathjob As String
    Dim addr_vlr As String
    Dim cmdshell As String
    
    
    
    
    ' Seta variavel com posição atual do cursor
    
    addr_vlr = Application.ActiveCell.Address
    pathjob = "\c\Program Files\LANDesk\LDClient\sdmcache\pacotel\G_AdministracaoOPVSAT\Infraestrutura\Aplicacoes\TEFContingenciaF5\Rollout"
    cmdshell = "cmd /c start \\"
    
    ' Verifica a linha onde esta o cursor
    ' alimenta variaveis com coordenadas de linha e coluna
    Select Case Len(addr_vlr)
    
    Case 4
    addr_lin = Int(Right(addr_vlr, 1))
    Case 5
    addr_lin = Int(Right(addr_vlr, 2))
    Case 6
    addr_lin = Int(Right(addr_vlr, 3))
    Case 7
    addr_lin = Int(Right(addr_vlr, 4))
    
    End Select
    
    ' Lê posição onde esta o numero do IP
    str_num_ip = Worksheets(1).Cells(addr_lin, 11)
    
    ' verifica se o campo de end ip foi preenchido
    
    If str_num_ip = "" Then MsgBox ("Preencher o campo IP antes de fazer acesso ao rollout"): End
    If str_num_ip = " " Then MsgBox ("Preencher o campo IP antes de fazer acesso ao rollout"): End
    
    x = Shell(cmdshell & str_num_ip & pathjob, vbNormalFocus)
     
    'y = Shell("cmd /c time /t>c:\share-end.txt")
    
     
End Sub


