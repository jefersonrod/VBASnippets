

Sub Abre_share_teste()


    ' Chama share remoto para o ip informado
    '
    '
    Dim str_num_ip As String
    Dim pathjob As String
    Dim pathjob2 As String
    Dim addr_vlr As String
    Dim cmdshell As String
    Dim pathunc As String
    
    
    
    
    ' Seta variavel com posição atual do cursor
    
    addr_vlr = Application.ActiveCell.Address
    pathjob = "\c\Program Files\LANDesk\LDClient"
    'pathjob = "c:\Program Files\LANDesk\LDClient\sdmcache"
    pathunc = "\\"
    cmdshell = "cmd /c start "
    
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
    'str_num_ip = Worksheets(1).Cells(addr_lin, 11)
    str_num_ip = "10.42.36.123"
    
    ' verifica se o campo de end ip foi preenchido
    
    If str_num_ip = "" Then MsgBox ("Preencher o campo IP antes de fazer acesso ao rollout"): End
    If str_num_ip = " " Then MsgBox ("Preencher o campo IP antes de fazer acesso ao rollout"): End
    
    'MsgBox (cmdshell & Chr(34) & pathunc & str_num_ip & pathjob & Chr(34))
    x = Shell(cmdshell & """\\""" & str_num_ip & """\c\Program Files\LANDesk\LDClient""", vbNormalFocus)
    'x = Shell("""cmd /c start \\"" & str_num_ip & ""\c\Program Files\LANDesk\LDClient""", vbNormalFocus)
    'x = Shell("cmd /c start \\" & str_num_ip & "\c\SUPERDB\Governanca\PracticoLive", vbNormalFocus)
    'x = Shell("C:\Program Files\Internet Explorer\iexplore.exe" & " " & "http://" & str_num_ip & ":16992", vbNormalNoFocus)
     
    'MsgBox (cmdshell & str_num_ip & pathjob & Chr(32) & pathjob2)
     
    'Start "c:\Program Files\LANDesk\LDClient\sdmcache\"
    
     
End Sub



