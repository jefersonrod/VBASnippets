Sub Abre_share()


    ' Chama share remoto para o ip informado
    '
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
    
    If str_num_ip = "" Then MsgBox ("Preencher o campo IP antes de fazer acesso ao share"): End
    If str_num_ip = " " Then MsgBox ("Preencher o campo IP antes de fazer acesso ao share"): End
   
    ' "cmd /k " & Chr(34) &
    ' strcmd = "cmd /k " & Chr(34) & "runas /u:" & str_num_ip & "\suporte /savecred cmd /c" & Chr(34)
    
    '/ u& Chr(34) &:& Chr(34) & Chr(34) &"
    ' & str_num_ip & Chr(34) & \suporte /savecred cmd /k start \\ & Chr(34) & str_num_ip  & Chr(34) & \c\temp
    'strcmd = "This is the Test only." & Chr(34) & "double ok" & Chr(34) & "."
    
    ' abre o comando ping -t
    ' x = Shell("start" & " " & "\\" & str_num_ip & "\c\temp", vbNormalNoFocus)
    ' x = Shell("ping" & " " & str_num_ip & " " & "-t", vbNormalNoFocus)
     
     'x = Shell("ping" & " " & str_num_ip, vbNormalFocus)
     'MsgBox ("start \\" & str_num_ip)
      ' x = Shell("cmd /k runas /u:" & str_num_ip & "\suporte" & " " & "/savecred" & " cmd /k " & "start \\" & str_num_ip & "\c\temp", vbNormalFocus)
     
     'MsgBox (strcmd)
     'Z = Shell(strcmd, vbNormalFocus)
     
     'y = Shell("cmd /c time /t>c:\share-ini.txt")
	 
	 'executa o comando pelo cmd
     x = Shell("cmd /c start \\" & str_num_ip & "\c\temp", vbNormalFocus)
     
     'y = Shell("cmd /c time /t>c:\share-end.txt")
    
     
End Sub
