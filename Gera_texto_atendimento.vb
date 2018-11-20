
Sub Gera_texto_atendimento()

' ----------------------------------------------------------------------
' Função: gerar arquivo texto com dados da planilha para bloco de notas
' Desenvolvido por: Jeferson Rodrigues - Analista Suporte - VSAT
' Data: 16/07/2013
' Hora: 14:09
'
' ----------------------------------------------------------------------

    Dim Arquivo As String
    Dim campo(4)
    Dim addr_vlr As String
    Dim DataCham As String
    
    Dim Arqv As String
    Dim Drive As String
    Dim FullArqv As String
    
    
        
    ' informa o endereço da celula selecionada para a variavel
    ' (linha, coluna)
    addr_vlr = Application.ActiveCell.Address
    
    ' Variaveis não usadas no momento
    '
    ' Cli = Worksheets(1).Cells(1, 3)
    ' Saud = Worksheets(1).Cells(1, 5)
    ' Resp = Worksheets(1).Cells(2, 5)
    ' Encer = Worksheets(1).Cells(3, 5)
    
    ' define o nome do arquivo do corpo do e-mail
    ' Arqv = "e-mail" & "_" & Cli & ".txt"
     Arqv = "Atendimento.txt"
    
    ' define o caminho completo do arquivo de e-mail
    ' ******* Surgiu problema com função shell ao abrir o arquivo usando variavel, verificar depois
     Drive = "C:\temp\"
     FullArqv = Drive + Arqv
    
    ' Este bloco todo tem erro na logica, verificar depois com calma
    ' converte o endereço em valor da linha
    'If Len(addr_vlr) = 4 Then
    'addr_lin = Int(Right(addr_vlr, 1))
    'Else: If Len(addr_vlr) = 5 Then Else
    'addr_lin = Int(Right(addr_vlr, 2)): If Len(addr_vlr) = 6 Then Else
    'addr_lin = Int(Right(addr_vlr, 3))
    'End If
    ' aqui termina o bloco com erro, substituido por case abaixo
    
    ' Verifica a linha onde esta o cursor
    
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
    
    '---------------------------------------------------------
    ' informações somente para localização
    ' Application.ActiveCell = addr_lin
    ' Application.ActiveCell = Application.ActiveCell.Address
    ' --------------------------------------------------------
    
    ' alimenta variaveis com coordenadas de linha e coluna
    atend_data = Worksheets(1).Cells(addr_lin, 1)
    atend_horario = Worksheets(1).Cells(addr_lin, 2)
    atend_loja = Worksheets(1).Cells(addr_lin, 3)
    atend_tecnico = Worksheets(1).Cells(addr_lin, 4)
    
    txt_problema = Worksheets(1).Cells(addr_lin, 8)
    txt_solucao = Worksheets(1).Cells(addr_lin, 9)
    
    ' verifica se o campo de horario foi preenchido
    
    If atend_data = "" Then MsgBox ("Preencher o campo Data antes de gerar atendimento"): End
    If atend_horario = "" Then MsgBox ("Preencher o campo Horario antes de gerar atendimento"): End
    If atend_loja = "" Then MsgBox ("Preencher o campo numero da loja antes de gerar atendimento"): End
    If atend_tecnico = "" Then MsgBox ("Preencher o campo nome do técnico antes de gerar atendimento"): End
    If txt_problema = "" Then MsgBox ("Preencher o campo Problema relatado antes de gerar atendimento"): End
    If txt_solucao = "" Then MsgBox ("Preencher o campo Solução antes de gerar atendimento"): End
        
    ' converte horario em Cdate e pega somente a parte hh:mm
    maldita_var = CDate(atend_horario)
    maldita_var = Left(maldita_var, 5)
    
    ' alimenta variaveis de campo com valores da planilha
    
    campo_atend_data = "Data: " & atend_data
    campo_atend_horario = "Horario: " & maldita_var
    
    campo_atend_loja = "Loja: " & atend_loja
    campo_atend_tecnico = "Técnico: " & atend_tecnico
    
    campo_txt_problema = "Problema: " & txt_problema
    campo_txt_solucao = "Solução: " & txt_solucao
    
    ' usado somente para e-mail
    ' campo_assunto = alert_serv_name & " - " & Hour(alert_horario) & ":" & Minute(alert_horario) & " - " & alert_data & " - " & txt_problema
    
    ' cria arquivo de texto e joga os dados coletados
    Open FullArqv For Output As #1
        ' Print #1, campo_assunto
        ' Print #1, " "
        ' Print #1, " "
        ' Print #1, Saud & " " & Resp
        ' Print #1, " "
        ' Print #1, "Favor verificar o servidor abaixo"
        ' Print #1, " "
        Print #1, campo_atend_data
        Print #1, campo_atend_horario
        Print #1, campo_atend_loja
        Print #1, campo_atend_tecnico
        ' Print #1, "Obs.: "
        'Print #1, "Mensagem Original: "
        Print #1, " "
        Print #1, campo_txt_problema
        Print #1, campo_txt_solucao
        ' Print #1, " "
        ' Print #1, " "
        
      
        
    ' fecha o arquivo de texto
    Close #1
    
    ' grava na coluna reg info que o registro foi gerado
    Worksheets(1).Cells(addr_lin, 15) = "X"
    
    
    ' abre o bloco de notas com o texto criado
    x = Shell("notepad" & " " & FullArqv, vbNormalNoFocus)
    
    ' seleciona e copia para area de transferencia o conteudo do bloco de notas
      AppActivate ("Notepad")
      SendKeys "^a"
      SendKeys "^c"
      
    'tempo de pausa
    Application.Wait (Now + TimeValue("0:00:01"))
      
    ' Abre checklist para gerar atendimento
    'y = Shell("C:\Program Files\Internet Explorer\iexplore http://checklist.redeboticario.vsat/admin/AtendimentoMigracoes/AtendimentoMigracoesBusca.cshtml", vbNormalFocus)
    'Windows("Checklist Migrações 2013 - Windows Internet Explorer provided by Grupo Boticario").Activate
    
    
    ' envia tab ate chegar no campo da loja
    
    'SendKeys "{TAB 5}"
    
    'tempo de pausa
    'Application.Wait (Now + TimeValue("0:00:01"))
    
    ' envia numero da loja
    'SendKeys atend_loja
    
    ' insere a data e hora atual e notificação E-mail ao enviar

    ' Worksheets(1).Cells(addr_lin, 5) = Date
    ' Worksheets(1).Cells(addr_lin, 6) = Time
    ' Worksheets(1).Cells(addr_lin, 7) = "E-mail"
    
    ' salva o arquivo
    Workbooks(Application.ActiveWorkbook.Name).Save
    
End Sub
