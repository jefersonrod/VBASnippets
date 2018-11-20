Sub Mail_Chamado_Bematech()

' ----------------------------------------------------------------------
' Função: gerar arquivo texto com dados da planilha para corpo de mail
' Desenvolvido por: Jeferson Rodrigues - Analista Suporte - VSAT
' Data: 16/07/2013
' Hora: 15:28
'
' ----------------------------------------------------------------------

    Dim Arquivo As String
    Dim campo(4)
    Dim addr_vlr As String
    Dim DataCham As String
    Dim Saud As String
    Dim Solic As String
    Dim Encer As String
    Dim campo_assunto As String
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
     Arqv = "Bematech.txt"
    
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
    chamad_cod_loja = Worksheets(1).Cells(addr_lin, 3)
    
    'txt_serv_tag = Worksheets(1).Cells(addr_lin, 12)
    'txt_defeito = Worksheets(1).Cells(addr_lin, 15)
    
   ' Alimenta variaveis de texto com valores sring
    campo_assunto = "Reparo de equipamento com defeito"
    
    Saud = "Olá"
    Solic = "Peço a gentileza de verificar o erro em anexo que esta ocorrendo na loja "
    Encer = "Fico no aguardo de retorno e agradeço pela atenção."
    subj = "Erro na loja "
    
    ' verifica se o campo de horario foi preenchido
    
    If chamad_cod_loja = "" Then MsgBox ("Preencher o campo NUMERO DA LOJA antes de gerar atendimento"): End
    
    ' If txt_serv_tag = "" Then MsgBox ("Preencher o campo SERVICE TAG antes de gerar atendimento"): End
    ' If txt_defeito = "" Then MsgBox ("Preencher o campo DEFEITO antes de gerar atendimento"): End
        
    ' converte horario em Cdate e pega somente a parte hh:mm
    ' maldita_var = CDate(atend_horario)
    ' maldita_var = Left(maldita_var, 5)
    
    ' alimenta variaveis de campo com valores da planilha
    
    ' campo_atend_data = "Data: " & atend_data
    ' campo_atend_horario = "Horario: " & maldita_var
    ' campo_atend_tecnico = "Técnico: " & atend_tecnico
    
    campo_chamad_cod_loja = "Código da loja: " & chamad_cod_loja
    ' campo_txt_serv_tag = "Service Tag: " & txt_serv_tag
    ' campo_txt_defeito = "Defeito: " & txt_defeito
    
    ' usado somente para e-mail
    campo_assunto = subj & chamad_cod_loja
    
    ' cria arquivo de texto e joga os dados coletados
    Open FullArqv For Output As #1
        Print #1, campo_assunto
        Print #1, " "
        Print #1, " "
        Print #1, Saud
        Print #1, " "
        Print #1, " "
        Print #1, Solic & chamad_cod_loja
        ' Print #1, "Favor verificar o servidor abaixo"
        Print #1, " "
        Print #1, " "
        'Print #1, campo_chamad_cod_loja
        'Print #1, campo_txt_serv_tag
        'Print #1, campo_txt_defeito
        ' Print #1, campo_atend_tecnico
        ' Print #1, "Obs.: "
        'Print #1, "Mensagem Original: "
        Print #1, " "
        ' Print #1, campo_txt_problema
        'Print #1, campo_txt_solucao
        ' Print #1, " "
        Print #1, Encer
        
      
        
    ' fecha o arquivo de texto
    Close #1
    
    ' abre o bloco de notas com o texto criado
    x = Shell("notepad" & " " & FullArqv, vbNormalNoFocus)
    
    ' seleciona e copia para area de transferencia o conteudo do bloco de notas
      AppActivate ("Notepad")
      SendKeys "^a"
      SendKeys "^c"
      
      ' teste
      'AppActivate ("Outlook")
      'SendKeys "^n"
    
    ' insere a data e hora atual e notificação E-mail ao enviar

    ' Worksheets(1).Cells(addr_lin, 5) = Date
    ' Worksheets(1).Cells(addr_lin, 6) = Time
    ' Worksheets(1).Cells(addr_lin, 7) = "E-mail"
    
    ' salva o arquivo
    ' Workbooks(Application.ActiveWorkbook.Name).Save ' re-ligado para verificar erros do excel / 24/12/11

End Sub

