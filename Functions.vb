Option Explicit
Public Function CopyText(Text As String)
    'VBA Macro using late binding to copy text to clipboard.
    'By Justin Kay, 8/15/2014
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    MSForms_DataObject.SetText Text
    MSForms_DataObject.PutInClipboard
    Set MSForms_DataObject = Nothing
End Function

Public Function Username()
    Dim name As String
    name = Mid(Application.Username, 9)
    MsgBox "Current user is " & name
End Function

Public Function CreateStatusHTML()
    Dim linhaatual As Integer
    Dim colData As Integer
    Dim colHora As Integer
    Dim colLoja As Integer
    
    Dim data As String
    Dim hora As String
    Dim loja As String
    Dim nome As String
    
    Dim Arqv As String
    Dim ArqvLog As String
    Dim Drive As String
    Dim FullArqv As String
    Dim FullArqvLog As String
    
    Dim divstart As String
    Dim divend As String
    Dim br As String
    
    Dim id As String
    
    divstart = "<div id=""col1""> "
    divend = "</div>"
    br = "<br>"
    id = Mid(Application.Username, 9)
    
    ' Arqv = "e-mail" & "_" & Cli & ".txt"
    Arqv = "status1.html"
    ArqvLog = "status1-log.html"
    ' define o caminho completo do arquivo de e-mail
    ' ******* Surgiu problema com função shell ao abrir o arquivo usando variavel, verificar depois
    Drive = "C:\Users\jefersonr\OneDrive - Grupo Boticario\Util\src\html\statusatend\"
    FullArqv = Drive + Arqv
    FullArqvLog = Drive + ArqvLog
    
    colData = 1
    colHora = 2
    colLoja = 3
    
    'get atual
    linhaatual = linha_Atual.linha_Atual
    
    loja = Worksheets("Atendimentos").Cells(linhaatual, colLoja)
    data = Worksheets("Atendimentos").Cells(linhaatual, colData)
    hora = Format(Worksheets("Atendimentos").Cells(linhaatual, colHora), "hh:mm")
    
    loja = "<h2>" + loja + "</h2>"
    
    'write file as html
    ' cria arquivo de texto e joga os dados coletados
    Open FullArqv For Output As #1
        Print #1, " "
        Print #1, divstart
        Print #1, loja
        Print #1, br
        Print #1, data
        Print #1, br
        Print #1, hora
        Print #1, br
        Print #1, id
        Print #1, divend
    Close #1
    
    ' cria arquivo de texto e joga os dados coletados
    Open FullArqvLog For Append As #1
        Print #1, " "
        Print #1, divstart
        Print #1, loja
        Print #1, br
        Print #1, data
        Print #1, br
        Print #1, hora
        Print #1, br
        Print #1, id
        Print #1, divend
    Close #1

End Function

Public Function EmptyStatusHTML()
    Dim linhaatual As Integer
    Dim colData As Integer
    Dim colHora As Integer
    Dim colLoja As Integer
    
    Dim data As String
    Dim hora As String
    Dim loja As String
    Dim nome As String
    
    Dim Arqv As String
    Dim ArqvLog As String
    Dim Drive As String
    Dim FullArqv As String
    Dim FullArqvLog As String
    
    Dim divstart As String
    Dim divend As String
    Dim br As String
    
    Dim id As String
    
    divstart = "<div id=""col1""> "
    divend = "</div>"
    br = "<br>"
    id = Mid(Application.Username, 9)
    
    ' Arqv = "e-mail" & "_" & Cli & ".txt"
    Arqv = "status1.html"
    ArqvLog = "status1-log.html"
    ' define o caminho completo do arquivo de e-mail
    ' ******* Surgiu problema com função shell ao abrir o arquivo usando variavel, verificar depois
    Drive = "C:\Users\jefersonr\OneDrive - Grupo Boticario\Util\src\html\statusatend\"
    FullArqv = Drive + Arqv
    FullArqvLog = Drive + ArqvLog
    
    colData = 1
    colHora = 13
    'colLoja = 3
    
    'get atual
    linhaatual = linha_Atual.linha_Atual
    
    'loja = Worksheets("Atendimentos").Cells(linhaatual, colLoja)
    'data = Worksheets("Atendimentos").Cells(linhaatual, colData)
    hora = Format(Worksheets("Atendimentos").Cells(linhaatual, colHora), "hh:mm")
    
    loja = "<h2>" + loja + "</h2>"
    
    'write file as html
    ' cria arquivo de texto e joga os dados coletados
    Open FullArqv For Output As #1
        Print #1, " "
        Print #1, divstart
        Print #1, "DISPONIVEL"
        Print #1, br
        Print #1, id
        Print #1, divend
    Close #1
    

End Function

