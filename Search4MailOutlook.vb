Sub Search4MailOutlook()

'set variable for use with Outlook
Dim myOlApp As New Outlook.Application
Dim myNameSpace As Outlook.Namespace
Dim myInbox As Outlook.MAPIFolder
Dim myitems As Outlook.Items
Dim olFldr As Outlook.MAPIFolder
Dim myitem As Object
Dim olkAtt As Outlook.Attachment
Dim aMail As MailItem
Dim aAttach As Attachment

'set variable for general use in WorkSheet
Dim plan As String
Dim procurar As String
Dim extensao As String
Dim linhaAtual As Integer
Dim caminhoSalvarRelat As String
Dim corVermelho As Long
Dim mailTotalCount As String
Dim mailItemCount As Integer
Dim mailFoundCount As Integer
Dim pastaCxEntrada As String
Dim log As Boolean
Dim timercron As Boolean
Dim StartTime As Double
Dim SecondsElapsed As Double

Dim Found As Boolean
Dim showEMail As Boolean


'set variable for log output in text
Dim Arqv As String
Dim ArqvLog As String
Dim Drive As String
Dim FullArqv As String
Dim FullArqvLog As String

'set positions variable
Dim colLoja As Integer
Dim colRelatOK As Integer
Dim linLoja As Integer
Dim colCaminhoSalvarRelat As Integer
Dim linCaminhoSalvarRelat As Integer
Dim colmailItemCount As Integer
Dim linmailItemCount As Integer
Dim colmailTotalCount As Integer
Dim linmailTotalCount As Integer
Dim colFolderSearch As Integer
Dim linFolderSearch As Integer
Dim colShowEmail As Integer
Dim linShowEmail As Integer
Dim colExtensao As Integer
Dim linExtensao As Integer
Dim colLog As Integer
Dim linLog As Integer
Dim colTimer As Integer
Dim linTimer As Integer
Dim colTimerShow As Integer
Dim linTimerShow As Integer

'set variables for path and filename to save attachment
Dim sPath As String
Dim sName As String
Dim sFile As String



'set variables for use with OutLook
Set myNameSpace = myOlApp.GetNamespace("MAPI")
Set myInbox = myNameSpace.GetDefaultFolder(olFolderInbox)
'Set myitems = myInbox.Items

'feed general variables
plan = "Buscar"
Found = False
linLoja = FunctionsBuscaRelatFotog.linha_Atual
corVermelho = RGB(255, 0, 0)

'feed position variables
colLoja = 1
colRelatOK = 2
colCaminhoSalvarRelat = 2
linCaminhoSalvarRelat = 1
colmailItemCount = 5
linmailItemCount = 1
colmailTotalCount = 7
linmailTotalCount = 1
linFolderSearch = 2
colFolderSearch = 2
linShowEmail = 3
mailItemCount = 0
colShowEmail = 2
linExtensao = 4
colExtensao = 2
colLog = 2
linLog = 5
colTimer = 2
linTimer = 6
colTimerShow = 4
linTimerShow = linLoja
'feed general variables with defined positions
log = Worksheets("config.ini").Cells(linLog, colLog)
timercron = Worksheets("config.ini").Cells(linLog, colLog)
'start timer if enable
If (timercron) Then
    StartTime = Timer 'Remember time when macro starts
    time1 = Time
End If
procurar = Worksheets(plan).Cells(linLoja, colLoja)
caminhoSalvarRelat = Worksheets("config.ini").Cells(linCaminhoSalvarRelat, colCaminhoSalvarRelat)
pastaCxEntrada = Worksheets("config.ini").Cells(linFolderSearch, colFolderSearch)
showEMail = Worksheets("config.ini").Cells(linShowEmail, colShowEmail)
extensao = Worksheets("config.ini").Cells(linExtensao, colExtensao)
sPath = caminhoSalvarRelat
sName = "Relatório Fotografico Loja " + procurar + extensao
sFile = sPath & sName

'set Outlook variables after get inbox name
Set olFldr = myInbox.Folders(pastaCxEntrada)
Set myitems = olFldr.Items
mailTotalCount = olFldr.Items.count
Worksheets(plan).Cells(linmailTotalCount, colmailTotalCount) = mailTotalCount

'set path for developers test text log
ArqvLog = "logSubject.txt"
Drive = "C:\temp\"
FullArqvLog = Drive + ArqvLog
'Open FullArqvLog For Append As #1

'check for some fields are filled
If (caminhoSalvarRelat = "") Then
    MsgBox ("Preencher o caminho para salvar o anexo na aba config.ini, ex.: c:\temp\")
    Exit Sub
End If
If (pastaCxEntrada = "") Then
    MsgBox ("Preencher o nome da pasta da caixa de entrada na aba config.ini")
    Exit Sub
End If
If (extensao = "") Then
    MsgBox ("Preencher a extensão do arquivo buscado na aba config.ini, ex.: .doc")
    Exit Sub
End If

'read each email in inbox defined
For Each myitem In myitems
    If myitem.Class = olMail Then
        Set aMail = myitem
        mailItemCount = mailItemCount + 1 'count item read
        Worksheets(plan).Cells(linmailItemCount, colmailItemCount) = mailItemCount 'update item count
        If InStr(1, myitem.subject, procurar) > 0 Then 'if subject is equal procurar returns 1, found
        mailFoundCount = mailItemCount
            For Each aAttach In aMail.Attachments 'search in attachments
                If Right(LCase(aAttach.Filename), Len(extensao)) = extensao Then 'check  attachment extension
                Debug.Print "Found: " + procurar + " | " + myitem.subject 'dev debug show item
                If (showEMail) Then 'check is shows email found or not
                    myitem.Display
                End If
                Found = True
                aAttach.SaveAsFile sFile 'save attachment
                Worksheets(plan).Cells(linLoja, colRelatOK) = "OK" 'setr status in worksheet
                If (log) Then 'check if log or not
                    Call LogBusca.LogBusca(Found, mailItemCount, procurar, myitem.subject)
                End If
                If (timercron) Then
                    time2 = Time
                    time3 = DateDiff("s", time1, time2)
                    Worksheets(plan).Cells(linTimerShow, colTimerShow) = CStr(time3)
                    Worksheets(plan).Cells(linTimerShow, colTimerShow + 1) = "segundos executando"
                End If
                'exit from search
                Exit For
                End If
           Next
        End If
    End If
Next myitem
Worksheets(plan).Cells(linmailItemCount, colmailItemCount) = mailFoundCount 'update final count item found
'Close #1 'close text log file

'If the subject isn't found:
If Not Found Then
    'NoResults.Show
    Worksheets(plan).Cells(linLoja, colRelatOK) = "NO" 'update status on worksheet
    Worksheets(plan).Cells(linLoja, colRelatOK).Font.Color = corVermelho 'update status on worksheet
    Debug.Print "NOT Found " + procurar 'dev debug show item status
    If (log) Then 'check if log or not
       Call LogBusca.LogBusca(Found, mailItemCount, procurar, "Não encontrado / Not found")
    End If
    If (timercron) Then
        time2 = Time
        time3 = DateDiff("s", time1, time2)
        Worksheets(plan).Cells(linTimerShow, colTimerShow) = CStr(time3)
        Worksheets(plan).Cells(linTimerShow, colTimerShow + 1) = "segundos executando"
    End If
End If

'myOlApp.Quit 'close outlook
Set myOlApp = Nothing
  
End Sub
