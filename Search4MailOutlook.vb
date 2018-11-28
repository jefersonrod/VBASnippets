Sub Search4MailOutlook()

Dim myOlApp As New Outlook.Application
Dim myNameSpace As Outlook.Namespace
Dim myInbox As Outlook.MAPIFolder
Dim myitems As Outlook.Items
Dim olFldr As Outlook.MAPIFolder
Dim myitem As Object
Dim Found As Boolean
Dim procurar As String
Dim olkAtt As Outlook.Attachment
Dim aMail As MailItem
Dim aAttach As Attachment

Dim Arqv As String
Dim ArqvLog As String
Dim Drive As String
Dim FullArqv As String
Dim FullArqvLog As String
Dim linhaAtual As Integer
Dim colLoja As Integer
Dim colRelatOK As Integer
Dim linLoja As Integer
Dim colCaminhoSalvarRelat As Integer
Dim linCaminhoSalvarRelat As Integer
Dim caminhoSalvarRelat As String
Dim sPath As String
Dim sName As String
Dim sFile As String
Dim corVermelho As Long
Dim mailTotalCount As String
Dim mailItemCount As Integer
Dim colmailItemCount As Integer
Dim linmailItemCount As Integer
Dim colmailTotalCount As Integer
Dim linmailTotalCount As Integer
Dim mailFoundCount As Integer
Dim colFolderSearch As Integer
Dim linFolderSearch As Integer
Dim pastaCxEntrada As String

Set myNameSpace = myOlApp.GetNamespace("MAPI")
Set myInbox = myNameSpace.GetDefaultFolder(olFolderInbox)

'Set myitems = myInbox.Items


Found = False
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
mailItemCount = 0
linLoja = Functions.linha_Atual
corVermelho = RGB(255, 0, 0)
procurar = Worksheets(1).Cells(linLoja, colLoja)
caminhoSalvarRelat = Worksheets("config.ini").Cells(linCaminhoSalvarRelat, colCaminhoSalvarRelat)
pastaCxEntrada = Worksheets("config.ini").Cells(linFolderSearch, colFolderSearch)
sPath = caminhoSalvarRelat
sName = "Relatório Fotografico Loja " + procurar + ".doc"
sFile = sPath & sName

Set olFldr = myInbox.Folders(pastaCxEntrada)
Set myitems = olFldr.Items
mailTotalCount = olFldr.Items.Count
'set path
ArqvLog = "logSubject.txt"
Drive = "C:\temp\"
FullArqvLog = Drive + ArqvLog
'MsgBox (mailTotalCount)
Worksheets(1).Cells(linmailTotalCount, colmailTotalCount) = mailTotalCount
'Open FullArqvLog For Append As #1

For Each myitem In myitems
    If myitem.Class = olMail Then
        Set aMail = myitem
        'Debug.Print CStr(myitem.subject)
        'Print #1, CStr(myitem.subject)
        mailItemCount = mailItemCount + 1
        Worksheets(1).Cells(linmailItemCount, colmailItemCount) = mailItemCount
        If InStr(1, myitem.subject, procurar) > 0 Then
        mailFoundCount = mailItemCount
            For Each aAttach In aMail.Attachments
                If Right(LCase(aAttach.Filename), 4) = ".doc" Then
                Debug.Print "Found"
                myitem.Display
                Found = True
                aAttach.SaveAsFile sFile
                Worksheets(1).Cells(linLoja, colRelatOK) = "OK"
                'No need to check any of this message's remaining attachments
                
                Exit For
                End If
           Next
        End If
    End If
Next myitem
Worksheets(1).Cells(linmailItemCount, colmailItemCount) = mailFoundCount
'Close #1

'If the subject isn't found:
If Not Found Then
    'NoResults.Show
    Worksheets(1).Cells(linLoja, colRelatOK) = "NO"
    Worksheets(1).Cells(linLoja, colRelatOK).Font.Color = corVermelho
    Debug.Print "NOT Found"
End If

'myOlApp.Quit
Set myOlApp = Nothing
  
End Sub
