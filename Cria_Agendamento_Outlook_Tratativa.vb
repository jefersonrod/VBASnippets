Sub Cria_Agendamento_Outlook_Tratativa()

Dim olApp As Outlook.Application
Dim olAppItem As Outlook.AppointmentItem

On Error Resume Next
On Error GoTo 0

Dim plan As String
Dim email As String
Dim trello As String
Dim assunto As String
Dim corpo As String
Dim tag As String
Dim emaildest As String

Dim linhaAtual As Integer
Dim colLoja As Integer
Dim colCP As Integer
Dim colCidade As Integer
Dim colUF As Integer
Dim colStatusMX As Integer
Dim colSituacao As Integer
Dim colPossibilidade As Integer
Dim colResponsavel As Integer
Dim colEmailDest As Integer
Dim colStatusAcao As Integer
Dim colDataRetorno As Integer
Dim colHoraRetorno As Integer

Dim loja As String
Dim cp As String
Dim cidade As String
Dim uf As String
Dim statusmx As String
Dim situacao As String
Dim possibilidade As String
Dim responsavel As String
Dim statusacao As String
Dim dataretorno As Date
Dim horaretorno As Date

Dim corAzulMarinho As String



nl = vbCrLf 'new line
br = "<br>" 'new line HTML
corAzulMarinho = RGB(68, 114, 196)
plan = "Tratativas"
tag = "#tratativas"

Worksheets(plan).Activate
'instance of Appointment
Set olApp = New Outlook.Application
'Set olApp = GetObject("", "Outlook.Application")
    'On Error GoTo 0
    'If olApp Is Nothing Then
     '   On Error Resume Next
     '   Set olApp = CreateObject("Outlook.Application")
     '   On Error GoTo 0
     '   If olApp Is Nothing Then
     '       MsgBox "Outlook não esta disponivel!"
     '       Exit Sub
     '   End If
    'End If

'feed vars col position
colLoja = 1
colCP = 2
colCidade = 3
colUF = 4
colStatusMX = 6
colSituacao = 8
colPossibilidade = 9
colResponsavel = 10
colEmailDest = 17
colDataRetorno = 15
colHoraRetorno = 16
colStatusAcao = 13

'get atual
linhaAtual = linha_Atual.linha_Atual

'Check main fields
If (Worksheets(plan).Cells(linhaAtual, colEmailDest) = "") Then
    MsgBox ("Favor preencher o endereço de e-mail na coluna" + CStr(colEmailDest))
    Exit Sub

    
Else
    loja = Worksheets(plan).Cells(linhaAtual, colLoja)
    dataretorno = Worksheets(plan).Cells(linhaAtual, colDataRetorno)
    horaretorno = Worksheets(plan).Cells(linhaAtual, colHoraRetorno)
    statusacao = Worksheets(plan).Cells(linhaAtual, colStatusAcao)
    'emaildest = Worksheets(plan).Cells(linhaAtual, colEmailDest)
    'MsgBox (linhaAtual)
End If


'Check mail address to send
If (Worksheets("config.ini").Cells(1, 2) <> "") Then
    email = Worksheets("config.ini").Cells(1, 2)
Else
    MsgBox ("Preencha o endereço de e-mail na aba CONFIG.INI")
End If

'Check Trello address to add as a member
If (Worksheets(plan).Cells(linhaAtual, colEmailDest) <> "") Then
    emaildest = Worksheets(plan).Cells(linhaAtual, colEmailDest)
Else
    MsgBox ("Preencha o endereço de responsavel do Outlook na coluna " + CStr(colEmailDest))
End If

'Feed email fields
assunto = "Agendamento lj " + loja
corpo = "loja: " + loja + nl + _
        "Data Retorno: " + CStr(dataretorno) + nl + _
        "Status Ação: " + statusacao
        


'setup app definitions
With Application
    .EnableEvents = False
    .ScreenUpdating = False
End With

'Build appointment
' adds a list of appontments to the Calendar in Outlook
'Set OutApp = CreateObject("Outlook.Application")
'Set OutMail = OutApp.CreateItem(0)




'Create Appointment

Set olAppItem = olApp.CreateItem(olAppointmentItem)
       With olAppItem
            ' set default appointment values
            '.Location = Cells(r, 3)
        .Body = corpo
        .ReminderSet = True
        .BusyStatus = olFree
        '.RequiredAttendees = emaildest
        .Recipients.Add (emaildest)
            'On Error Resume Next
        '.Start = #11/26/2018 9:00:00 AM#
        '.Start = dataretorno
        .Start = dataretorno + horaretorno
        .Duration = 90
        'myItem.End = DateValue(dataretorno) + "10:00:00 AM"
        .Subject = assunto
            '.Attachments.Add ("c:\temp\somefile.msg")
        .Location = "Telefone"
            '.Body = .Subject & ", " & Cells(r, 4).Value
        .BusyStatus = olBusy
        .Categories = "Tratativas reativação" ' add this to be able to delete the testappointments
        .Display 'show appointment
        .Save ' saves the new appointment to the default folder
        .Send
        
       End With
        
'Set olAppItem = Nothing
'Set olApp = Nothing
'MsgBox "Done !"
    
'change color of the loja (colunm 1) to mark as sent
Worksheets(plan).Cells(linhaAtual, colEmailDest).Font.Color = corAzulMarinho

'setup app definitions enable
With Application
    .EnableEvents = True
    .ScreenUpdating = True
End With

End Sub




