Sub CreateAppointmentOutlookAtendimentosTimeModeloX()

Dim olApp As Outlook.Application
Dim olAppItem As Outlook.AppointmentItem

On Error Resume Next
On Error GoTo 0

Dim plan As String
Dim email As String
Dim assunto As String
Dim corpo As String
Dim linhaAtual As Integer
Dim colLoja As Integer
Dim colTecnico As Integer
Dim colResponsavel As Integer
Dim colEmailDest As Integer
Dim colProblema As Integer
Dim colSolucao As Integer
Dim colDataRetorno As Integer
Dim colHoraRetorno As Integer

Dim loja As String
Dim tecnico As String
Dim responsavel As String
Dim problema As String
Dim solucao As String
Dim dataretorno As Date
Dim horaretorno As Date
Dim linhaConfig As Integer
Dim corAzulMarinho As String



nl = vbCrLf 'new line
br = "<br>" 'new line HTML
corAzulMarinho = RGB(68, 114, 196)
plan = FunctionsTimeModelX.ActualSheetName
linhaConfig = CheckAnalystLine

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
colTecnico = 4
colResponsavel = 6
colProblema = 8
colSolucao = 9
colDataRetorno = 17
colHoraRetorno = 18
colEmailDest = 5

'get atual
linhaAtual = linha_Atual.linha_Atual

'Check main fields
If (Worksheets("config.ini").Cells(linhaConfig, colEmailDest) = "") Then
    MsgBox ("Preencha o endereço de e-mail na aba CONFIG.INI, na coluna " + CStr(colEmailDest))
    Exit Sub
Else
    email = Worksheets("config.ini").Cells(linhaConfig, colEmailDest)
    loja = Worksheets(plan).Cells(linhaAtual, colLoja)
    dataretorno = Worksheets(plan).Cells(linhaAtual, colDataRetorno)
    horaretorno = Worksheets(plan).Cells(linhaAtual, colHoraRetorno)
    tecnico = Worksheets(plan).Cells(linhaAtual, colTecnico)
    responsavel = Worksheets(plan).Cells(linhaAtual, colResponsavel)
    problema = Worksheets(plan).Cells(linhaAtual, colProblema)
    solucao = Worksheets(plan).Cells(linhaAtual, colSolucao)
    
    If CStr(dataretorno) = "00:00:00" Then MsgBox ("Data de agendamento não infomado, verifique"): Exit Sub
    If CStr(horaretorno) = "00:00:00" Then MsgBox ("Hora do agendamento não infomado, verifique"): Exit Sub
    
End If

'Feed email fields
assunto = loja
corpo = "loja: " + loja + nl + _
        "Data Retorno: " + CStr(dataretorno) + nl + _
        "Hora Retorno: " + CStr(horaretorno) + nl + _
        "Técnico: " + tecnico + nl + _
        "Responsavel: " + responsavel + nl + _
        "Problema: " + problema + nl + _
        "Solução: " + solucao + nl + nl



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
        .Recipients.Add (email)
            'On Error Resume Next
        '.Start = #11/26/2018 9:00:00 AM#
        .Start = dataretorno + horaretorno '"09:00:00 AM"
        .Duration = 30
        'myItem.End = DateValue(dataretorno) + "10:00:00 AM"
        .Subject = assunto
            '.Attachments.Add ("c:\temp\somefile.msg")
        .Location = "Telefone"
            '.Body = .Subject & ", " & Cells(r, 4).Value
        .BusyStatus = olBusy
        .Display 'show appointment
        .Save ' saves the new appointment to the default folder
        .Send
        
       End With
        
'Set olAppItem = Nothing
'Set olApp = Nothing
'MsgBox "Done !"
    
'change color of the loja (colunm 1) to mark as sent
'Worksheets(plan).Cells(linhaAtual, colEmailDest).Font.Color = corAzulMarinho

'setup app definitions enable
With Application
    .EnableEvents = True
    .ScreenUpdating = True
End With

End Sub
