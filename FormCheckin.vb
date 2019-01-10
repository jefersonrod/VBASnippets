Sub FormCheckin()
'get info from sheet and generate card in Trello
On Error Resume Next
On Error GoTo 0

'set general vars
Dim plan As String
Dim corpo As String

Dim loja As String
Dim lojaCodigo As String
Dim line As String
Dim checkin As String
Dim checkout As String
'Técnico
Dim tecnico As String
Dim recebeuContatodaPrimesys As String
Dim recebeuContatodaPrimesysRESP As String
Dim recebeuOrientacõesSobreoManualdeMigracao As String
Dim recebeuOrientacõesSobreoManualdeMigracaoRESP As String
Dim jaRealizouMigracao As String
Dim jaRealizouMigracaoRESP As String
Dim possuiWhatsappQual As String
Dim possuiWhatsappQualRESP As String
Dim informarSobreoLinkqueEstaSendoInstalado As String
Dim informarSobreoLinkqueEstaSendoInstaladoRESP As String
Dim envioFotosRackRetaguardaBalcao As String
Dim envioFotosRackRetaguardaBalcaoRESP As String

'Responsavel
Dim resposavel As String
Dim InformarSobreAcompanhamento As String
Dim InformarSobreAcompanhamentoRESP As String
Dim temAlgumChamadoAberto As String
Dim temAlgumChamadoAbertoRESP As String
Dim estaComAlgumProblemaSistemico As String
Dim estaComAlgumProblemaSistemicoRESP As String
Dim orientarAssinarOSSomenteApos As String
Dim orientarAssinarOSSomenteAposRESP As String


Dim corAzul As String
Dim corAzulMarinho As String
Dim corAzulMarinhoCode As String

'set general vars
plan = FunctionsTimeModelX.ActualSheetName
nl = vbCrLf 'new line
line = "---------------------------------"
checkin = "*Check-In*"
checkout = "*Check-Out*"

    'Check loja field
    If (Worksheets(plan).Cells(5, 3) = "") Then
        MsgBox ("Favor preencher o numero da loja na linha 5")
        Worksheets(plan).Cells(5, 3).Select
        Exit Sub
        
    Else
        'check tecnico answer
        For i = 9 To 19 Step 2
            If (Worksheets(plan).Cells(i, 3) = "") Then
            MsgBox ("Favor preencher a linha " + CStr(i))
            Worksheets(plan).Cells(i, 3).Select
            Exit Sub
            End If
            
        Next i
        
        'check responsavel field
        For i = 23 To 29 Step 2
            If (Worksheets(plan).Cells(i, 3) = "") Then
            MsgBox ("Favor preencher a linha " + CStr(i))
            Worksheets(plan).Cells(i, 3).Select
            Exit Sub
            End If
            
        Next i
        
        loja = Worksheets(plan).Cells(4, 3)
        lojaCodigo = Worksheets(plan).Cells(5, 3)
        'Técnico
        tecnico = Worksheets(plan).Cells(7, 3)
        recebeuContatodaPrimesys = Worksheets(plan).Cells(8, 3)
        recebeuContatodaPrimesysRESP = Worksheets(plan).Cells(9, 3)
        recebeuOrientacõesSobreoManualdeMigracao = Worksheets(plan).Cells(10, 3)
        recebeuOrientacõesSobreoManualdeMigracaoRESP = Worksheets(plan).Cells(11, 3)
        jaRealizouMigracao = Worksheets(plan).Cells(12, 3)
        jaRealizouMigracaoRESP = Worksheets(plan).Cells(13, 3)
        possuiWhatsappQual = Worksheets(plan).Cells(14, 3)
        possuiWhatsappQualRESP = Worksheets(plan).Cells(15, 3)
        informarSobreoLinkqueEstaSendoInstalado = Worksheets(plan).Cells(16, 3)
        informarSobreoLinkqueEstaSendoInstaladoRESP = Worksheets(plan).Cells(17, 3)
        envioFotosRackRetaguardaBalcao = Worksheets(plan).Cells(18, 3)
        envioFotosRackRetaguardaBalcaoRESP = Worksheets(plan).Cells(19, 3)
        'Responsavel
        responsavel = Worksheets(plan).Cells(21, 3)
        InformarSobreAcompanhamento = Worksheets(plan).Cells(22, 3)
        InformarSobreAcompanhamentoRESP = Worksheets(plan).Cells(23, 3)
        temAlgumChamadoAberto = Worksheets(plan).Cells(24, 3)
        temAlgumChamadoAbertoRESP = Worksheets(plan).Cells(25, 3)
        estaComAlgumProblemaSistemico = Worksheets(plan).Cells(26, 3)
        estaComAlgumProblemaSistemicoRESP = Worksheets(plan).Cells(27, 3)
        orientarAssinarOSSomenteApos = Worksheets(plan).Cells(28, 3)
        orientarAssinarOSSomenteAposRESP = Worksheets(plan).Cells(29, 3)
        
    End If
    
    
    
    'generate content text
    
    corpo = checkin + nl + _
            loja + lojaCodigo + nl + nl + _
            tecnico + nl + nl + _
            recebeuContatodaPrimesys + nl + _
            recebeuContatodaPrimesysRESP + nl + _
            recebeuOrientacõesSobreoManualdeMigracao + nl + _
            recebeuOrientacõesSobreoManualdeMigracaoRESP + nl + _
            jaRealizouMigracao + nl + _
            jaRealizouMigracaoRESP + nl + _
            possuiWhatsappQual + nl + _
            possuiWhatsappQualRESP + nl + _
            informarSobreoLinkqueEstaSendoInstalado + nl + _
            informarSobreoLinkqueEstaSendoInstaladoRESP + nl + _
            envioFotosRackRetaguardaBalcao + nl + _
            envioFotosRackRetaguardaBalcaoRESP + nl + nl + _
            line + nl + _
            responsavel + nl + nl + _
            InformarSobreAcompanhamento + nl + _
            InformarSobreAcompanhamentoRESP + nl + _
            temAlgumChamadoAberto + nl + _
            temAlgumChamadoAbertoRESP + nl + _
            estaComAlgumProblemaSistemico + nl + _
            estaComAlgumProblemaSistemicoRESP + nl + _
            orientarAssinarOSSomenteApos + nl + _
            orientarAssinarOSSomenteAposRESP + nl
            
    
    'generate content
    CopyText corpo 'call copy to clipboard function
    
    'change color of the loja (colunm 1) to mark as sent
    'Worksheets(plan).Cells(linhaAtual, colLoja).Font.Color = corAzul

End Sub
