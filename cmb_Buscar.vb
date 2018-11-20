
Sub cmb_Buscar()

Dim ExtracaoSSOL_LINI As Integer
Dim Demandas_LINI As Integer
Dim usuario As String

Dim a As Double
Dim b As Double
Dim CriterioBusca As Double
Dim NovaLinha As Double
Dim Adicionas As Boolean
Dim ApagarPintura As Boolean

Dim planilhademandas As String
Dim planilhaextracao As String

planilhademandas = Sheets("config.ini").Range("B2")
planilhaextracao = Sheets("config.ini").Range("B7")

ExtracaoSSOL_LINI = 2
Demandas_LINI = 3

Application.ScreenUpdating = False

If CVDate(Format$(Now(), "DD/MM/YYYY")) = CVDate(Sheets("config.ini").Range("B9")) Then

    a = MsgBox("Você já fez essa busca hoje. Deseja remover o destaque em amarelo das novas SSs nessa busca?", vbCritical + vbYesNo, "Atenção")
    
        If a = 6 Then ApagarPintura = True Else ApagarPintura = False
           
Else
    Sheets("config.ini").Range("B9") = Format$(Now(), "DD/MM/YYYY")
    ApagarPintura = True
    
End If

'Busca linha NOVA disponível

a = Demandas_LINI

 Do While Not Sheets(planilhademandas).Range("A" & a) = ""
 
    Sheets(planilhademandas).Activate
    Range("A" & a).Select
        If ApagarPintura = True Then
    
            With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
  a = a + 1
 Loop
     
 NovaLinha = a
 
'Busca Novas SSs

 a = ExtracaoSSOL_LINI
 CriterioBusca = 0
 
 Do While Not Sheets(planilhaextracao).Range("A" & a) = ""
 
    CriterioBusca = Val(Sheets(planilhaextracao).Range("A" & a))
 
    For b = Demandas_LINI To NovaLinha
    
        If CriterioBusca = Sheets(planilhademandas).Range("A" & b) Then
            Adicionar = False
            Exit For
        Else
            Adicionar = True
        End If
     
    Next b
    
    If Adicionar = True Then
        Sheets(planilhademandas).Range("A" & NovaLinha) = Sheets(planilhaextracao).Range("A" & a)
        Range("A" & NovaLinha).Select
        
            Sheets(planilhademandas).Activate
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        
        NovaLinha = NovaLinha + 1
    End If
     
 a = a + 1
 Loop


' Busca Datas na ExtraçãOdoSSOL
 a = Demandas_LINI
 Do While Not Sheets(planilhademandas).Range("A" & a) = ""
 
    CriterioBusca = Val(Sheets(planilhademandas).Range("A" & a))
    
        b = ExtracaoSSOL_LINI
        Do While Not Sheets(planilhaextracao).Range("a" & b) = ""
        
            If CriterioBusca = Val(Sheets(planilhaextracao).Range("a" & b)) Then
            
                If (Sheets(planilhaextracao).Range("G" & b) <> "--------" And Sheets(planilhaextracao).Range("G" & b) <> "---------") Then
                    Sheets(planilhademandas).Range("G" & a) = Sheets(planilhaextracao).Range("G" & b)
                Else
                    Sheets(planilhademandas).Range("G" & a) = ""
                End If
                
                If (Sheets(planilhaextracao).Range("H" & b) <> "--------" And Sheets(planilhaextracao).Range("H" & b) <> "---------") Then
                    Sheets(planilhademandas).Range("I" & a) = Sheets(planilhaextracao).Range("H" & b)
                Else
                    Sheets(planilhademandas).Range("I" & a) = ""
                End If
                
                If (Sheets(planilhaextracao).Range("I" & b) <> "--------" And Sheets(planilhaextracao).Range("I" & b) <> "---------") Then
                    Sheets(planilhademandas).Range("J" & a) = Sheets(planilhaextracao).Range("I" & b)
                Else
                    Sheets(planilhademandas).Range("J" & a) = ""
                End If
                
                Sheets(planilhademandas).Range("B" & a) = Sheets(planilhaextracao).Range("C" & b)
                Sheets(planilhademandas).Range("D" & a) = Sheets(planilhaextracao).Range("K" & b)
                Sheets(planilhademandas).Range("C" & a) = Sheets(planilhaextracao).Range("D" & b)
                Sheets(planilhademandas).Range("R" & a) = Sheets(planilhaextracao).Range("B" & b)
                Sheets(planilhademandas).Range("Q" & a) = Sheets(planilhaextracao).Range("J" & b)
                
                'pega codigo do usuário da planilha de extração
                usuario = Sheets(planilhaextracao).Range("F" & b)
                'pinta cel usuario de branco
                 Sheets(planilhademandas).Range("S" & a).Select
                 Sheets(planilhademandas).Activate
                 With Selection.Interior
                                .Pattern = xlNone
                                .TintAndShade = 0
                                .PatternTintAndShade = 0
                End With
                'procura usuario de acordo com o codigo na tabela config.ini
                Select Case usuario
                    Case Sheets("config.ini").Range("A" & 12)
                        Sheets(planilhademandas).Range("S" & a) = Sheets("config.ini").Range("B" & 12)
                    Case Sheets("config.ini").Range("A" & 13)
                        Sheets(planilhademandas).Range("S" & a) = Sheets("config.ini").Range("B" & 13)
                    Case Sheets("config.ini").Range("A" & 14)
                        Sheets(planilhademandas).Range("S" & a) = Sheets("config.ini").Range("B" & 14)
                    Case Sheets("config.ini").Range("A" & 15)
                        Sheets(planilhademandas).Range("S" & a) = Sheets("config.ini").Range("B" & 15)
                    Case Sheets("config.ini").Range("A" & 16)
                        Sheets(planilhademandas).Range("S" & a) = Sheets("config.ini").Range("B" & 16)
                    Case Sheets("config.ini").Range("A" & 17)
                        Sheets(planilhademandas).Range("S" & a) = Sheets("config.ini").Range("B" & 17)
                    Case Sheets("config.ini").Range("A" & 18)
                        Sheets(planilhademandas).Range("S" & a) = Sheets("config.ini").Range("B" & 18)
                    Case Sheets("config.ini").Range("A" & 19)
                        Sheets(planilhademandas).Range("S" & a) = Sheets("config.ini").Range("B" & 19)
                    Case Sheets("config.ini").Range("A" & 20)
                        Sheets(planilhademandas).Range("S" & a) = Sheets("config.ini").Range("B" & 20)
                    Case Sheets("config.ini").Range("A" & 21)
                        Sheets(planilhademandas).Range("S" & a) = Sheets("config.ini").Range("B" & 21)
                    Case Sheets("config.ini").Range("A" & 22)
                        Sheets(planilhademandas).Range("S" & a) = Sheets("config.ini").Range("B" & 22)
                    Case Sheets("config.ini").Range("A" & 23)
                        Sheets(planilhademandas).Range("S" & a) = Sheets("config.ini").Range("B" & 23)
                    Case Sheets("config.ini").Range("A" & 24)
                        Sheets(planilhademandas).Range("S" & a) = Sheets("config.ini").Range("B" & 24)
                    'caso não tenha o codigo cadastrado na config.ini, msg para verificar usuário
                    Case Else
                        Sheets(planilhademandas).Range("S" & a) = "VERIFICAR USUARIO"
                        Range("S" & a).Select
                        Sheets(planilhademandas).Activate
                        With Selection.Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 255
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                End Select
                                
                Exit Do
             End If
             
            b = b + 1
        Loop
        
    a = a + 1
 Loop

Application.ScreenUpdating = True

End Sub