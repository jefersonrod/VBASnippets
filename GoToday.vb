Sub GoToday()
    Dim hoje As Date
    Dim planilhaFilaDev As String
    Dim rng As Range
    Dim cel As Range
    
    planilhaFilaDev = Worksheets("config.ini").Cells(5, 2)
    
    hoje = Date
    
    'procura data inicial para preencher
    With Worksheets(planilhaFilaDev)
        'LastCol = Worksheets("Fila Dev").Cells(2, .Columns.count).End(xlToLeft).Column
        LastColAddress = Worksheets(planilhaFilaDev).Cells(2, .Columns.count).End(xlToLeft).Address
        'LastColValor = Worksheets(planilhaFilaDev).Cells(2, .Columns.count).End(xlToLeft).Value
        lastAddressColDate = Worksheets(planilhaFilaDev).Range(LastColAddress).Address
    End With
    
    rangeTotalCalendario = "D2:" + lastAddressColDate
            'MsgBox (rangeTotalCalendario)
            
    Set rng = Worksheets(planilhaFilaDev).Range(rangeTotalCalendario)
    For Each cel In rng.Cells
    
        With cel
            
            If hoje = .Value Then
                'MsgBox ("Found")
                Debug.Print .Address & ":" & .Value
                address_found = .Address
            End If
            
        End With

    Next cel
    
    Worksheets(planilhaFilaDev).Range(address_found).Activate

End Sub


