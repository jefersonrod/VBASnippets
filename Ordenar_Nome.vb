Sub Ordenar_Nome()
'
' Ordenar_Nome Macro
' Ordenar pela coluna de nomes
'

'
    Columns("A:A").Select
    ActiveWorkbook.Worksheets("Fila Dev").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Fila Dev").Sort.SortFields.Add Key:=Range("A1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Fila Dev").Sort
        .SetRange Range("A3:BN1000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
