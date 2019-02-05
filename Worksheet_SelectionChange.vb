Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)
'set color of header and filters
Dim planCfg As String
Dim lineAnalist As Integer
Dim corAreaBts As String
Dim corAreaFiltro As String
Dim corFonteFiltro As String

planCfg = "config.ini"
lineAnalist = FunctionsTimeModelX.CheckAnalystLine
corAreaBts = Worksheets(planCfg).Cells(lineAnalist, 8)
corAreaFiltro = Worksheets(planCfg).Cells(lineAnalist, 9)
corFonteFiltro = Worksheets(planCfg).Cells(lineAnalist, 10)

Range("A1:R1").Interior.Color = FunctionsTimeModelX.convertRGB(corAreaBts)
Range("A2:R2").Interior.Color = FunctionsTimeModelX.convertRGB(corAreaFiltro)
Range("A2:R2").Font.Color = FunctionsTimeModelX.convertRGB(corFonteFiltro)
'Update 20140904 by Jef *change to highlight only selected line
    Static xRow
    Static xColumn
            
        If xRow <> "" Then
                       
            With Rows(xRow).Interior
            .ColorIndex = xlNone
            End With
        End If
        
    pRow = Selection.Row
    
    xRow = pRow
               
            With Rows(pRow).Interior
            .ColorIndex = 15
            .Pattern = xlSolid
            End With
End Sub

