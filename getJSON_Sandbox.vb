Sub getJSON()

    Dim sJSONString As String
    Dim vJSON
    Dim sState As String
    Dim aData()
    Dim aHeader()
    Dim startLineSheet As Integer
    Dim startColSheet As Integer
    Dim i As Integer
    Dim Plan As String
    
    Plan = "JSONTest"
    startLineSheet = 6
    
    ' start of loop to get API rest
    For j = 0 To 13
     ' Retrieve data
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", "http://hchecklist.redeboticario.vsat/franquia/demandaJSON", False
        .Send
        sJSONString = .ResponseText
    End With
    
    ' Parse JSON response
    JSON.Parse sJSONString, vJSON, sState
    
    'set positions
    startColSheet = 1
    
    Set vJSON = vJSON(CStr(j))
    
    'debug print
    
    'Debug.Print vJSON(CStr(i))
    'Debug.Print "startLineSheet: " & startLineSheet
    'Debug.Print "Id: " & vJSON("Id")
    'Debug.Print "CodigoFranquia: " & vJSON("CodigoFranquia")
    
    ' Access to each item in dictionary
    
    
    Worksheets(Plan).Cells(startLineSheet, startColSheet) = vJSON("Id")
    startColSheet = startColSheet + 1
    Worksheets(Plan).Cells(startLineSheet, startColSheet) = vJSON("CodigoFranquia")
    startColSheet = startColSheet + 1
    Worksheets(Plan).Cells(startLineSheet, startColSheet) = vJSON("Projeto")
    startColSheet = startColSheet + 1
    
    Call CheckJsonDate(vJSON("Data"), startLineSheet, startColSheet)
    
    startColSheet = startColSheet + 1
    Worksheets(Plan).Cells(startLineSheet, startColSheet) = vJSON("Contato")
    startColSheet = startColSheet + 1
    Worksheets(Plan).Cells(startLineSheet, startColSheet) = vJSON("Telefone")
    startColSheet = startColSheet + 1
    Worksheets(Plan).Cells(startLineSheet, startColSheet) = vJSON("Email")
    startColSheet = startColSheet + 1
    Call CheckClube(vJSON("Clube"), startLineSheet, startColSheet)
    'Worksheets(Plan).Cells(startLineSheet, startColSheet) = vJSON("Clube")
    startColSheet = startColSheet + 1
    Worksheets(Plan).Cells(startLineSheet, startColSheet) = vJSON("Mobshop")
    startColSheet = startColSheet + 1
    Worksheets(Plan).Cells(startLineSheet, startColSheet) = vJSON("Switch")
    startColSheet = startColSheet + 1
        
    Call CheckJsonDate(vJSON("DataSolicitacaoViabilidade"), startLineSheet, startColSheet)
    
    startColSheet = startColSheet + 1
            
    Call CheckJsonDate(vJSON("DataRetornoViabilidade"), startLineSheet, startColSheet)
    
    startColSheet = startColSheet + 1
    
    Worksheets(Plan).Cells(startLineSheet, startColSheet) = vJSON("Link")
    
    startColSheet = startColSheet + 1
    
    Call CheckJsonDate(vJSON("DataContracaoLink"), startLineSheet, startColSheet)
    
    startColSheet = startColSheet + 1
        
    Call CheckJsonDate(vJSON("DataInstalacaoLink"), startLineSheet, startColSheet)
    
    startColSheet = startColSheet + 1
    
    Call CheckJsonDate(vJSON("DataConclusao"), startLineSheet, startColSheet)
    
    startColSheet = startColSheet + 1
        
    Worksheets(Plan).Cells(startLineSheet, startColSheet) = vJSON("StatusDemanda")
    
    startColSheet = startColSheet + 1
    Worksheets(Plan).Cells(startLineSheet, startColSheet) = vJSON("Usuario")
    startColSheet = startColSheet + 1
    Worksheets(Plan).Cells(startLineSheet, startColSheet) = vJSON("Obs")
    startColSheet = startColSheet + 1
    
    Call CheckJsonDate(vJSON("DataDevolucaoKitVSAT"), startLineSheet, startColSheet)
    
    startColSheet = startColSheet + 1
    
    Call CheckJsonDate(vJSON("DataDevolucaoAntena"), startLineSheet, startColSheet)
    
    startColSheet = startColSheet + 1
    
    Call CheckJsonDate(vJSON("DataInauguracaoLoja"), startLineSheet, startColSheet)
    
    startColSheet = startColSheet + 1
    
    startLineSheet = startLineSheet + 1
    
    'Debug.Print "=========================="
    Next j

    
End Sub

