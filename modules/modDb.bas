Attribute VB_Name = "modDb"
Function FetchNewToken() As String
    Dim xhr As New XMLHTTP60
    Dim url As String, username As String, pwd As String, dataToPost As String
    Dim jsonResp As String, token As String
    Dim initPos As Integer, length As Integer
    
    username = cfConfiguracoes.Cells().Find(what:="Login Sísifo", lookat:=xlWhole).Offset(0, 1).Formula
    pwd = cfConfiguracoes.Cells().Find(what:="Senha Sísifo", lookat:=xlWhole).Offset(0, 1).Formula
    dataToPost = "?email=" & username & "&password=" & pwd
    url = SISIFO_URL & "login" & dataToPost
    xhr.Open "POST", url
    xhr.Send
    
    jsonResp = Replace(xhr.ResponseText, "\/", "/")
    initPos = InStr(1, jsonResp, """access_token"":""") + 16
    length = InStr(1, jsonResp, """,""token_type"":""") - initPos
    token = Mid(jsonResp, initPos, length)
    FetchNewToken = token
End Function

Function FetchEseloInfoByEspaiderJuizoSlug(slug As String) As String()
    Dim token As String, jsonResp As String, response() As String, eseloEntries() As String
    Dim initPos As Integer, length As Integer
    Dim EseloInfo() As Variant
    Dim retried As Boolean, invalidToken As Boolean
    
    Do
        token = GetStoredApiTokenEnsuringItExists
        jsonResp = RequestEseloInfoByEspaiderJuizoSlug(slug, token)
        If InStr(1, LCase(jsonResp), "token expirado") > 0 Then
            If invalidToken = True Then retried = True
            invalidToken = True
            SetStoredApiToken ""
        Else
            invalidToken = False
        End If
    Loop Until invalidToken = False Or (retried = True And invalidToken = True)
    
    eseloEntries = FormatEseloJuizosResponseAsArray(jsonResp)
    
    If eseloEntries(0) = "" Then
        ReDim response(0, 0)
        response(0, 0) = ""
    Else
        response = ParseEseloEntries(eseloEntries)
    End If
    
    FetchEseloInfoByEspaiderJuizoSlug = response
End Function

Function RequestEseloInfoByEspaiderJuizoSlug(slug As String, token As String) As String
    Dim xhr As New XMLHTTP60
    Dim url As String, response As String
    
    url = SISIFO_URL & "info-eselo-via-espaider/" & slug
    xhr.Open "GET", url
    xhr.SetRequestHeader "Authorization", "Bearer " & token
    xhr.Send
    response = Replace(xhr.ResponseText, "\/", "/")
    RequestEseloInfoByEspaiderJuizoSlug = response
End Function

Function FormatEseloJuizosResponseAsArray(json As String) As String()
    Dim regDrv As New RegExp
    Dim matches As MatchCollection
    Dim i As Integer
    Dim response() As String
    
    With regDrv
        .Global = True
        .pattern = "({""eseloJuizo"":){1}.*?}{1}"
        Set matches = .Execute(json)
    End With
    
    If matches.Count = 0 Then
        ReDim response(0)
        response(0) = ""
    Else
        ReDim response(matches.Count - 1)
        For i = 0 To matches.Count - 1 Step 1
            response(i) = matches(i)
        Next i
    End If
    
    FormatEseloJuizosResponseAsArray = response
End Function

Function ParseEseloEntries(eseloEntries() As String) As String()
    Dim regDrv As New RegExp
    Dim matches As MatchCollection
    Dim juizoPattern As String, comarcaPattern As String, juizo As String, comarca As String
    Dim i As Integer
    Dim response() As String
    
    juizoPattern = "(""eseloJuizo"":"").*?(""){1}"
    comarcaPattern = "(""eseloComarca"":"").*?(""){1}"
    
    regDrv.Global = True
    ReDim response(UBound(eseloEntries), 1)
    For i = 0 To UBound(eseloEntries) Step 1
        regDrv.pattern = juizoPattern
        Set matches = regDrv.Execute(eseloEntries(i))
        juizo = Replace(matches(0), """eseloJuizo"":""", "")
        juizo = Left(juizo, Len(juizo) - 1)
        
        regDrv.pattern = comarcaPattern
        Set matches = regDrv.Execute(eseloEntries(i))
        comarca = Replace(matches(0), """eseloComarca"":""", "")
        comarca = Left(comarca, Len(comarca) - 1)
        
        response(i, 0) = juizo
        response(i, 1) = comarca
    Next i
    
    ParseEseloEntries = response
End Function

Function FetchSapConfigs() As SapInfo
    Dim token As String, jsonResp As String
    Dim response As SapInfo
    Dim retried As Boolean, invalidToken As Boolean
    
    Do
        token = GetStoredApiTokenEnsuringItExists
        jsonResp = RequestSapConfigs(token)
        If InStr(1, LCase(jsonResp), "token expirado") > 0 Then
            If invalidToken = True Then retried = True
            invalidToken = True
            SetStoredApiToken ""
        Else
            invalidToken = False
        End If
    Loop Until invalidToken = False Or (retried = True And invalidToken = True)
    
    response = ParseSapConfigsJson(jsonResp)
    FetchSapConfigs = response
End Function

Function RequestSapConfigs(token As String) As String
    Dim xhr As New XMLHTTP60
    Dim url As String
    
    url = SISIFO_URL & "custas-configs/"
    xhr.Open "GET", url
    xhr.SetRequestHeader "Authorization", "Bearer " & token
    xhr.Send
    
    RequestSapConfigs = Replace(xhr.ResponseText, "\/", "/")
End Function

Function ParseSapConfigsJson(json As String) As SapInfo
    Dim response As SapInfo
    Dim i As Integer
    
    With response
        .TipoDocumento = GetOneConfigFromJson("Tipo de Documento", json)
        .ReferenciaCabecalho = GetOneConfigFromJson("Referência Cabeçalho", json)
        .TextoCabecalho = GetOneConfigFromJson("Texto Cabeçalho", json)
        .NumContaCliente = GetOneConfigFromJson("Nº Conta Cliente", json)
        .Nome = GetOneConfigFromJson("Nome", json)
        .Cnpj = GetOneConfigFromJson("CNPJ", json)
        .Cpf = GetOneConfigFromJson("CPF", json)
        .Rua = GetOneConfigFromJson("Rua", json)
        .Local = GetOneConfigFromJson("Local", json)
        .Cep = GetOneConfigFromJson("CEP", json)
        .CondicoesPagamento = GetOneConfigFromJson("Condições de Pagamento", json)
        .FormaPagamento = GetOneConfigFromJson("Forma de Pagamento", json)
        .AtribuicaoFornecedor = GetOneConfigFromJson("Atribuição Fornecedor", json)
        .BancoEmpresa = GetOneConfigFromJson("Banco Empresa", json)
        .ChaveBreveConta = GetOneConfigFromJson("Chave Breve da conta", json)
        .ContaRazao = GetOneConfigFromJson("Conta Razao", json)
        .AtribuicaoDespesa = GetOneConfigFromJson("Atribuição Despesa", json)
        .CentroCusto = GetOneConfigFromJson("Centro de Custo", json)
    End With
    
    ParseSapConfigsJson = response
End Function

Function GetOneConfigFromJson(configName As String, json As String) As String
    Dim regDrv As New RegExp
    Dim matches As MatchCollection
    Dim pattern As String, response As String
    
    pattern = "(""nome"":""" & configName & """,""valor"":""){1}.*?("",)"
    regDrv.Global = True
    regDrv.pattern = pattern
    Set matches = regDrv.Execute(json)
    response = Replace(matches(0), """nome"":""" & configName & """,""valor"":""", "")
    response = Left(response, Len(response) - 2)
    GetOneConfigFromJson = response
End Function

Function PostDajeToDb(info As DajeInfo) As String
    Dim token As String, jsonResp As String
    Dim response As String
    Dim retried As Boolean, invalidToken As Boolean
    
    Do
        token = GetStoredApiTokenEnsuringItExists
        jsonResp = MakeRequestPostDaje(info, token)
        If InStr(1, LCase(jsonResp), "token expirado") > 0 Then
            If invalidToken = True Then retried = True
            invalidToken = True
            SetStoredApiToken ""
        Else
            invalidToken = False
        End If
    Loop Until invalidToken = False Or (retried = True And invalidToken = True)
    
    If InStr(1, LCase(jsonResp), "criado com sucesso") > 0 Then jsonResp = "sucesso"
    PostDajeToDb = jsonResp
End Function

Function MakeRequestPostDaje(info As DajeInfo, token As String) As String
    Dim xhr As New XMLHTTP60
    Dim url As String, dataToPost As String
    Dim emissionDate As String, dueDate As String, value As String
    
    emissionDate = Format(info.emissionDate, "YYYY-MM-DD")
    dueDate = Format(info.dueDate, "YYYY-MM-DD")
    value = Replace(info.valor, ".", "")
    value = Replace(value, ",", ".")
    
    dataToPost = "numero=" & info.dajeNumber _
                & "&processo=" & info.processoNumber _
                & "&parte_adversa=" & info.Adverso _
                & "&valor=" & value _
                & "&emissao=" & emissionDate _
                & "&vencimento=" & dueDate _
                & "&tipo=" & info.actType _
                & "&qtd_atos=" & info.actsQuantity _
                & "&eventos_atos=" & info.actEventId _
                & "&gerencia=" & info.gerenciaEmbasa _
                & "&codigo_barras=" & info.barCode
                
    url = SISIFO_URL & "dajes"
    xhr.Open "POST", url
    xhr.SetRequestHeader "Authorization", "Bearer " & token
    xhr.SetRequestHeader "Content-type", "application/x-www-form-urlencoded"
    xhr.Send dataToPost
    
    MakeRequestPostDaje = Replace(xhr.ResponseText, "\/", "/")
End Function
