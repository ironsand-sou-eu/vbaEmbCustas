Attribute VB_Name = "modEselo"
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const contribuinte As String = "Empresa Baiana de Águas e Saneamento", _
    endereco As String = "4 avenida, 420, CAB", _
    cidade As String = "Salvador", _
    cnpjEmbasa As String = "13.504.675/0001-10"
    
Type EseloInfo
    juizo As String
    comarca As String
End Type

Type DajeGenerationInfo
    numProcesso As String
    Adverso As String
    juizosComarcas() As EseloInfo
    juizosComarcasOriginal() As EseloInfo
    diretorio As String
    atribuicaoDaje As String
    tipoAto As String
    valorAto As String
    eventosAtos As String
    atosQt As Integer
End Type

Type DajeInfo
    dajeNumber As String
    processoNumber As String
    Adverso As String
    valor As Currency
    emissionDate As Date
    dueDate As Date
    barCode As String
    actType As String
    actsQuantity As Integer
    actEventId As String
    gerenciaEmbasa As String
End Type

Type TaxableActsInfo
    celulaProcesso As Excel.Range
    comEletronicas As Integer
    comPostais As Integer
    comMandados As Integer
    litisconsortes As Integer
    eventosDigitalizacoes As String
    eventosCalculos As String
    eventosPenhoras As String
    eventosPrecatorias As String
    eventosConfComp As String
    possuiConfComp As Boolean
    faseExecucao As Boolean
End Type

Type SapInfo
    TipoDocumento As String
    ReferenciaCabecalho As String
    TextoCabecalho As String
    NumContaCliente As String
    Nome As String
    Cnpj As String
    Cpf As String
    Rua As String
    Local As String
    Cep As String
    CondicoesPagamento As String
    FormaPagamento As String
    AtribuicaoFornecedor As String
    BancoEmpresa As String
    ChaveBreveConta As String
    ContaRazao As String
    AtribuicaoDespesa As Integer
    CentroCusto As String
End Type
 
Sub GerarUmDAJE(info As DajeGenerationInfo)
    Dim contentDiff As String, dajesFolderInitialContent As String, dajesFolderFinalContent As String
    Dim fileNameWithPath As String
    Dim juizoComarca() As EseloInfo
    Dim valorCodBarras As Currency
    Dim dspDAJE As New Despesa
    Dim by As New Selenium.by
    Dim infoSap As SapInfo
    
    info.juizosComarcas = GetJuizoComarcaAccordingToTipoAto(info.tipoAto, info.juizosComarcasOriginal)
    dajesFolderInitialContent = GetFilenamesInFolder(info.diretorio)
    
    GenerateAndDownloadDaje info, dajesFolderInitialContent
    With dspDAJE
        .NumeroProcesso = info.numProcesso
        .comarca = info.juizosComarcas(0).comarca
        .Vencimento = Date + 5
        .TipoDespesa = GetTipoDespesaAccordingToTipoAto(info.tipoAto)
        .Adverso = info.Adverso
        .Emissao = Date
        .QtdAtos = info.atosQt
        .eventosAtos = info.eventosAtos
    End With
    dajesFolderFinalContent = GetFolderContentAfterWaitingChange(info.diretorio, dajesFolderInitialContent)
    contentDiff = GetInitialToFinalContentDiff(dajesFolderInitialContent, dajesFolderFinalContent)
    fileNameWithPath = GetFileNameWithPath(info.diretorio, contentDiff)
    dspDAJE.CodigoDAJE = GetDajeNumber(contentDiff)
    dspDAJE.CodigoBarras = PegarCodBarrasPdfDaje(fileNameWithPath, dspDAJE.CodigoDAJE)
    valorCodBarras = PegarValorCodBarras(dspDAJE.CodigoBarras)
    dspDAJE.ValorDAJE = valorCodBarras
    
    infoSap = FetchSapConfigs
    dspDAJE.ExportaLinhasDespesasEspaider
End Sub

Sub GenerateAndDownloadDaje(info As DajeGenerationInfo, initialFolderContent As String)
    Dim chrome As New Selenium.ChromeDriver
    Dim dummyOutput As String
    
    SetChromeDownloadFolder chrome, info.diretorio
    chrome.Get "http://eselo.tjba.jus.br/"
    CloseHelpPopup chrome
    FillSelectField chrome, "atribuicoes", info.atribuicaoDaje
    FillSelectField chrome, "tiposatos", info.tipoAto
    CloseHelpPopup chrome
    If info.valorAto <> "" Then
        info.valorAto = Replace(info.valorAto, ".", "")
        FillValorAtoField chrome, info.valorAto
        ClickButton chrome, "botaovalordeclarado"
    End If
    CloseHelpPopup chrome
    FillTextField chrome, IIf(info.atribuicaoDaje = "RECURSOS", "maskProcesso", "numProcesso"), info.numProcesso
    ClickButton chrome, IIf(info.atribuicaoDaje = "RECURSOS", "buttonProcesso", "buttonNumProcesso")
    
    If ErrorMsgBeforeComarcaTextbox(chrome) Then
        MsgBox "Mestre, o e-Selo está afirmando que o processo informado não existe. O DAJE atual não será " & _
        "gerado. Recomendamos verificar se foi gerado algum DAJE para este processo.", vbCritical + vbOKOnly, _
        "Sísifo - Erro"
        Exit Sub
    End If
    FillSelectField chrome, "comarcas", info.juizosComarcas(0).comarca
    FillCartorioField chrome, "cartorios", info.juizosComarcas
    CloseHelpPopup chrome
    If info.atosQt <> 0 Then FillTextField chrome, "sonumeros", info.atosQt
    FillTextField chrome, "contribuinte", contribuinte
    FillTextField chrome, "endereco_completo", endereco
    FillTextField chrome, "cidade", cidade
    ClickRadioButton chrome, "tipo_doc:1"
    FillTextField chrome, "maskcpf", cnpjEmbasa
    FillTextField chrome, "complemento", info.Adverso
    ClickButton chrome, "commandButtonEmitirDaj"
    ClickDownloadImage chrome
    dummyOutput = GetFolderContentAfterWaitingChange(info.diretorio, initialFolderContent)
    ClickEmitirNovoDajeButton chrome
End Sub

Function GetJuizoComarcaAccordingToTipoAto(tipoAto As String, info() As EseloInfo) As EseloInfo()
    Dim response() As EseloInfo
    
    Select Case tipoAto
    Case "XXVII - RECURSOS (EXCLUÍDAS DESPESAS COM PORTE E REMESSA E/OU RETORNO, QUANDO CABÍVEIS) C) RECURSO INOMINADO (JUIZADOS ESPECIAIS)"
        ReDim response(0)
        response(0).juizo = "TURMA RECURSAL - SALVADOR"
        response(0).comarca = "SALVADOR"
    Case "XXVII - RECURSOS (EXCLUÍDAS DESPESAS COM PORTE E REMESSA E/OU RETORNO, QUANDO CABÍVEIS) A) APELAÇÃO, RECURSO ADESIVO"
        ReDim response(0)
        response(0).juizo = "DIRETORIA DE DISTRIBUIÇÃO DO 2º GRAU - SALVADOR"
        response(0).comarca = "SALVADOR"
    Case Else
        response = info
    End Select
    
    GetJuizoComarcaAccordingToTipoAto = response
End Function

Function GetTipoDespesaAccordingToTipoAto(tipoAto As String) As String
    Dim response As String
    
    Select Case tipoAto
    Case "I - DAS CAUSAS EM GERAL"
        response = "Valor da causa"
    Case "XXVII - RECURSOS (EXCLUÍDAS DESPESAS COM PORTE E REMESSA E/OU RETORNO, QUANDO CABÍVEIS) C) RECURSO INOMINADO (JUIZADOS ESPECIAIS)"
        response = "Recurso Inominado"
    Case "XXVII - RECURSOS (EXCLUÍDAS DESPESAS COM PORTE E REMESSA E/OU RETORNO, QUANDO CABÍVEIS) A) APELAÇÃO, RECURSO ADESIVO"
        response = "Apelação"
    Case "XXVI - ENVIO ELETRÔNICO DE CITAÇÕES, INTIMAÇÕES, OFÍCIOS E NOTIFICAÇÕES."
        response = "Comunicações eletrônicas"
    Case "III - TARIFA DE POSTAGEM - CITAÇÃO OU INTIMAÇÃO VIA POSTAL"
        response = "Comunicações postais"
    Case "XXVIII - CITAÇÃO, INTIMAÇÃO, NOTIFICAÇÃO E ENTREGA DE OFÍCIO"
        response = "Comunicações por mandado"
    Case "VII - LITISCONSÓRCIO ATIVO OU PASSIVO, POR PARTE EXCEDENTE"
        response = "Litisconsórcios"
    Case "XXI - DIGITALIZAÇÃO DE DOCUMENTO REALIZADA NO ÂMBITO DESTE PODER JUDICIÁRIO, POR DOCUMENTO (DENTRE ELES, A DIGITALIZAÇÃO DE PETIÇÃO, INLCUINDO-SE OS DOCUMENTOS ANEXADOS A ESTA, ENDEREÇADA A PROCESSO ELETRÔNICO POR MEIO FÍSICO, I.E., PAPEL)"
        response = "Digitalizações"
    Case "XVIII - AVALIAÇÕES E CÁLCULOS JUDICIAIS, POR MANDADO"
        response = "Cálculos"
    Case "XIX - REQUISIÇÃO DE INFORMAÇÕES POR MEIO ELETRÔNICO (BACENJUD, RENAJUD, INFOJUD, SERASAJUD E ASSEMELHADOS), POR CADA CONSULTA"
        response = "Penhoras"
    Case "IV - EXCEÇÃO DE IMPEDIMENTO E SUSPEIÇÃO DOS JUÍZES, CONFLITO DE COMPETÊNCIA OU DE JURISDIÇÃO SUSCITADOS PELA PARTE - DESAFORAMENTO."
        response = "Conflito de competência, suspeição ou impedimento"
    Case "XV - DEMAIS PROCESSOS OU PROCEDIMENTOS SEM VALOR DECLARADO, INCLUSIVE INCIDENTAIS, E DE IMPUGNAÇÕES EM GERAL"
        response = "Sentença de Embargos à execução"
    Case "VI - CARTA PRECATORIA, DE ORDEM E ROGATORIA, INCLUIDO PORTE DE RETORNO."
        response = "Cartas precatórias"
    Case "XVI - DESARQUIVAMENTO DE PROCESSOS, INCLUSIVE ELETRÔNICOS, POR PROCESSO"
        response = "Desarquivamento"
    End Select
    
    GetTipoDespesaAccordingToTipoAto = response
End Function

Function GetFilenamesInFolder(folder As String) As String
    Dim i As String, response As String
    
    i = Dir(folder)
    response = i
    Do While i <> ""
        i = Dir
        response = response & "," & i
    Loop
    
    GetFilenamesInFolder = response
End Function

Function GetFolderContentAfterWaitingChange(folder As String, initialFolderContent As String) As String
    Dim finalFolderContent As String
    
    Do
        finalFolderContent = GetFilenamesInFolder(folder)
    Loop While initialFolderContent = finalFolderContent Or TempDownloadFileExists(finalFolderContent)
    
    GetFolderContentAfterWaitingChange = finalFolderContent
End Function

Function TempDownloadFileExists(folderContent As String) As Boolean
    TempDownloadFileExists = InStr(1, folderContent, ".tmp") > 0 Or InStr(1, folderContent, "download") > 0
End Function

Sub SetChromeDownloadFolder(chrome As ChromeDriver, folder As String)
    chrome.SetPreference "download.default_directory", folder
    chrome.SetPreference "download.prompt_for_download", False
    chrome.SetPreference "plugins.always_open_pdf_externally", True
End Sub

Sub CloseHelpPopup(chrome As ChromeDriver)
    Dim helpDiv As Selenium.WebElement
    Dim divStyle As String
    
    On Error Resume Next
    Do
        Set helpDiv = chrome.FindElementById("ajuda")
        divStyle = helpDiv.Attribute("style")
        If InStr(1, divStyle, "display: none;") = 0 Then
            helpDiv.ExecuteScript "document.getElementById('ajuda').setAttribute('style', 'display: none;')"
        End If
    Loop Until InStr(1, divStyle, "display: none;") > 0
    On Error GoTo 0
End Sub

Sub FillSelectField(chrome As ChromeDriver, fieldId As String, valueToSelect As String)
    Dim by As New Selenium.by
    
    On Error Resume Next
    Do
        If chrome.IsElementPresent(by.ID(fieldId), 10000) Then
            chrome.FindElementById(fieldId).AsSelect.SelectByValue valueToSelect
        End If
    Loop Until chrome.FindElementById(fieldId).AsSelect.SelectedOption.value = valueToSelect
    On Error GoTo 0
End Sub

Sub FillTextField(chrome As ChromeDriver, fieldId As String, ByVal valueToFill As String)
    Dim textbox As Selenium.WebElement
    Dim by As New Selenium.by
    
    On Error Resume Next
    Do
        If chrome.IsElementPresent(by.ID(fieldId), 10000) Then
            Set textbox = chrome.FindElementById(fieldId)
            textbox.Clear
            textbox.Click
            textbox.SendKeys valueToFill
        End If
    Loop Until chrome.ExecuteScript("return document.getElementById('" & fieldId & "').value;") = valueToFill
    On Error GoTo 0
End Sub

Sub FillValorAtoField(chrome As ChromeDriver, valorAto As String)
    Dim valorTextbox As Selenium.WebElement
    Dim by As New Selenium.by
    
    On Error Resume Next
    Do
        If chrome.IsElementPresent(by.Class("real"), 10000) Then
            Set valorTextbox = chrome.FindElementByClass("real")
            valorTextbox.Clear
            valorTextbox.Click
            valorTextbox.SendKeys valorAto
        End If
    Loop Until GetEseloValorAtoValue(chrome) = valorAto
    On Error GoTo 0
End Sub

Sub FillCartorioField(chrome As ChromeDriver, fieldId As String, juizos() As EseloInfo)
    Dim by As Selenium.by
    Dim i As Integer
    
    On Error Resume Next
    Do
        Sleep 300
        For i = 0 To UBound(juizos) Step 1
            If chrome.IsElementPresent(by.ID(fieldId), 10000) Then
                chrome.FindElementById(fieldId).AsSelect.SelectByValue juizos(i).juizo
                Exit For
            End If
        Next i
    Loop Until chrome.FindElementById(fieldId).AsSelect.SelectedOption.value = juizos(i).juizo
    On Error GoTo 0
End Sub

Sub ClickButton(chrome As ChromeDriver, buttonId As String)
    Dim by As New Selenium.by
    
    On Error Resume Next
    If chrome.IsElementPresent(by.ID(buttonId), 10000) Then chrome.FindElementById(buttonId).Click
    On Error GoTo 0
End Sub

Sub ClickRadioButton(chrome As ChromeDriver, fieldIdToSelect As String)
    Dim radio As Selenium.WebElement
    Dim by As New Selenium.by
    
    On Error Resume Next
    Do
        If chrome.IsElementPresent(by.ID(fieldIdToSelect), 10000) Then
            Set radio = chrome.FindElementById(fieldIdToSelect)
            radio.Click
            chrome.ExecuteScript ("document.getElementById(fieldIdToSelect).setAttribute('checked', 'checked')")
        End If
    Loop Until chrome.FindElementById(fieldIdToSelect).IsSelected
    On Error GoTo 0
End Sub

Function GetEseloValorAtoValue(chrome As ChromeDriver) As String
    Dim valor As String
    valor = chrome.ExecuteScript("return document.getElementsByClassName('real')[0].value;")
    valor = Replace(valor, "R$", "")
    valor = Replace(valor, ".", "")
    valor = Trim(valor)
    GetEseloValorAtoValue = valor
End Function

Function ErrorMsgBeforeComarcaTextbox(chrome As ChromeDriver) As Boolean
    Dim container As Selenium.WebElement, comarcaFieldrow As Selenium.WebElement
    Dim comarcaDivLoaded As Boolean, errorMsgLoaded As Boolean
    
    On Error Resume Next
    Do
        Sleep 300
        Set container = chrome.FindElementByClass("passengerContainer")
        If container.FindElementByXPath("./*").tagName = "ul" Then errorMsgLoaded = True
        If comarcaFieldrow = container.FindElementById("detalhe_daje").FindElementsByClass("fieldRow")(2).text = "" Then comarcaDivLoaded = True
    Loop Until comarcaDivLoaded = True Or errorMsgLoaded = True
    On Error GoTo 0
    
    ErrorMsgBeforeComarcaTextbox = errorMsgLoaded
End Function

Function GetDajeValue(chrome As ChromeDriver) As Currency
    Dim response As Currency
    Dim by As Selenium.by
    
    Do
        response = Trim(Replace(chrome.FindElementById("totalizado_final").text, "R$", ""))
    Loop Until response = Trim(Replace(chrome.FindElementById("totalizado_final").text, "R$", ""))
    GetDajeValue = response
End Function

Sub ClickDownloadImage(chrome As ChromeDriver)
    Dim by As New Selenium.by
    
    On Error Resume Next
    If chrome.IsElementPresent(by.ID("img_print"), 10000) Then
        chrome.FindElementById("img_print").FindElementsByTag("a")(1).FindElementsByTag("img")(1).Click
    End If
    On Error GoTo 0
End Sub

Sub ClickEmitirNovoDajeButton(chrome As ChromeDriver)
    On Error Resume Next
    chrome.FindElementByClass("bookingButtons").FindElementByTag("input").Click
    On Error GoTo 0
End Sub

Function GetInitialToFinalContentDiff(initialFolderContent As String, finalFolderContent As String) As String
    Dim initialContentArray() As String, response As String
    Dim i As Integer
    
    initialContentArray = Split(initialFolderContent, ",")
    response = finalFolderContent
    If UBound(initialContentArray) <> -1 Then
        For i = 0 To UBound(initialContentArray) Step 1
            response = Replace(response, initialContentArray(i), "")
        Next i
    End If
    response = Replace(response, ",", "")
    GetInitialToFinalContentDiff = response
End Function

Function GetDajeNumber(fileName As String)
    Dim response As String
    
    response = Replace(fileName, "daje_", "")
    response = Replace(response, ".pdf", "")
    GetDajeNumber = response
End Function

Function GetFileNameWithPath(folder As String, fileName As String) As String
    GetFileNameWithPath = folder & fileName
End Function

Function BuscaJuizoEselo(juizoEspaider As String) As EseloInfo()
    Dim slug As String, apiResponse() As String
    
    slug = GetSlug(juizoEspaider)
    apiResponse = FetchEseloInfoByEspaiderJuizoSlug(slug)
    BuscaJuizoEselo = FormatEseloInfoFromResponseArray(apiResponse)
End Function

Function FormatEseloInfoFromResponseArray(respArray() As String) As EseloInfo()
    Dim i As Integer
    Dim response() As EseloInfo
    
    If respArray(0, 0) = "" Then
        ReDim response(0)
        response(0).comarca = ""
        response(0).juizo = ""
    Else
        ReDim response(UBound(respArray))
        For i = 0 To UBound(respArray) Step 1
            response(i).juizo = respArray(i, 0)
            response(i).comarca = respArray(i, 1)
        Next i
    End If
    
    FormatEseloInfoFromResponseArray = response
End Function
