Attribute VB_Name = "modCustas"
Option Explicit

Sub DetectarProvGerarDaje(control As IRibbonControl)
''
'' Verifica a providência e chama a função correspondente.
''
    Dim strProvidencia As String
    Dim plan As Worksheet
    Dim contLinha As Long
    
    Set plan = ActiveSheet
    contLinha = ActiveCell.Row
    
    strProvidencia = plan.Cells(contLinha, 4).text
    
    Select Case strProvidencia
    Case "Emitir DAJE - Projudi"
        ContarEGerarDajesCognicaoProjudi
        
    Case "Emitir DAJE - Projudi - Execução"
        ContarEGerarDajesExecucaoProjudi
        
    Case "Emitir DAJE - PJe, eSAJ e outros sistemas", "Emitir DAJE - Cobrança"
        GerarDajesSemContar False
        
    Case "Emitir DAJE de desarquivamento"
        GerarDajeDesarquivamento control
    
    Case Else
        MsgBox "Sinto muito, " & DeterminarTratamento & "! O comando escolhido só serve para gerar DAJEs de Projudi ou de desarquivamento. " & _
                "Será que um destes não vos satisfaria?", vbInformation + vbOKOnly, "Sísifo em treinamento"
    
    End Select
    
End Sub

Sub ContarEGerarDajesCognicaoProjudi(Optional ByVal control As IRibbonControl)
    Dim IE As InternetExplorer
    Dim actsInfo As TaxableActsInfo
    
    Set actsInfo.celulaProcesso = ActiveCell.Worksheet.Cells(ActiveCell.Row, 1)
    actsInfo = ContarDajes(IE)
    actsInfo.faseExecucao = False
    GerarDajes actsInfo, IE
End Sub

Sub ContarEGerarDajesExecucaoProjudi(Optional ByVal control As IRibbonControl)
    Dim IE As InternetExplorer
    Dim actsInfo As TaxableActsInfo
    
    Set actsInfo.celulaProcesso = ActiveCell.Worksheet.Cells(ActiveCell.Row, 1)
    actsInfo = ContarDajes(IE)
    actsInfo.faseExecucao = True
    GerarDajes actsInfo, IE
End Sub

Sub GerarDajesSemContar(faseExecucao As Boolean)
    Dim actsInfo As TaxableActsInfo
    
    Set actsInfo.celulaProcesso = ActiveCell.Worksheet.Cells(ActiveCell.Row, 1)
    GerarDajes actsInfo
End Sub

Sub GerarDajesOutrosSistemas()
    Dim actsInfo As TaxableActsInfo
    actsInfo.celulaProcesso
    
    GerarDajes actsInfo
End Sub

Function ContarDajes(ByRef IE As InternetExplorer) As TaxableActsInfo
    Dim htmlDoc As HTMLDocument
    Dim tbTabela As HTMLTable, trCont As HTMLTableRow
    Dim resp As TaxableActsInfo
    Dim perfilUsuario As String, urlProcesso As String, strCont As String, iAndamento As String, iObsAndamento As String
    Dim advs As String
    Dim eventoInicial As Integer, numEvento As Integer
    Dim intCont As Integer
    Dim bolSoHtmlMp3 As Boolean
    Dim answer As VbMsgBoxResult
    
    'Descobre se o perfil é de parte ou advogado, em seguida cria nova janela para o processo
    Set IE = RecuperarIE("projudi.tjba.jus.br")
    perfilUsuario = DescobrirPerfil(IE.document)
    
    Set IE = New InternetExplorer
    IE.Visible = False
    
    'Pegar link pelo número CNJ
    Set resp.celulaProcesso = ActiveCell.Worksheet.Cells(ActiveCell.Row, 1)
    urlProcesso = PegaLinkProcesso(resp.celulaProcesso.Formula, IIf(perfilUsuario = "Advogado", True, False), IE, htmlDoc)
        
    Select Case urlProcesso
    Case "Sessão expirada"
        IE.Quit
        resp.celulaProcesso.Interior.Color = 9420794
        MsgBox DeterminarTratamento & ", a sessão expirou. Faça login no Projudi e tente novamente.", vbCritical + vbOKOnly, "Sísifo - Sessão do Projudi expirada"
        Exit Function
    Case "Processo não encontrado"
        IE.Quit
        resp.celulaProcesso.Interior.Color = 9420794
        MsgBox DeterminarTratamento & ", o processo não foi encontrado. Verifique se o número está correto e tente novamente.", vbCritical + vbOKOnly, "Sísifo - Processo não encontrado"
        Exit Function
    Case "Processo não encontrado ou perfil sem acesso"
        IE.Quit
        resp.celulaProcesso.Interior.Color = 9420794
        MsgBox DeterminarTratamento & ", o processo " & resp.celulaProcesso.Formula & " não foi encontrado. Talvez o número esteja errado, ou o perfil do " & _
            "usuário atualmente logado no Projudi não consiga acessá-lo. Imploro que confira o número ou tente novamente com outro usuário.", _
            vbCritical + vbOKOnly, "Sísifo - Processo não encontrado"
        Exit Function
    Case "Mais de um processo encontrado"
        IE.Quit
        resp.celulaProcesso.Interior.Color = 9420794
        MsgBox DeterminarTratamento & ", foi encontrado mais de um processo para o número " & resp.celulaProcesso.Formula & ". Isso é completamente inesperado! " & _
            "Suplico que confira o número e tente novamente.", vbCritical + vbOKOnly, "Sísifo - Mais de um processo encontrado"
        Exit Function
    End Select
    
    IE.navigate urlProcesso & "&consentimentoAcesso=true"
    
    Do
        DoEvents
    Loop Until IE.readyState = READYSTATE_COMPLETE
    
    'Se for segredo de justiça, avisa e para tudo
    If htmlDoc.body.Children(2).innerText = "Processo sob Segredo de Justiça" Then
        MsgBox DeterminarTratamento & ", o processo está em segredo de justiça. Tente novamente com um usuário com acesso.", vbCritical + vbOKOnly, "Sísifo - Processo em segredo de justiça"
        On Error Resume Next
        IE.Quit
        On Error GoTo 0
        resp.celulaProcesso.Interior.Color = 9420794
        Exit Function
    End If
    
    'Expande os botões de arquivos e observações de andamento
    ExpandirBotoesProcesso IE, htmlDoc
    
    'Conta a quantidade de partes
    resp.litisconsortes = PegarQtdPartes(IE, htmlDoc) - 2
    
    'Pega os advogados que já atuaram no processo nos históricos de advogados, para ajudar a contar as digitalizações
    advs = PegarHistoricoDeAdvogados(IE, htmlDoc)
    
    'Vai para o primeiro evento e os itera, contando os atos a pagar
    Set tbTabela = htmlDoc.getElementById("Arquivos").Children(0).Children(0)
    
    'Pergunta a partir de quando contar
    If tbTabela.Children.length <= btQtdEventosPossivelExecucao + 1 Then
        eventoInicial = 1
    Else
        eventoInicial = CInt(InputBox(DeterminarTratamento & ", há muitos eventos, talvez o processo seja uma execução! Devo contar os atos a partir de qual evento? Digite ""1"" para contar os atos do processo inteiro", "Sísifo - Contagem de atos", "1"))
    End If
    
    For intCont = 1 To tbTabela.Children.length - eventoInicial ' Desconsidera o "0", que seria a linha de cabeçalho dos eventos. Só vai até o evento inicial.
    
        numEvento = tbTabela.Children(intCont).Children(0).Children(0).Children(0).Children(0).Children(0).innerText
        iAndamento = tbTabela.Children(intCont).Children(0).Children(0).Children(0).Children(0).Children(1).Children(0).Children(0).innerText
        iObsAndamento = Mid(tbTabela.Children(intCont).Children(0).Children(0).Children(0).Children(0).Children(1).innerHTML, _
                                InStr(1, tbTabela.Children(intCont).Children(0).Children(0).Children(0).Children(0).Children(1).innerHTML, " <br>") + 5)
        iObsAndamento = Trim(Replace(iObsAndamento, Chr(10), ""))
        strCont = tbTabela.Children(intCont).Children(0).Children(0).Children(0).Children(0).Children(3).innerText
        
        'Aumenta as variáveis, de acordo com o evento
        If iAndamento = "Citação expedido(a)" Or iAndamento = "Intimação expedido(a)" Then
            If tbTabela.Children(intCont).Children(0).Children(2).Children(0).innerText = "Movimentação sem arquivos" Then ' Se não tiver arquivos, é eletrônica
                If Left(iObsAndamento, 8) <> "(P/ Advg" Then
                    resp.comEletronicas = resp.comEletronicas + 1 'Se for para advogado não paga, segundo o manual.
                End If
            Else ' Se tiver, é postal
                resp.comPostais = resp.comPostais + 1
                tbTabela.Children(intCont).Children(0).Children(0).Children(0).Children(0).Children(4).Children(0).Children(0).Children(0).Children(0).Children(0).Children(0).Children(0).Click
            End If
            
        ElseIf iAndamento = "Expedição de Mandado" Then
            resp.comMandados = resp.comMandados + 1
        
        ElseIf InStr(1, LCase(iAndamento), "precat") <> 0 Then
            resp.eventosPrecatorias = resp.eventosPrecatorias & numEvento & ", "
            
        ElseIf InStr(1, LCase(iAndamento), "competência declinada") <> 0 Then
            resp.eventosConfComp = resp.eventosConfComp & numEvento & ", "
            
        ElseIf InStr(1, LCase(iAndamento), "penhora") <> 0 Then
            resp.eventosPenhoras = resp.eventosPenhoras & numEvento & ", "
        
        ElseIf InStr(1, LCase(iAndamento), "impugnação de cálculo") = 0 And _
            (InStr(1, LCase(iAndamento), "cálculo") <> 0 Or InStr(1, LCase(iAndamento), "contadoria") <> 0 Or _
            InStr(1, LCase(iObsAndamento), "cálculo") <> 0 Or InStr(1, LCase(iObsAndamento), "contadoria")) <> 0 Then
            
            resp.eventosCalculos = resp.eventosCalculos & numEvento & ", "
            
        ' Guarda evento se não for protocolado por advogados da Embasa, nem por advogados nos históricos, nem por agentes eletrônicos
        '   do sistema Projudi, nem forem retornos de mandados...
        ElseIf InStr(1, advs, strCont & ",") = 0 And InStr(1, strAdvsEmbasa, strCont & ",") = 0 _
            And InStr(1, strAgentesAutomaticosProjudi, strCont & ",") = 0 And InStr(1, LCase(iAndamento), "mandado") = 0 _
            And Not tbTabela.Children(intCont).Children(0).Children(2).Children(0).innerText = "Movimentação sem arquivos" Then
            
            'Iterar os arquivos do andamento para testar se não é apenas uma combinação de "online.html" e arquivos mp3
            bolSoHtmlMp3 = True
            For Each trCont In tbTabela.Children(intCont).Children(0).Children(2).Children(0).Children(0).Children(0).Children(0).Children
                If Not (trCont.Children(trCont.Children.length - 1).Children(0).text = "online.html" Or Right(trCont.Children(trCont.Children.length - 1).Children(0).text, 4) = ".mp3") Then
                    bolSoHtmlMp3 = False
                    Exit For
                End If
            Next trCont
            
            ' Se não for apenas o arquivo online.html ou .mp3, adiciona às possíveis digitalizações
            If bolSoHtmlMp3 = False Then
                resp.eventosDigitalizacoes = resp.eventosDigitalizacoes & numEvento & ", "
            Else
                 If iAndamento = "Juntada de Intimação Telefônica" Then
                    resp.comEletronicas = resp.comEletronicas + 1
                End If
            End If
            
    
        ' Se for documento de advogado do processo, da Embasa ou de agente eletrônico do Projudi, fecha
        ElseIf (InStr(1, advs, strCont & ",") <> 0 Or InStr(1, strAdvsEmbasa, strCont & ",") <> 0 Or InStr(1, strAgentesAutomaticosProjudi, strCont & ",") <> 0) _
            And Not tbTabela.Children(intCont).Children(0).Children(2).Children(0).innerText = "Movimentação sem arquivos" Then
            
                tbTabela.Children(intCont).Children(0).Children(0).Children(0).Children(0).Children(4).Children(0).Children(0).Children(0).Children(0).Children(0).Children(0).Children(0).Click
                
        End If
        
    Next intCont
    
    If resp.eventosDigitalizacoes <> "" Then _
        resp.eventosDigitalizacoes = InputBox(DeterminarTratamento & ", acho que os seguintes eventos podem ter digitalizacoes:" & vbCrLf & _
            Left(resp.eventosDigitalizacoes, Len(resp.eventosDigitalizacoes) - 2) & vbCrLf & vbCrLf & _
            "Por favor, confira-os e me diga quantos destes eventos realmente contêm digitalizações:", _
            "Sísifo - Contar digitalizações", Len(resp.eventosDigitalizacoes) - Len(Replace(resp.eventosDigitalizacoes, ",", "")))
    
    If resp.eventosCalculos <> "" Then _
        resp.eventosCalculos = InputBox(DeterminarTratamento & ", acho que os seguintes eventos podem ter cálculos judiciais:" & vbCrLf & _
            Left(resp.eventosCalculos, Len(resp.eventosCalculos) - 2) & "." & vbCrLf & vbCrLf & _
            "Por favor, confira-os e me diga quantos destes eventos realmente contêm cálculos:", "Sísifo - Contar cálculos", "1")
    
    If resp.eventosPenhoras <> "" Then _
        resp.eventosPenhoras = InputBox(DeterminarTratamento & ", acho que os seguintes eventos podem ter solicitações de informações judiciais (penhoras, " & _
            "bacenjud, infojud, etc):" & vbCrLf & Left(resp.eventosPenhoras, Len(resp.eventosPenhoras) - 2) & "." & vbCrLf & vbCrLf & _
            "Por favor, confira-os e me diga quantos destes eventos realmente contêm penhoras:", "Sísifo - Contar penhoras", "1")
    
    If resp.eventosPrecatorias <> "" Then _
        resp.eventosPrecatorias = InputBox(DeterminarTratamento & ", acho que os seguintes eventos podem ter cartas precatórias:" & vbCrLf & _
            Left(resp.eventosPrecatorias, Len(resp.eventosPrecatorias) - 2) & "." & vbCrLf & vbCrLf & _
            "Por favor, confira-os e me diga quantos destes eventos realmente contêm cartas precatórias:", "Sísifo - Contar precatórias", "1")
    
    If resp.eventosConfComp <> "" Then
        answer = MsgBox(DeterminarTratamento & ", acho que os seguintes eventos podem ter conflitos de competência:" & vbCrLf & _
            Left(resp.eventosConfComp, Len(resp.eventosConfComp) - 2) & "." & vbCrLf & vbCrLf & _
            "Por favor, confira-os e me diga: algum desses eventos contém conflito de competência?", vbQuestion + vbYesNo, "Sísifo - Contar conflitos de competência")
        resp.possuiConfComp = IIf(answer = vbYes, True, False)
    End If
    
    ContarDajes = resp

End Function
    
Sub GerarDajes(actsInfo As TaxableActsInfo, Optional IE As InternetExplorer)
    Dim form As New frmRelatorioAtos
    Dim info As DajeGenerationInfo
    Dim intCont As Integer
    
    'Apresenta num formulário pra confirmação
    With form
        .txtNumProc = actsInfo.celulaProcesso.Formula
        .txtAdverso = actsInfo.celulaProcesso.Offset(0, 1).Formula
        .txtComEletronicas.text = actsInfo.comEletronicas
        .txtComPostais.text = actsInfo.comPostais
        .txtComMandados.text = actsInfo.comMandados
        .txtLitisconsortes.text = actsInfo.litisconsortes
        .txtDigitalizacoes.text = IIf(actsInfo.eventosDigitalizacoes = "", 0, actsInfo.eventosDigitalizacoes)
        .txtCalculos.text = IIf(actsInfo.eventosCalculos = "", 0, actsInfo.eventosCalculos)
        .txtPenhoras.text = IIf(actsInfo.eventosPenhoras = "", 0, actsInfo.eventosPenhoras)
        .txtPrecatorias.text = IIf(actsInfo.eventosPrecatorias = "", 0, actsInfo.eventosPrecatorias)
        .chbConfComp.value = IIf(actsInfo.possuiConfComp = True, True, False)
        .chbSentencaEmbargos.value = IIf(actsInfo.faseExecucao = True, True, False)
        .txtValor.SetFocus
        .txtValor.SelStart = 0
        .txtValor.SelLength = Len(.txtValor.text)
        
        form.Show
        
        On Error Resume Next
        IE.Quit
        On Error GoTo 0
        
        If form.chbDeveGerar.value = False Then Exit Sub
        
        info = PegarInformacoesDaje(actsInfo.celulaProcesso)
        
        If info.juizosComarcas(0).juizo = "" Then 'Se o juízo não estiver cadastrado, avisa e para.
            MsgBox DeterminarTratamento & ", o juízo """ & actsInfo.celulaProcesso.Worksheet.Cells(actsInfo.celulaProcesso.Row, 11).Formula & """ não está cadastrado na minha base de dados. Favor cadastrá-lo e tentar novamente.", vbCritical + vbOKOnly, "Sísifo - Juízo não cadastrado"
            Exit Sub
        End If

        'Gera e baixa os DAJEs.
        If .chbNaoGerarValorDaCausa.value = False Then
            info.atribuicaoDaje = "PROCESSOS_EM_GERAL"
            info.tipoAto = "I - DAS CAUSAS EM GERAL"
            info.valorAto = Replace(.txtValor.text, ".", "")
            info.eventosAtos = info.valorAto
            GerarUmDAJE info
            info = ResetVariableInfo(info)
        End If
        If .optRi.value = True Then
            info.atribuicaoDaje = "RECURSOS"
            info.tipoAto = "XXVII - RECURSOS (EXCLUÍDAS DESPESAS COM PORTE E REMESSA E/OU RETORNO, QUANDO CABÍVEIS) C) RECURSO INOMINADO (JUIZADOS ESPECIAIS)"
            GerarUmDAJE info
            info = ResetVariableInfo(info)
        End If
        If .optApelacao.value = True Then
            info.atribuicaoDaje = "RECURSOS"
            info.tipoAto = "XXVII - RECURSOS (EXCLUÍDAS DESPESAS COM PORTE E REMESSA E/OU RETORNO, QUANDO CABÍVEIS) A) APELAÇÃO, RECURSO ADESIVO"
            info.valorAto = Replace(.txtValor.text, ".", "")
            info.eventosAtos = info.valorAto
            GerarUmDAJE info
            info = ResetVariableInfo(info)
        End If
        If .txtComEletronicas.text <> 0 Then
            info.atribuicaoDaje = "PROCESSOS_EM_GERAL"
            info.tipoAto = "XXVI - ENVIO ELETRÔNICO DE CITAÇÕES, INTIMAÇÕES, OFÍCIOS E NOTIFICAÇÕES."
            info.atosQt = .txtComEletronicas.text
            info.eventosAtos = actsInfo.comEletronicas
            GerarUmDAJE info
            info = ResetVariableInfo(info)
        End If
        If .txtComPostais.text <> 0 Then
            info.atribuicaoDaje = "DESPESAS_JUDICIAIS_EXTRAJUDICIAIS"
            info.tipoAto = "III - TARIFA DE POSTAGEM - CITAÇÃO OU INTIMAÇÃO VIA POSTAL"
            info.atosQt = .txtComPostais.text
            info.eventosAtos = actsInfo.comPostais
            GerarUmDAJE info
            info = ResetVariableInfo(info)
        End If
        If .txtComMandados.text <> 0 Then
            info.atribuicaoDaje = "ATOS_DOS_OFICIAIS_JUSTICA"
            info.tipoAto = "XXVIII - CITAÇÃO, INTIMAÇÃO, NOTIFICAÇÃO E ENTREGA DE OFÍCIO"
            info.atosQt = .txtComMandados.text
            info.eventosAtos = actsInfo.comMandados
            GerarUmDAJE info
            info = ResetVariableInfo(info)
        End If
        If .txtLitisconsortes.text <> 0 Then
            info.atribuicaoDaje = "PROCESSOS_EM_GERAL"
            info.tipoAto = "VII - LITISCONSÓRCIO ATIVO OU PASSIVO, POR PARTE EXCEDENTE"
            info.atosQt = .txtLitisconsortes.text
            GerarUmDAJE info
            info = ResetVariableInfo(info)
        End If
        If .txtDigitalizacoes.text <> 0 Then
            info.atribuicaoDaje = "PROCESSOS_EM_GERAL"
            info.tipoAto = "XXI - DIGITALIZAÇÃO DE DOCUMENTO REALIZADA NO ÂMBITO DESTE PODER JUDICIÁRIO, POR DOCUMENTO (DENTRE ELES, A DIGITALIZAÇÃO DE PETIÇÃO, INLCUINDO-SE OS DOCUMENTOS ANEXADOS A ESTA, ENDEREÇADA A PROCESSO ELETRÔNICO POR MEIO FÍSICO, I.E., PAPEL)"
            info.atosQt = .txtDigitalizacoes.text
            info.eventosAtos = actsInfo.eventosDigitalizacoes
            GerarUmDAJE info
            info = ResetVariableInfo(info)
        End If
        If .txtCalculos.text <> 0 Then
            For intCont = 1 To .txtCalculos.text Step 1
                info.atribuicaoDaje = "ATOS_DOS_OFICIAIS_JUSTICA"
                info.tipoAto = "XVIII - AVALIAÇÕES E CÁLCULOS JUDICIAIS, POR MANDADO"
                info.eventosAtos = actsInfo.eventosCalculos
                GerarUmDAJE info
            Next intCont
            info = ResetVariableInfo(info)
        End If
        If .txtPenhoras.text <> 0 Then
            info.atribuicaoDaje = "PROCESSOS_EM_GERAL"
            info.tipoAto = "XIX - REQUISIÇÃO DE INFORMAÇÕES POR MEIO ELETRÔNICO (BACENJUD, RENAJUD, INFOJUD, SERASAJUD E ASSEMELHADOS), POR CADA CONSULTA"
            info.atosQt = .txtPenhoras.text
            info.eventosAtos = actsInfo.eventosPenhoras
            GerarUmDAJE info
            info = ResetVariableInfo(info)
        End If
        If .chbConfComp.value = True Then
            info.atribuicaoDaje = "PROCESSOS_EM_GERAL"
            info.tipoAto = "IV - EXCEÇÃO DE IMPEDIMENTO E SUSPEIÇÃO DOS JUÍZES, CONFLITO DE COMPETÊNCIA OU DE JURISDIÇÃO SUSCITADOS PELA PARTE - DESAFORAMENTO."
            info.eventosAtos = actsInfo.eventosConfComp
            GerarUmDAJE info
            info = ResetVariableInfo(info)
        End If
        If .chbSentencaEmbargos.value = True Then
            info.atribuicaoDaje = "PROCESSOS_EM_GERAL"
            info.tipoAto = "XV - DEMAIS PROCESSOS OU PROCEDIMENTOS SEM VALOR DECLARADO, INCLUSIVE INCIDENTAIS, E DE IMPUGNAÇÕES EM GERAL"
            GerarUmDAJE info
            info = ResetVariableInfo(info)
        End If
        If .txtPrecatorias.text <> 0 Then
            For intCont = 1 To .txtPrecatorias.text Step 1
                info.atribuicaoDaje = "PROCESSOS_EM_GERAL"
                info.tipoAto = "VI - CARTA PRECATORIA, DE ORDEM E ROGATORIA, INCLUIDO PORTE DE RETORNO."
                info.eventosAtos = actsInfo.eventosPrecatorias
                GerarUmDAJE info
            Next intCont
            info = ResetVariableInfo(info)
        End If
    End With
    
    actsInfo.celulaProcesso.Interior.Color = 14857357
    actsInfo.celulaProcesso.Offset(1, 0).Select
End Sub

Function ResetVariableInfo(info As DajeGenerationInfo) As DajeGenerationInfo
    info.atosQt = 0
    info.atribuicaoDaje = ""
    info.tipoAto = ""
    info.valorAto = ""
    info.eventosAtos = ""
    ResetVariableInfo = info
End Function


Sub GerarDajeDesarquivamento(control As IRibbonControl)
    Dim rngProcesso As Excel.Range
    Dim info As DajeGenerationInfo
    
    Set rngProcesso = ActiveCell.Worksheet.Cells(ActiveCell.Row, 1)
    info = PegarInformacoesDaje(rngProcesso)
    If info.juizosComarcas(0).juizo = "" Then 'Se o juízo não estiver cadastrado, avisa e para.
        MsgBox DeterminarTratamento & ", o juízo """ & rngProcesso.Worksheet.Cells(rngProcesso.Row, 11).Formula & """ não está cadastrado na minha base de dados. Favor cadastrá-lo e tentar novamente.", vbCritical + vbOKOnly, "Sísifo - Juízo não cadastrado"
       Exit Sub
    End If
    
    info.atribuicaoDaje = "PROCESSOS_EM_GERAL"
    info.tipoAto = "XVI - DESARQUIVAMENTO DE PROCESSOS, INCLUSIVE ELETRÔNICOS, POR PROCESSO"
    GerarUmDAJE info
    
    rngProcesso.Interior.Color = 14857357
    rngProcesso.Offset(1, 0).Select
End Sub

Function PegarInformacoesDaje(rngProcesso As Excel.Range) As DajeGenerationInfo
    Dim resposta As DajeGenerationInfo
    Dim eseloInfos() As EseloInfo
    Dim juizoEspaider As String, comarcaEspaider As String
    
    resposta.numProcesso = Trim(rngProcesso.Formula)
    resposta.Adverso = Trim(rngProcesso.Offset(0, 1).Formula)
    
    ' Pega o Juízo na redação do Espaider, depois pega as informações na redação do Eselo.
    juizoEspaider = rngProcesso.Worksheet.Cells(rngProcesso.Row, 11).Formula ' Juízo na redação Espaider
    eseloInfos = BuscaJuizoEselo(juizoEspaider)
    resposta.juizosComarcas = eseloInfos
    resposta.juizosComarcasOriginal = eseloInfos

    ' Cria a pasta e configura o Chrome para salvar os DAJEs na pasta
    resposta.diretorio = CaminhoDesktop & "\Sisifo DAJEs\" & resposta.numProcesso & "\"
    If Dir(CaminhoDesktop & "\Sisifo DAJEs", vbDirectory) = "" Then MkDir CaminhoDesktop & "\Sisifo DAJEs\"
    If Dir(resposta.diretorio, vbDirectory) = "" Then MkDir resposta.diretorio

    
    PegarInformacoesDaje = resposta

End Function
