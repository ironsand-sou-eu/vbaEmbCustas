Attribute VB_Name = "modCustas"
Option Explicit

Sub DetectarProvGerarDaje(control As IRibbonControl)
''
'' Verifica a provid�ncia e chama a fun��o correspondente.
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
        
    Case "Emitir DAJE - Projudi - Execu��o"
        ContarEGerarDajesExecucaoProjudi
        
    Case "Emitir DAJE - PJe, eSAJ e outros sistemas", "Emitir DAJE - Cobran�a"
        GerarDajesSemContar False
        
    Case "Emitir DAJE de desarquivamento"
        GerarDajeDesarquivamento control
    
    Case Else
        MsgBox "Sinto muito, " & DeterminarTratamento & "! O comando escolhido s� serve para gerar DAJEs de Projudi ou de desarquivamento. " & _
                "Ser� que um destes n�o vos satisfaria?", vbInformation + vbOKOnly, "S�sifo em treinamento"
    
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
    
    'Descobre se o perfil � de parte ou advogado, em seguida cria nova janela para o processo
    Set IE = RecuperarIE("projudi.tjba.jus.br")
    perfilUsuario = DescobrirPerfil(IE.document)
    
    Set IE = New InternetExplorer
    IE.Visible = False
    
    'Pegar link pelo n�mero CNJ
    Set resp.celulaProcesso = ActiveCell.Worksheet.Cells(ActiveCell.Row, 1)
    urlProcesso = PegaLinkProcesso(resp.celulaProcesso.Formula, IIf(perfilUsuario = "Advogado", True, False), IE, htmlDoc)
        
    Select Case urlProcesso
    Case "Sess�o expirada"
        IE.Quit
        resp.celulaProcesso.Interior.Color = 9420794
        MsgBox DeterminarTratamento & ", a sess�o expirou. Fa�a login no Projudi e tente novamente.", vbCritical + vbOKOnly, "S�sifo - Sess�o do Projudi expirada"
        Exit Function
    Case "Processo n�o encontrado"
        IE.Quit
        resp.celulaProcesso.Interior.Color = 9420794
        MsgBox DeterminarTratamento & ", o processo n�o foi encontrado. Verifique se o n�mero est� correto e tente novamente.", vbCritical + vbOKOnly, "S�sifo - Processo n�o encontrado"
        Exit Function
    Case "Processo n�o encontrado ou perfil sem acesso"
        IE.Quit
        resp.celulaProcesso.Interior.Color = 9420794
        MsgBox DeterminarTratamento & ", o processo " & resp.celulaProcesso.Formula & " n�o foi encontrado. Talvez o n�mero esteja errado, ou o perfil do " & _
            "usu�rio atualmente logado no Projudi n�o consiga acess�-lo. Imploro que confira o n�mero ou tente novamente com outro usu�rio.", _
            vbCritical + vbOKOnly, "S�sifo - Processo n�o encontrado"
        Exit Function
    Case "Mais de um processo encontrado"
        IE.Quit
        resp.celulaProcesso.Interior.Color = 9420794
        MsgBox DeterminarTratamento & ", foi encontrado mais de um processo para o n�mero " & resp.celulaProcesso.Formula & ". Isso � completamente inesperado! " & _
            "Suplico que confira o n�mero e tente novamente.", vbCritical + vbOKOnly, "S�sifo - Mais de um processo encontrado"
        Exit Function
    End Select
    
    IE.navigate urlProcesso & "&consentimentoAcesso=true"
    
    Do
        DoEvents
    Loop Until IE.readyState = READYSTATE_COMPLETE
    
    'Se for segredo de justi�a, avisa e para tudo
    If htmlDoc.body.Children(2).innerText = "Processo sob Segredo de Justi�a" Then
        MsgBox DeterminarTratamento & ", o processo est� em segredo de justi�a. Tente novamente com um usu�rio com acesso.", vbCritical + vbOKOnly, "S�sifo - Processo em segredo de justi�a"
        On Error Resume Next
        IE.Quit
        On Error GoTo 0
        resp.celulaProcesso.Interior.Color = 9420794
        Exit Function
    End If
    
    'Expande os bot�es de arquivos e observa��es de andamento
    ExpandirBotoesProcesso IE, htmlDoc
    
    'Conta a quantidade de partes
    resp.litisconsortes = PegarQtdPartes(IE, htmlDoc) - 2
    
    'Pega os advogados que j� atuaram no processo nos hist�ricos de advogados, para ajudar a contar as digitaliza��es
    advs = PegarHistoricoDeAdvogados(IE, htmlDoc)
    
    'Vai para o primeiro evento e os itera, contando os atos a pagar
    Set tbTabela = htmlDoc.getElementById("Arquivos").Children(0).Children(0)
    
    'Pergunta a partir de quando contar
    If tbTabela.Children.length <= btQtdEventosPossivelExecucao + 1 Then
        eventoInicial = 1
    Else
        eventoInicial = CInt(InputBox(DeterminarTratamento & ", h� muitos eventos, talvez o processo seja uma execu��o! Devo contar os atos a partir de qual evento? Digite ""1"" para contar os atos do processo inteiro", "S�sifo - Contagem de atos", "1"))
    End If
    
    For intCont = 1 To tbTabela.Children.length - eventoInicial ' Desconsidera o "0", que seria a linha de cabe�alho dos eventos. S� vai at� o evento inicial.
    
        numEvento = tbTabela.Children(intCont).Children(0).Children(0).Children(0).Children(0).Children(0).innerText
        iAndamento = tbTabela.Children(intCont).Children(0).Children(0).Children(0).Children(0).Children(1).Children(0).Children(0).innerText
        iObsAndamento = Mid(tbTabela.Children(intCont).Children(0).Children(0).Children(0).Children(0).Children(1).innerHTML, _
                                InStr(1, tbTabela.Children(intCont).Children(0).Children(0).Children(0).Children(0).Children(1).innerHTML, " <br>") + 5)
        iObsAndamento = Trim(Replace(iObsAndamento, Chr(10), ""))
        strCont = tbTabela.Children(intCont).Children(0).Children(0).Children(0).Children(0).Children(3).innerText
        
        'Aumenta as vari�veis, de acordo com o evento
        If iAndamento = "Cita��o expedido(a)" Or iAndamento = "Intima��o expedido(a)" Then
            If tbTabela.Children(intCont).Children(0).Children(2).Children(0).innerText = "Movimenta��o sem arquivos" Then ' Se n�o tiver arquivos, � eletr�nica
                If Left(iObsAndamento, 8) <> "(P/ Advg" Then
                    resp.comEletronicas = resp.comEletronicas + 1 'Se for para advogado n�o paga, segundo o manual.
                End If
            Else ' Se tiver, � postal
                resp.comPostais = resp.comPostais + 1
                tbTabela.Children(intCont).Children(0).Children(0).Children(0).Children(0).Children(4).Children(0).Children(0).Children(0).Children(0).Children(0).Children(0).Children(0).Click
            End If
            
        ElseIf iAndamento = "Expedi��o de Mandado" Then
            resp.comMandados = resp.comMandados + 1
        
        ElseIf InStr(1, LCase(iAndamento), "precat") <> 0 Then
            resp.eventosPrecatorias = resp.eventosPrecatorias & numEvento & ", "
            
        ElseIf InStr(1, LCase(iAndamento), "compet�ncia declinada") <> 0 Then
            resp.eventosConfComp = resp.eventosConfComp & numEvento & ", "
            
        ElseIf InStr(1, LCase(iAndamento), "penhora") <> 0 Then
            resp.eventosPenhoras = resp.eventosPenhoras & numEvento & ", "
        
        ElseIf InStr(1, LCase(iAndamento), "impugna��o de c�lculo") = 0 And _
            (InStr(1, LCase(iAndamento), "c�lculo") <> 0 Or InStr(1, LCase(iAndamento), "contadoria") <> 0 Or _
            InStr(1, LCase(iObsAndamento), "c�lculo") <> 0 Or InStr(1, LCase(iObsAndamento), "contadoria")) <> 0 Then
            
            resp.eventosCalculos = resp.eventosCalculos & numEvento & ", "
            
        ' Guarda evento se n�o for protocolado por advogados da Embasa, nem por advogados nos hist�ricos, nem por agentes eletr�nicos
        '   do sistema Projudi, nem forem retornos de mandados...
        ElseIf InStr(1, advs, strCont & ",") = 0 And InStr(1, strAdvsEmbasa, strCont & ",") = 0 _
            And InStr(1, strAgentesAutomaticosProjudi, strCont & ",") = 0 And InStr(1, LCase(iAndamento), "mandado") = 0 _
            And Not tbTabela.Children(intCont).Children(0).Children(2).Children(0).innerText = "Movimenta��o sem arquivos" Then
            
            'Iterar os arquivos do andamento para testar se n�o � apenas uma combina��o de "online.html" e arquivos mp3
            bolSoHtmlMp3 = True
            For Each trCont In tbTabela.Children(intCont).Children(0).Children(2).Children(0).Children(0).Children(0).Children(0).Children
                If Not (trCont.Children(trCont.Children.length - 1).Children(0).text = "online.html" Or Right(trCont.Children(trCont.Children.length - 1).Children(0).text, 4) = ".mp3") Then
                    bolSoHtmlMp3 = False
                    Exit For
                End If
            Next trCont
            
            ' Se n�o for apenas o arquivo online.html ou .mp3, adiciona �s poss�veis digitaliza��es
            If bolSoHtmlMp3 = False Then
                resp.eventosDigitalizacoes = resp.eventosDigitalizacoes & numEvento & ", "
            Else
                 If iAndamento = "Juntada de Intima��o Telef�nica" Then
                    resp.comEletronicas = resp.comEletronicas + 1
                End If
            End If
            
    
        ' Se for documento de advogado do processo, da Embasa ou de agente eletr�nico do Projudi, fecha
        ElseIf (InStr(1, advs, strCont & ",") <> 0 Or InStr(1, strAdvsEmbasa, strCont & ",") <> 0 Or InStr(1, strAgentesAutomaticosProjudi, strCont & ",") <> 0) _
            And Not tbTabela.Children(intCont).Children(0).Children(2).Children(0).innerText = "Movimenta��o sem arquivos" Then
            
                tbTabela.Children(intCont).Children(0).Children(0).Children(0).Children(0).Children(4).Children(0).Children(0).Children(0).Children(0).Children(0).Children(0).Children(0).Click
                
        End If
        
    Next intCont
    
    If resp.eventosDigitalizacoes <> "" Then _
        resp.eventosDigitalizacoes = InputBox(DeterminarTratamento & ", acho que os seguintes eventos podem ter digitalizacoes:" & vbCrLf & _
            Left(resp.eventosDigitalizacoes, Len(resp.eventosDigitalizacoes) - 2) & vbCrLf & vbCrLf & _
            "Por favor, confira-os e me diga quantos destes eventos realmente cont�m digitaliza��es:", _
            "S�sifo - Contar digitaliza��es", Len(resp.eventosDigitalizacoes) - Len(Replace(resp.eventosDigitalizacoes, ",", "")))
    
    If resp.eventosCalculos <> "" Then _
        resp.eventosCalculos = InputBox(DeterminarTratamento & ", acho que os seguintes eventos podem ter c�lculos judiciais:" & vbCrLf & _
            Left(resp.eventosCalculos, Len(resp.eventosCalculos) - 2) & "." & vbCrLf & vbCrLf & _
            "Por favor, confira-os e me diga quantos destes eventos realmente cont�m c�lculos:", "S�sifo - Contar c�lculos", "1")
    
    If resp.eventosPenhoras <> "" Then _
        resp.eventosPenhoras = InputBox(DeterminarTratamento & ", acho que os seguintes eventos podem ter solicita��es de informa��es judiciais (penhoras, " & _
            "bacenjud, infojud, etc):" & vbCrLf & Left(resp.eventosPenhoras, Len(resp.eventosPenhoras) - 2) & "." & vbCrLf & vbCrLf & _
            "Por favor, confira-os e me diga quantos destes eventos realmente cont�m penhoras:", "S�sifo - Contar penhoras", "1")
    
    If resp.eventosPrecatorias <> "" Then _
        resp.eventosPrecatorias = InputBox(DeterminarTratamento & ", acho que os seguintes eventos podem ter cartas precat�rias:" & vbCrLf & _
            Left(resp.eventosPrecatorias, Len(resp.eventosPrecatorias) - 2) & "." & vbCrLf & vbCrLf & _
            "Por favor, confira-os e me diga quantos destes eventos realmente cont�m cartas precat�rias:", "S�sifo - Contar precat�rias", "1")
    
    If resp.eventosConfComp <> "" Then
        answer = MsgBox(DeterminarTratamento & ", acho que os seguintes eventos podem ter conflitos de compet�ncia:" & vbCrLf & _
            Left(resp.eventosConfComp, Len(resp.eventosConfComp) - 2) & "." & vbCrLf & vbCrLf & _
            "Por favor, confira-os e me diga: algum desses eventos cont�m conflito de compet�ncia?", vbQuestion + vbYesNo, "S�sifo - Contar conflitos de compet�ncia")
        resp.possuiConfComp = IIf(answer = vbYes, True, False)
    End If
    
    ContarDajes = resp

End Function
    
Sub GerarDajes(actsInfo As TaxableActsInfo, Optional IE As InternetExplorer)
    Dim form As New frmRelatorioAtos
    Dim info As DajeGenerationInfo
    Dim intCont As Integer
    
    'Apresenta num formul�rio pra confirma��o
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
        
        If info.juizosComarcas(0).juizo = "" Then 'Se o ju�zo n�o estiver cadastrado, avisa e para.
            MsgBox DeterminarTratamento & ", o ju�zo """ & actsInfo.celulaProcesso.Worksheet.Cells(actsInfo.celulaProcesso.Row, 11).Formula & """ n�o est� cadastrado na minha base de dados. Favor cadastr�-lo e tentar novamente.", vbCritical + vbOKOnly, "S�sifo - Ju�zo n�o cadastrado"
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
            info.tipoAto = "XXVII - RECURSOS (EXCLU�DAS DESPESAS COM PORTE E REMESSA E/OU RETORNO, QUANDO CAB�VEIS) C) RECURSO INOMINADO (JUIZADOS ESPECIAIS)"
            GerarUmDAJE info
            info = ResetVariableInfo(info)
        End If
        If .optApelacao.value = True Then
            info.atribuicaoDaje = "RECURSOS"
            info.tipoAto = "XXVII - RECURSOS (EXCLU�DAS DESPESAS COM PORTE E REMESSA E/OU RETORNO, QUANDO CAB�VEIS) A) APELA��O, RECURSO ADESIVO"
            info.valorAto = Replace(.txtValor.text, ".", "")
            info.eventosAtos = info.valorAto
            GerarUmDAJE info
            info = ResetVariableInfo(info)
        End If
        If .txtComEletronicas.text <> 0 Then
            info.atribuicaoDaje = "PROCESSOS_EM_GERAL"
            info.tipoAto = "XXVI - ENVIO ELETR�NICO DE CITA��ES, INTIMA��ES, OF�CIOS E NOTIFICA��ES."
            info.atosQt = .txtComEletronicas.text
            info.eventosAtos = actsInfo.comEletronicas
            GerarUmDAJE info
            info = ResetVariableInfo(info)
        End If
        If .txtComPostais.text <> 0 Then
            info.atribuicaoDaje = "DESPESAS_JUDICIAIS_EXTRAJUDICIAIS"
            info.tipoAto = "III - TARIFA DE POSTAGEM - CITA��O OU INTIMA��O VIA POSTAL"
            info.atosQt = .txtComPostais.text
            info.eventosAtos = actsInfo.comPostais
            GerarUmDAJE info
            info = ResetVariableInfo(info)
        End If
        If .txtComMandados.text <> 0 Then
            info.atribuicaoDaje = "ATOS_DOS_OFICIAIS_JUSTICA"
            info.tipoAto = "XXVIII - CITA��O, INTIMA��O, NOTIFICA��O E ENTREGA DE OF�CIO"
            info.atosQt = .txtComMandados.text
            info.eventosAtos = actsInfo.comMandados
            GerarUmDAJE info
            info = ResetVariableInfo(info)
        End If
        If .txtLitisconsortes.text <> 0 Then
            info.atribuicaoDaje = "PROCESSOS_EM_GERAL"
            info.tipoAto = "VII - LITISCONS�RCIO ATIVO OU PASSIVO, POR PARTE EXCEDENTE"
            info.atosQt = .txtLitisconsortes.text
            GerarUmDAJE info
            info = ResetVariableInfo(info)
        End If
        If .txtDigitalizacoes.text <> 0 Then
            info.atribuicaoDaje = "PROCESSOS_EM_GERAL"
            info.tipoAto = "XXI - DIGITALIZA��O DE DOCUMENTO REALIZADA NO �MBITO DESTE PODER JUDICI�RIO, POR DOCUMENTO (DENTRE ELES, A DIGITALIZA��O DE PETI��O, INLCUINDO-SE OS DOCUMENTOS ANEXADOS A ESTA, ENDERE�ADA A PROCESSO ELETR�NICO POR MEIO F�SICO, I.E., PAPEL)"
            info.atosQt = .txtDigitalizacoes.text
            info.eventosAtos = actsInfo.eventosDigitalizacoes
            GerarUmDAJE info
            info = ResetVariableInfo(info)
        End If
        If .txtCalculos.text <> 0 Then
            For intCont = 1 To .txtCalculos.text Step 1
                info.atribuicaoDaje = "ATOS_DOS_OFICIAIS_JUSTICA"
                info.tipoAto = "XVIII - AVALIA��ES E C�LCULOS JUDICIAIS, POR MANDADO"
                info.eventosAtos = actsInfo.eventosCalculos
                GerarUmDAJE info
            Next intCont
            info = ResetVariableInfo(info)
        End If
        If .txtPenhoras.text <> 0 Then
            info.atribuicaoDaje = "PROCESSOS_EM_GERAL"
            info.tipoAto = "XIX - REQUISI��O DE INFORMA��ES POR MEIO ELETR�NICO (BACENJUD, RENAJUD, INFOJUD, SERASAJUD E ASSEMELHADOS), POR CADA CONSULTA"
            info.atosQt = .txtPenhoras.text
            info.eventosAtos = actsInfo.eventosPenhoras
            GerarUmDAJE info
            info = ResetVariableInfo(info)
        End If
        If .chbConfComp.value = True Then
            info.atribuicaoDaje = "PROCESSOS_EM_GERAL"
            info.tipoAto = "IV - EXCE��O DE IMPEDIMENTO E SUSPEI��O DOS JU�ZES, CONFLITO DE COMPET�NCIA OU DE JURISDI��O SUSCITADOS PELA PARTE - DESAFORAMENTO."
            info.eventosAtos = actsInfo.eventosConfComp
            GerarUmDAJE info
            info = ResetVariableInfo(info)
        End If
        If .chbSentencaEmbargos.value = True Then
            info.atribuicaoDaje = "PROCESSOS_EM_GERAL"
            info.tipoAto = "XV - DEMAIS PROCESSOS OU PROCEDIMENTOS SEM VALOR DECLARADO, INCLUSIVE INCIDENTAIS, E DE IMPUGNA��ES EM GERAL"
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
    If info.juizosComarcas(0).juizo = "" Then 'Se o ju�zo n�o estiver cadastrado, avisa e para.
        MsgBox DeterminarTratamento & ", o ju�zo """ & rngProcesso.Worksheet.Cells(rngProcesso.Row, 11).Formula & """ n�o est� cadastrado na minha base de dados. Favor cadastr�-lo e tentar novamente.", vbCritical + vbOKOnly, "S�sifo - Ju�zo n�o cadastrado"
       Exit Sub
    End If
    
    info.atribuicaoDaje = "PROCESSOS_EM_GERAL"
    info.tipoAto = "XVI - DESARQUIVAMENTO DE PROCESSOS, INCLUSIVE ELETR�NICOS, POR PROCESSO"
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
    
    ' Pega o Ju�zo na reda��o do Espaider, depois pega as informa��es na reda��o do Eselo.
    juizoEspaider = rngProcesso.Worksheet.Cells(rngProcesso.Row, 11).Formula ' Ju�zo na reda��o Espaider
    eseloInfos = BuscaJuizoEselo(juizoEspaider)
    resposta.juizosComarcas = eseloInfos
    resposta.juizosComarcasOriginal = eseloInfos

    ' Cria a pasta e configura o Chrome para salvar os DAJEs na pasta
    resposta.diretorio = CaminhoDesktop & "\Sisifo DAJEs\" & resposta.numProcesso & "\"
    If Dir(CaminhoDesktop & "\Sisifo DAJEs", vbDirectory) = "" Then MkDir CaminhoDesktop & "\Sisifo DAJEs\"
    If Dir(resposta.diretorio, vbDirectory) = "" Then MkDir resposta.diretorio

    
    PegarInformacoesDaje = resposta

End Function
