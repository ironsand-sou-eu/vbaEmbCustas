Attribute VB_Name = "modProjudi"
Public Const strAdvsEmbasa As String = "ADEVALDO DE SANTANA GOMES,ANA PAULA AMORIM CORTES," & _
                                    "AGLAY LIMA COSTA MACHADO PEDREIRA," & _
                                    "ANALYZ PESSOA BRAZ DE OLIVEIRA," & _
                                    "ANANDA ATMAN AZEVEDO DOS SANTOS," & _
                                    "ANGELA MOISES FARIA LANTYER," & _
                                    "CARLOS HENRIQUE MARTINS JUNIOR," & _
                                    "CESAR BRAGA LINS BAMBERG RODRIGUEZ,CRISTHIANO PAULO TEIXEIRA DE CASTRO," & _
                                    "DANILO BARRETO FEDULO DE ALMEIDA,ELISANGELA DE QUEIROZ FERNANDES BRITO," & _
                                    "FABIO JUNIO SOUZA OLIVEIRA," & _
                                    "FERNANDA BARRETO MOTA," & _
                                    "GENYSSON SANTOS ARAUJO,GILDEMAR BITTENCOURT SANTOS SILVA," & _
                                    "IZABELA RIOS LEITE," & _
                                    "JAIRO BRAGA LIMA,JEFFERSON MESSIAS," & _
                                    "JORGE KIDELMIR NASCIMENTO DE OLIVEIRA FILHO," & _
                                    "JULIANA CARDOSO NASCIMENTO," & _
                                    "LEIDIANE CARVALHO FRAGA MAGALHAES," & _
                                    "LIVIA MOURA MARQUES DE OLIVEIRA," & _
                                    "LIVIA REGINA OLIVEIRA DE SOUZA," & _
                                    "MARCOS MOTA DE ALMEIDA FILHO,MARIA QUINTAS RADEL," & _
                                    "MARIANA BRASIL NOGUEIRA LIMA,MARIVALDO SILVA NETTO," & _
                                    "MILA LEITE NASCIMENTO," & _
                                    "PEDRO CAMERA PACHECO,ROMULO RAMOS DONATO," & _
                                    "SILVIA DE MATOS CARVALHO MATINELLI," & _
                                    "TANIA MARIA REBOUCAS," 'Agentes que ser�o ignorados ao contar digitaliza��es

Public Const SISIFO_URL As String = "https://embasa.sisifo.tec.br/api/"
Public Const strAgentesAutomaticosProjudi As String = "ECT,SISTEMA CNJ," 'Agentes que ser�o ignorados ao contar digitaliza��es
Public Const btQtdEventosPossivelExecucao As Byte = 60 'Quantidade de eventos at� a qual o sistema presume que n�o � poss�vel ser um RI de execu��o.

Function PegaLinkProcesso(ByVal strNumeroCNJ As String, ByVal bolLoginDeAdvogado As Boolean, ByRef IE As InternetExplorer, ByRef DocHTML As HTMLDocument) As String
''
'' Retorna o link da p�gina principal do processo strNumeroCNJ.
'' DEVO LIDAR COM O ERRO DE N�O ESTAR LOGADO!!!!!!!
''

    Dim strContNumeroProcesso As String
    Dim intCont As Integer
    Dim btContLinkProcesso As Byte

    IE.Visible = True
    IE.navigate IIf(bolLoginDeAdvogado = True, "https://projudi.tjba.jus.br/projudi/buscas/ProcessosQualquerAdvogado", "https://projudi.tjba.jus.br/projudi/buscas/ProcessosParte")
        
    Set IE = RecuperarIE(IIf(bolLoginDeAdvogado = True, "https://projudi.tjba.jus.br/projudi/buscas/ProcessosQualquerAdvogado", "https://projudi.tjba.jus.br/projudi/buscas/ProcessosParte"))
    If IE Is Nothing Then
        On Error Resume Next
        IE.Quit
        On Error GoTo 0
        PegaLinkProcesso = "N�o abriu por demora"
        Exit Function
    End If
    IE.Visible = True
    
    On Error GoTo Volta1
Volta1:
    Do
        DoEvents
    Loop Until IE.document.readyState = "complete"
    
    Do
        DoEvents
    Loop Until IE.document.getElementsByTagName("body")(0).Children(2).Children(0).Children(0).Children(0).Children(1).Children(0).innerText = "N�mero Processo"
    On Error GoTo 0
    
    ' Preenche o n�mero do processo na busca e submete o formul�rio
    Set DocHTML = IE.document
    
    If DocHTML.Title = "Sistema CNJ - A sess�o expirou" Then
        PegaLinkProcesso = "Sess�o expirada"
        Exit Function
    End If
    
    DocHTML.getElementById("numeroProcesso").value = strNumeroCNJ
    DocHTML.forms("busca").submit
    
    'Esperar 1
    ' No futuro: observar a requisi��o, para ver que valores j� voltam preenchidos e quais s�o criados de forma ass�ncrona, a� testar bom base em algum ass�ncrono.
    On Error GoTo Volta2
Volta2:
    Do
        DoEvents
    Loop Until IE.readyState = 4
    
    Do
        DoEvents
    Loop Until DocHTML.images(DocHTML.images.length - 1).href = "https://projudi.tjba.jus.br/projudi/imagens/botoes/bot-Imprimir.gif"
    
    If bolLoginDeAdvogado = True Then
        intCont = DocHTML.forms(1).Children(0).Children(0).Children.length
        btContLinkProcesso = 4
    Else
        intCont = DocHTML.getElementsByName("formProcessos")(0).getElementsByTagName("table")(0).getElementsByTagName("tbody")(0).getElementsByTagName("tr").length
        btContLinkProcesso = 3
    End If
    
    Select Case intCont
    Case 3 'Nenhum processo, s� os cabe�alhos da tabela
        PegaLinkProcesso = "Processo n�o encontrado ou perfil sem acesso"
    Case btContLinkProcesso + 1 'Um processo s�
        If DocHTML.getElementsByName("formProcessos")(0).getElementsByTagName("table")(0).getElementsByTagName("tbody")(0).getElementsByTagName("tr")(3).getElementsByTagName("td")(1).getElementsByTagName("a")(0).innerText = strNumeroCNJ Then
            PegaLinkProcesso = DocHTML.getElementsByName("formProcessos")(0).getElementsByTagName("table")(0).getElementsByTagName("tbody")(0).getElementsByTagName("tr")(3).getElementsByTagName("td")(1).getElementsByTagName("a")(0).href
        End If
    Case Is > btContLinkProcesso + 1 'Mais de um processo
        PegaLinkProcesso = "Mais de um processo encontrado"
    End Select
    
    'intCont = DocHTML.getElementsByTagName("a").length - 1
    'For intCont = 0 To intCont Step 1
    '    If DocHTML.getElementsByTagName("a")(intCont).innerText = strNumeroCNJ Then
    '        strContNumeroProcesso = strNumeroCNJ
    '        Exit For
    '    End If
    'Next intCont
    On Error GoTo 0
    
    'COLOCAR UM TIMEOUT AQUI
    
    ' Procura pelo link
    'If DocHTML.getElementsByTagName("a")(2) Is Nothing Then
    '    PegaLinkProcesso = "Processo n�o encontrado"
    'End If
    
    'PegaLinkProcesso = DocHTML.getElementsByTagName("a")(intCont)
    
End Function

Sub ExpandirBotoesProcesso(ByRef IE As InternetExplorer, ByRef DocHTML As HTMLDocument, Optional ByVal intQuantidadeAExpandir As Integer, Optional ByVal bolExpandirAdvogados As Boolean)
''
'' Expande os "intQuantidadeAExpandir" primeiros bot�es de arquivos para download e informa��es de andamentos.
'' Se "intQuantidadeAExpandir" n�o tiver sido passada, abre tudo.
'' DEVO LIDAR COM O ERRO DE N�O SER PASSADA UMA P�GINA!!!!!!!
''

    Dim elCont As IHTMLElement
    Dim elLink As HTMLAnchorElement
    Dim intCont As Integer, intContAbertos As Integer
    
    'Expande advogados, se for o caso
    If bolExpandirAdvogados = True Then
        For Each elLink In DocHTML.getElementsByTagName("a")
            If elLink.innerText = "Mostrar/Ocultar" And InStr(1, elLink.href, "Adv") <> 0 Then
                elLink.Click
            ElseIf elLink.innerText = "Hist�rico de Ju�zes" Then
                Exit For
            End If
        Next elLink
    End If
    
    'Expande os bot�es
    If intQuantidadeAExpandir <> 0 Then intContAbertos = 0
    
    For intCont = 0 To DocHTML.getElementsByTagName("img").length - 1
        Set elCont = DocHTML.getElementsByTagName("img")(intCont)
        If (InStr(1, elCont.outerHTML, "src=""/projudi/imagens/observacao.png""") <> 0) Or (InStr(1, elCont.outerHTML, "src=""/projudi/imagens/arquivos.png""") <> 0) Then
            elCont.parentElement.Click
            If intQuantidadeAExpandir <> 0 Then
                intContAbertos = intContAbertos + 1
                If intContAbertos > intQuantidadeAExpandir - 1 Then Exit For
            End If
        End If
    Next intCont
    
End Sub

Function DescobrirPerfil(DocHTML As HTMLDocument) As String
''
'' Descobre o perfil do documento aberto e, conforme o caso, retorna "Parte" ou "Advogado"
''
    Dim frFrame As HTMLFrameElement
    
    Set frFrame = DocHTML.getElementsByName("mainFrame")(0)
    
    If InStr(1, frFrame.contentDocument.getElementById("Stm0p0i0eHR").href, "Parte") <> 0 Then '� parte
        DescobrirPerfil = "Parte"
    ElseIf InStr(1, frFrame.contentDocument.getElementById("Stm0p0i0eHR").href, "Advogado") <> 0 Then ' � Advogado
        DescobrirPerfil = "Advogado"
    Else '� outra coisa
        DescobrirPerfil = "Outro"
    End If
    
End Function
