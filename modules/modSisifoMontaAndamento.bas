Attribute VB_Name = "modSisifoMontaAndamento"
Function PegarHistoricoDeAdvogados(ByRef IE As InternetExplorer, ByRef DocHTML As HTMLDocument) As String
''
'' Pega todos os nomes de advogados do Hist�rico de Advogados de todas as partes e retorna em uma string, separados por v�rgulas, com uma v�rgula no final
'' DEVO LIDAR COM O ERRO DE N�O SER PASSADA UMA P�GINA!!!!!!!
''
    Dim elCont As Variant, elContLinha As Variant
    Dim strContAdv As String, strFeitos As String, strAdvs As String
    Dim intCont As Integer
    Dim snInicioTimer As Single
    
    ' Para cada link de Hist�rico de advogados...
    For Each elCont In DocHTML.getElementsByTagName("a")
        If elCont.innerText = "Hist�rico de Advogados" Then
            elCont.Click 'elcont.parentelement.parentelement.parentelement.parentelement.parentelement.parentelement.parentelement.parentelement.parentelement.parentelement.parentelement.parentelement.parentelement.id
            
            ' Procurar a frame que cont�m o Hist�rico de Advogados exibido
            On Error Resume Next
AguardarHistoricoAdv:
            snInicioTimer = Timer
            Do
                For intCont = 0 To DocHTML.Frames.length - 1 Step 1
                    If InStr(1, DocHTML.Frames(intCont).document.url, "HistoricoAdvogado") <> 0 And _
                        InStr(1, strFeitos, DocHTML.Frames(intCont).document.url) = 0 Then Exit Do ' Tem que ter HistoricoAdvogado e n�o ter a URL da parte na lista das feitas
                    
                    If Err.Number <> 0 Then
                        Err.Clear
                        GoTo AguardarHistoricoAdv
                    End If
                Next intCont
            Loop Until Timer >= snInicioTimer + 10
            
            If Timer >= snInicioTimer + 10 Then 'Se saiu por causa do timer, pergunta se quer continuar esperando
                If MsgBox(DeterminarTratamento & ", a p�gina com o hist�rico de advogados parece estar demorando para carregar. Caso ela j� esteja carregada e " & _
                    "eu n�o tenha percebido, ou para pular essa lista de advogados, clique em ""Cancelar"". Para aguardar mais 10 segundos, clique em " & _
                    """Tentar novamente"". Isso n�o afeta diretamente a contagem de atos -- apenas indiretamente, pois pode fazer com que eu n�o saiba " & _
                    "que algumas pessoas s�o advogados do processo, mas o senhor pode verificar para mim!", vbQuestion + vbRetryCancel, _
                    "S�sifo - Lista de advogados demorando de carregar") = vbRetry Then
                    
                    GoTo AguardarHistoricoAdv
                Else
                    GoTo HistoricoVazio
                End If
                    
            End If
            
            On Error GoTo 0
            
AdvVazio:
            ' Iterar as linhas da frame para pegar os nomes dos advogados
            On Error Resume Next
            strContAdv = ""
            For Each elContLinha In DocHTML.Frames(intCont).document.getElementsByTagName("tr")
                If elContLinha.ClassName = "primeiraLinha" Then
                    If elContLinha.parentElement.Rows(elContLinha.RowIndex + 1).ClassName = "ultimaLinha" Then GoTo HistoricoVazio
                ElseIf elContLinha.ClassName = "tBranca" Or elContLinha.ClassName = "tCinza" Then
                    If Err.Number = 70 Then GoTo AdvVazio
                    strContAdv = strContAdv & elContLinha.Children(0).innerText & ","
                End If
            Next elContLinha
            
            On Error GoTo 0
            
            If strContAdv = "" Then
                GoTo AdvVazio
            Else
                strAdvs = strAdvs & strContAdv
                strFeitos = strFeitos & DocHTML.Frames(intCont).document.url & ","
            End If
            DocHTML.parentWindow.execScript "hidePopWin(false);"
            
        ElseIf elCont.innerText = "Hist�rico de Ju�zes" Then
            Exit For
        End If
HistoricoVazio:
    Next elCont
    
    PegarHistoricoDeAdvogados = strAdvs
    
End Function

Function PegarQtdPartes(IE As InternetExplorer, DocHTML As HTMLDocument) As Integer
''
'' Retorna a quantidade de imagens de parte existentes no processo.
''
    Dim intCont As Integer
    Dim varCont As Variant
    
    intCont = 0
    
    For Each varCont In DocHTML.getElementsByTagName("img")
        If InStr(1, varCont.src, "/projudi/imagens/dadosParte.png") <> 0 Then intCont = intCont + 1
    Next varCont
    
    PegarQtdPartes = intCont
    
End Function
