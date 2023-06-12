Attribute VB_Name = "modFuncoesApoioCustas"
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)
 
Private Sub AoCarregarRibbonDajes(Ribbon As IRibbonUI)
    ' Chama a função geral AoCarregarRibbon com os parâmetros corretos.
    AoCarregarRibbon cfConfiguracoes, Ribbon
End Sub

Sub LiberarEdicaoDajes(ByVal controle As IRibbonControl)
    ' Chama a função geral LiberarEdicao
    LiberarEdicao ThisWorkbook
End Sub

Sub RestringirEdicaoRibbonDajes(ByVal controle As IRibbonControl)
    ' Chama a função geral RestringirEdicaoRibbon
    RestringirEdicaoRibbon ThisWorkbook, controle
    
End Sub

Sub FechaConfigDajesVisivel(ByVal controle As IRibbonControl, Optional ByRef returnedVal)
    FechaConfigVisivel "btFechaConfigDajes", controle, returnedVal
End Sub

Function PegarValorCodBarras(codBarras As String) As Currency
    Dim resposta As Currency
    Dim valorBruto As String
    
    valorBruto = Mid(codBarras, 5, 7) & Mid(codBarras, 13, 4)
    resposta = CCur(valorBruto) / 100
    
    PegarValorCodBarras = resposta
    
End Function

Function PegarCodBarrasPdfDaje(nomeComCaminho As String, numeroDaje As String) As String
    Dim wd As Word.Application
    Dim wdDoc As Word.document
    Dim wdPar As Word.Paragraph
    Dim codBarras As String, resposta As String
    
    Application.StatusBar = "Sísifo - Lendo o código de barras com o Word"
    Set wd = New Word.Application
    Set wdDoc = AbrirPdf(wd, nomeComCaminho)
    For Each wdPar In wdDoc.Paragraphs
        If Len(wdPar.Range) >= 55 And Len(wdPar.Range) <= 56 Then
            codBarras = wdPar.Range
            Exit For
        End If
    Next wdPar
    
    codBarras = Replace(codBarras, " ", "")
    If Not IsNumeric(Right(codBarras, 1)) Then
        codBarras = Left(codBarras, Len(codBarras) - 1)
    End If
    
    numDajeParte1 = Left(numeroDaje, 8)
    numDajeParte2 = Right(numeroDaje, 5)
    posicaoparte1 = InStr(25, codBarras, numDajeParte1)
    posicaoparte2 = InStr(35, codBarras, numDajeParte2)
    If posicaoparte1 > 0 And posicaoparte2 > 0 Then
        resposta = codBarras
    Else
        resposta = "Conferir manualmente: " & codBarras
    End If
    wdDoc.Close wdDoNotSaveChanges
    wd.Quit
    Application.StatusBar = ""

    
    PegarCodBarrasPdfDaje = resposta
    
End Function

Function AbrirPdf(wd As Word.Application, nomeComCaminho As String) As Word.document
    Set AbrirPdf = wd.Documents.Open(nomeComCaminho, ConfirmConversions:=False, ReadOnly:=False, Revert:=False)
End Function

Function GetSlug(str As String) As String
    Dim response As String
    
    response = StrConv(str, vbLowerCase)
    response = Replace(response, " ", "-")
    response = Replace(response, "ª", "a")
    response = Replace(response, "º", "o")
    response = StripDiacrytics(response)
    response = Replace(response, "---", "-")
    GetSlug = response
End Function

Function StripDiacrytics(text As String)
    Dim A As String * 1, B As String * 1
    Dim response As String
    Dim i As Integer
    
    Const AccChars = "ŠšŸÀÁÂÃÄÅÇÈÉÊËÌÍÎÏĞÑÒÓÔÕÖÙÚÛÜİàáâãäåçèéêëìíîïğñòóôõöùúûüıÿ"
    Const RegChars = "SZszYAAAAAACEEEEIIIIDNOOOOOUUUUYaaaaaaceeeeiiiidnooooouuuuyy"
    response = text
    For i = 1 To Len(AccChars)
        A = Mid(AccChars, i, 1)
        B = Mid(RegChars, i, 1)
        response = Replace(response, A, B)
    Next i
    StripDiacrytics = response
End Function

Function GetStoredApiTokenEnsuringItExists() As String
    Dim token As String, response As String
    
    token = cfConfiguracoes.Cells().Find(what:="API Token", lookat:=xlWhole).Offset(0, 1).Formula
    If token = "" Then
        token = FetchNewToken
        SetStoredApiToken token
    End If
    GetStoredApiTokenEnsuringItExists = token
End Function

Sub SetStoredApiToken(token As String)
    cfConfiguracoes.Cells().Find(what:="API Token", lookat:=xlWhole).Offset(0, 1).Formula = token
End Sub
