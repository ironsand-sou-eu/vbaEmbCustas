VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Andamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Esta Classe possiu muitas propriedades, representadas pelas seguintes vari�veis:
Private strNumProc As String, strNomeAndamento As String, strObsAndamento As String, strIDPublicacao As String
Private dtDataAndamento As Date
Private curValorCausa As Currency
Private listaProvidencias As Collection

Private Sub Class_Initialize()
    
    'Seta as cole��es
    Set listaProvidencias = New Collection
    
End Sub

'Propriedade Providencias:
Public Property Get Providencias() As Collection
Set Providencias = listaProvidencias
End Property

'Propriedade NumeroProcesso:
Public Property Get NumeroProcesso() As String
NumeroProcesso = strNumProc
End Property
Public Property Let NumeroProcesso(str As String)
strNumProc = str
End Property

'Propriedade ValorCausa:
Public Property Get ValorCausa() As Currency
ValorCausa = curValorCausa
End Property
Public Property Let ValorCausa(valor As Currency)
curValorCausa = valor
End Property

'Propriedade NomeAndamento:
Public Property Get NomeAndamento() As String
NomeAndamento = strNomeAndamento
End Property
Public Property Let NomeAndamento(str As String)
strNomeAndamento = str
End Property

Public Function CompactaObsAndamento(strObsAndamento) As String
''
'' Remove quebras de espa�o duplicadas em ObsAndamento.
''
    Dim strObsAndamentoFinal As String
    
    If InStr(1, strObsAndamento, vbCrLf & vbCrLf) Then
        strObsAndamentoFinal = Replace(strObsAndamento, vbCrLf & vbCrLf, vbCrLf)
    End If
    
    CompactaObsAndamento = strObsAndamentoFinal
    
End Function

'Propriedade DataAndamento:
Public Property Get DataAndamento() As Date
DataAndamento = dtDataAndamento
End Property
Public Property Let DataAndamento(dt As Date)
dtDataAndamento = dt
End Property

'Propriedade ObsAndamento:
Public Property Get ObsAndamento() As String
ObsAndamento = strObsAndamento
End Property
Public Property Let ObsAndamento(str As String)
'strObsAndamento = Replace(str, ChrW(&H2028), "") ' Esse funcionava!
strObsAndamento = Replace(str, ChrW(2), "") ' Remove caracteres especiais Unicode que bugam o XML do Excel.
strObsAndamento = Replace(strObsAndamento, ChrW(11), "") ' Remove  caracteres especiais Unicode que bugam o XML do Excel.
strObsAndamento = Replace(strObsAndamento, ChrW(30), "-") ' Remove caracteres especiais Unicode que bugam o XML do Excel.
strObsAndamento = CompactaObsAndamento(strObsAndamento)
End Property

'Propriedade IDPublicacao:
Public Property Get IDPublicacao() As String
IDPublicacao = strIDPublicacao
End Property
Public Property Let IDPublicacao(str As String)
strIDPublicacao = str
End Property

'M�todo ExportaLinhasEspaider:
Public Sub ExportaLinhasAndamentosEspaider(arq As Workbook, plan As Worksheet, Optional strOrigem As String, Optional CelulaAtiva As Range)
''
'' Exporta as linhas. A planilha passada como par�metro deve estar limpa!
'' Se houver mais de uma parte adversa, cria tamb�m uma outra planilha na mesma pasta de trabalho, preparada
'' para exportar para ProcessoParticipantes.
''
    Dim intCont As Integer
    Dim lnUltimaLinha As Long

    
    With plan
        .Cells(1, 1) = strNumProc
        .Cells(1, 2) = dtDataAndamento
        .Cells(1, 3) = strNomeAndamento
        .Cells(1, 4) = strObsAndamento
        .Cells(1, 18) = strIDPublicacao
    
        intCont = 1
        If listaProvidencias.Count > 0 Then
            For intCont = 1 To listaProvidencias.Count
                .Cells(intCont, 5) = strNumProc
                .Cells(intCont, 6) = listaProvidencias(intCont).Nome
                .Cells(intCont, 7) = listaProvidencias(intCont).Nucleo
                .Cells(intCont, 8) = listaProvidencias(intCont).Responsavel
                .Cells(intCont, 9) = listaProvidencias(intCont).DataFinal
                .Cells(intCont, 10) = listaProvidencias(intCont).PrazoDias
                .Cells(intCont, 11) = listaProvidencias(intCont).Situacao
                .Cells(intCont, 12) = IIf(listaProvidencias(intCont).Alertar, "Sim", "N�o")
                .Cells(intCont, 13) = listaProvidencias(intCont).PeriodicidadeAlerta
                .Cells(intCont, 14) = listaProvidencias(intCont).Observacao
                .Cells(intCont, 15) = IIf(listaProvidencias(intCont).GerarAndamento, "Sim", "N�o")
                .Cells(intCont, 16) = listaProvidencias(intCont).DataAndamento
                .Cells(intCont, 17) = listaProvidencias(intCont).Andamento
            Next intCont
        End If
        
    End With
    
    ' Pede ao usu�rio que confirme a inclus�o. � importante principalmente para ele conferir os dados e para perceber
    ' eventual erro. O importante � confirmar o processo, andamento e data do andamento.
    plan.UsedRange.Columns.AutoFit
    plan.Activate
    plan.Range("J1:L1").Select
    
    If MsgBox("Confira os dados do processo. Deseja continuar com a inclus�o?", vbQuestion + vbYesNo, "S�sifo - Confirmar inclus�o") = vbNo Then
        ' Se n�o confirmar, avisa e n�o salva esse processo.
        'arq.Close False
        Do
        Loop Until MsgBox(DeterminarTratamento & ", a inclus�o foi cancelada. Descartamos os dados." & vbCrLf & _
        "Processo: " & strNumProc & vbCrLf & _
        "Andamento: " & strNomeAndamento & "em " & dtDataAndamento & vbCrLf & _
        "Clique em 'Cancelar' e insira novamente o processo.", vbCritical + vbOKCancel, "S�sifo - Inclus�o cancelada pelo usu�rio") = vbCancel
        Exit Sub
    End If
    
    ' Confirmados os dados, confere se o processo j� est� na planilha. OBS: Se j� estiver na planilha o ID (XML) ou o n�mero (Projudi), n�o deixar� cadastrar, mesmo que sejam andamentos diferentes.
    If (strOrigem = "Projudi" And Not ThisWorkbook.Sheets("Sisifo CadAndamento v1").Cells().Find(strNumProc) Is Nothing) Or _
        strOrigem = "XML" And Not ThisWorkbook.Sheets("Sisifo CadAndamento v1").Cells().Find(strIDPublicacao) Is Nothing Then
        'arq.Close False
        Do
        Loop Until MsgBox(DeterminarTratamento & ", o processo j� existe na planilha! A inclus�o foi cancelada. Descartamos os dados." & vbCrLf & _
        "Processo: " & strNumProc & vbCrLf & _
        "Andamento: " & strNomeAndamento & "em " & dtDataAndamento & vbCrLf & _
        "Clique em 'Cancelar' e insira o pr�ximo processo.", vbCritical + vbOKCancel, "S�sifo - Processo repetido") = vbCancel
        Exit Sub
    End If
    
    ' N�o estando na planilha, pimba na gorduchinha!
    plan.UsedRange.Copy
    lnUltimaLinha = ThisWorkbook.Sheets("Sisifo CadAndamento v1").UsedRange.Rows(ThisWorkbook.Sheets("Sisifo CadAndamento v1").UsedRange.Rows.Count).Row
    With ThisWorkbook.Sheets("Sisifo CadAndamento v1")
        .Cells(lnUltimaLinha + 1, 1).PasteSpecial xlPasteAll
        .Rows(lnUltimaLinha + 1).WrapText = False
    End With
    
    'Ap�s salvar, confirma a inclus�o e fecha a pasta de trabalho de rascunho.
    'Application.DisplayAlerts = False
    ThisWorkbook.Save
    'ThisWorkbook.SaveAs Filename:=ThisWorkbook.FullName, FileFormat:=xlOpenXMLAddIn
    'Application.DisplayAlerts = True
    MsgBox DeterminarTratamento & ", o Andamento foi registrado com sucesso!" & vbCrLf & _
    "Processo: " & strNumProc & vbCrLf & _
    "Andamento: " & strNomeAndamento & " em " & dtDataAndamento & vbCrLf & _
    "Lembre-se de inserir no Espaider no final do dia.", vbInformation + vbOKOnly, "S�sifo - Processo salvo para inclus�o"
    'arq.Close False
    
    If strOrigem = "Projudi" Then CelulaAtiva.Interior.Color = 14857357

End Sub


