VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Despesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Esta Classe possui muitas propriedades, representadas pelas seguintes variáveis:
Private strNumProc As String, strAdverso As String, strTipoDespesa As String, strComarca As String, strCodigoDAJE As String
Private strCodigoBarras As String, strEventosAtos As String
Private dtEmissao As Date, dtVencimento As Date
Private intQtdAtos As Integer
Private curValorDAJE As Currency

'Propriedade NumeroProcesso:
Public Property Get NumeroProcesso() As String
NumeroProcesso = strNumProc
End Property
Public Property Let NumeroProcesso(str As String)
strNumProc = str
End Property

'Propriedade Adverso:
Public Property Get Adverso() As String
Adverso = strAdverso
End Property
Public Property Let Adverso(str As String)
strAdverso = str
End Property

'Propriedade ValorDAJE:
Public Property Get ValorDAJE() As Currency
ValorDAJE = curValorDAJE
End Property
Public Property Let ValorDAJE(valor As Currency)
curValorDAJE = valor
End Property

'Propriedade TipoDespesa:
Public Property Get TipoDespesa() As String
TipoDespesa = strTipoDespesa
End Property
Public Property Let TipoDespesa(str As String)
strTipoDespesa = str
End Property

'Propriedade EventosAtos:
Public Property Get eventosAtos() As String
eventosAtos = strEventosAtos
End Property
Public Property Let eventosAtos(str As String)
strEventosAtos = str
End Property

'Propriedade QtdAtos:
Public Property Get QtdAtos() As Integer
QtdAtos = intQtdAtos
End Property
Public Property Let QtdAtos(val As Integer)
intQtdAtos = val
End Property

'Propriedade Comarca:
Public Property Get comarca() As String
comarca = strComarca
End Property
Public Property Let comarca(str As String)
strComarca = str
End Property

'Propriedade CodigoDAJE:
Public Property Get CodigoDAJE() As String
CodigoDAJE = strCodigoDAJE
End Property
Public Property Let CodigoDAJE(str As String)
strCodigoDAJE = str
End Property

'Propriedade CodigoBarras:
Public Property Get CodigoBarras() As String
CodigoBarras = strCodigoBarras
End Property
Public Property Let CodigoBarras(str As String)
strCodigoBarras = str
End Property

'Propriedade Emissao:
Public Property Get Emissao() As Date
Emissao = dtEmissao
End Property
Public Property Let Emissao(dt As Date)
dtEmissao = dt
End Property

'Propriedade Vencimento:
Public Property Get Vencimento() As Date
Vencimento = dtVencimento
End Property
Public Property Let Vencimento(dt As Date)
dtVencimento = dt
End Property

'Método ExportaLinhasEspaider:
Public Sub ExportaLinhasDespesasEspaider()
    Dim info As DajeInfo
    Dim jsonResp As String
    
    With info
        .dajeNumber = strCodigoDAJE
        .processoNumber = strNumProc
        .Adverso = strAdverso
        .valor = curValorDAJE
        .emissionDate = dtEmissao
        .dueDate = dtVencimento
        .barCode = strCodigoBarras
        .actType = strTipoDespesa
        .actsQuantity = intQtdAtos
        .actEventId = strEventosAtos
        .gerenciaEmbasa = "PPJCM"
    End With

    jsonResp = PostDajeToDb(info)
    
    If jsonResp <> "sucesso" Then
        MsgBox "Ocorreu o seguinte erro. Favor enviar um print ao administrador do sistema. Talvez o último DAJE não " & _
            "tenha sido salvo corretamente, favor conferir." & vbCrLf & "Resposta do backend: " & jsonResp
    End If
End Sub






