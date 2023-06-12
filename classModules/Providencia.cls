VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Providencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Esta Classe possiu muitas propriedades, representadas pelas seguintes variáveis:

Private dtDataAlerta As Date, dtDataFinal As Date
Private intPrazo As Integer, intSituacao As Integer, intAntecAlerta As Integer, intPeriodicidade As Integer, intDataAndamento As Integer
Private strNomeProv As String, strNucleo As String, strResponsavel As String, strObs As String, strAndamento As String
Private bolAlertar As Boolean, bolGerarAndamento As Boolean

Private Sub Class_Initialize()
    ' Carrega os valores padrão
   
End Sub

'Propriedade Nome:
Public Property Get Nome() As String
Nome = strNomeProv
End Property
Public Property Let Nome(str As String)
strNomeProv = str
End Property

'Propriedade Andamento:
Public Property Get Andamento() As String
Andamento = strAndamento
End Property
Public Property Let Andamento(str As String)
strAndamento = str
End Property

'Propriedade DataAndamento:
Public Property Get DataAndamento() As Integer
DataAndamento = intDataAndamento
End Property
Public Property Let DataAndamento(num As Integer)
intDataAndamento = num
End Property

'Propriedade GerarAndamento:
Public Property Get GerarAndamento() As Boolean
GerarAndamento = bolGerarAndamento
End Property
Public Property Let GerarAndamento(bol As Boolean)
bolGerarAndamento = bol
End Property

'Propriedade Núcleo:
Public Property Get Nucleo() As String
Nucleo = strNucleo
End Property
Public Property Let Nucleo(str As String)
strNucleo = str
End Property

'Propriedade Responsável:
Public Property Get Responsavel() As String
Responsavel = strResponsavel
End Property
Public Property Let Responsavel(str As String)
strResponsavel = str
End Property

'Propriedade Observação:
Public Property Get Observacao() As String
Observacao = strObs
End Property
Public Property Let Observacao(str As String)
strObs = str
End Property

'Propriedade Alertar:
Public Property Get Alertar() As Boolean
Alertar = bolAlertar
End Property
Public Property Let Alertar(bol As Boolean)
bolAlertar = bol
End Property

'Propriedade DataFinal:
Public Property Get DataFinal() As Date
DataFinal = dtDataFinal
End Property
Public Property Let DataFinal(dt As Date)
dtDataFinal = dt
End Property

'Propriedade DataAlerta:
Public Property Get DataAlerta() As Date
DataAlerta = dtDataAlerta
End Property
Public Property Let DataAlerta(dt As Date)
dtDataAlerta = dt
End Property

'Propriedade PrazoDias:
Public Property Get PrazoDias() As Integer
PrazoDias = intPrazo
End Property
Public Property Let PrazoDias(num As Integer)
intPrazo = num
End Property

'Propriedade Situacao:
Public Property Get Situacao() As Integer
Situacao = intSituacao
End Property
Public Property Let Situacao(num As Integer)
intSituacao = num
End Property

'Propriedade DiasAntecedenciaAlerta:
Public Property Get DiasAntecedenciaAlerta() As Integer
DiasAntecedenciaAlerta = intAntecAlerta
End Property
Public Property Let DiasAntecedenciaAlerta(num As Integer)
intAntecAlerta = num
End Property

'Propriedade PeriodicidadeAlerta:
Public Property Get PeriodicidadeAlerta() As Integer
PeriodicidadeAlerta = intPeriodicidade
End Property
Public Property Let PeriodicidadeAlerta(num As Integer)
intPeriodicidade = num
End Property
