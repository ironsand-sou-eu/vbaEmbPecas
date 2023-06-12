VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Faixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Esta Classe possiu muitas propriedades, representadas pelas seguintes variáveis:
Private strAbrangenciaFaixa As String, strTipoTarifa As String, strMensagemErro As String
Private intConsumoNaFaixa As Integer, intConsMaxAcumulado As Integer
Private curTarifa As Currency, curSubTotal As Currency

'Propriedade AbrangenciaFaixa:
Public Property Get AbrangenciaFaixa() As String
AbrangenciaFaixa = strAbrangenciaFaixa
End Property
Public Property Let AbrangenciaFaixa(str As String)
strAbrangenciaFaixa = str
End Property

'Propriedade TipoTarifa:
Public Property Get TipoTarifa() As String
TipoTarifa = strTipoTarifa
End Property
Public Property Let TipoTarifa(str As String)
strTipoTarifa = str
End Property

'Propriedade ConsumoNaFaixa:
Public Property Get ConsumoNaFaixa() As Integer
ConsumoNaFaixa = intConsumoNaFaixa
End Property
Public Property Let ConsumoNaFaixa(num As Integer)
intConsumoNaFaixa = num
End Property

'Propriedade ConsMaxAcumulado:
Public Property Get ConsMaxAcumulado() As Integer
ConsMaxAcumulado = intConsMaxAcumulado
End Property
Public Property Let ConsMaxAcumulado(num As Integer)
intConsMaxAcumulado = num
End Property

'Propriedade Tarifa:
Public Property Get Tarifa() As Currency
Tarifa = curTarifa
End Property
Public Property Let Tarifa(valor As Currency)
curTarifa = valor
End Property

'Propriedade SubTotal:
Public Property Get SubTotal() As Currency
SubTotal = curSubTotal
End Property
Public Property Let SubTotal(valor As Currency)
curSubTotal = valor
End Property

'Propriedade MensagemErro:
Public Property Get MensagemErro() As String
MensagemErro = strMensagemErro
End Property
Public Property Let MensagemErro(str As String)
strMensagemErro = str
End Property

'Método CalculaSubtotal:
Public Sub CalcularSubtotal()
''
'' Calcula o subtotal desta faixa e armazena em strSubTotal.
''
    
    Select Case strTipoTarifa
    Case "fixo"
        curSubTotal = curTarifa
        
    Case "por m3"
        curSubTotal = curTarifa * intConsumoNaFaixa
        
    Case Else
        strMensagemErro = "Tipo de tarifa inválido"
    
    End Select
    
End Sub
