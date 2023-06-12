VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Categoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Esta Classe possiu muitas propriedades, representadas pelas seguintes variáveis:
Private intQtdEconomias As Integer
Private btQtdFaixas As Byte
Private strCategoria As String, strMensagemErro As String
Private curAguaPorEconomia As Currency, curSubtotalAgua As Currency
Private listaFaixas As Collection

Private Sub Class_Initialize()
    'Seta as coleções
    Set listaFaixas = New Collection
    
End Sub

'Propriedade Faixas:
Public Property Get Faixas() As Collection
Set Faixas = listaFaixas
End Property

'Propriedade Categoria:
Public Property Get Categoria() As String
Categoria = strCategoria
End Property
Public Property Let Categoria(str As String)
strCategoria = str
End Property

'Propriedade QtdEconomias:
Public Property Get QtdEconomias() As Integer
QtdEconomias = intQtdEconomias
End Property
Public Property Let QtdEconomias(num As Integer)
intQtdEconomias = num
End Property

'Propriedade AguaPorEconomia:
Public Property Get AguaPorEconomia() As Currency
AguaPorEconomia = curAguaPorEconomia
End Property
Public Property Let AguaPorEconomia(valor As Currency)
curAguaPorEconomia = valor
End Property

'Propriedade SubtotalAgua:
Public Property Get SubtotalAgua() As Currency
SubtotalAgua = curSubtotalAgua
End Property
Public Property Let SubtotalAgua(valor As Currency)
curSubtotalAgua = valor
End Property

'Propriedade MensagemErro:
Public Property Get MensagemErro() As String
MensagemErro = strMensagemErro
End Property
Public Property Let MensagemErro(str As String)
strMensagemErro = str
End Property

'Método BuscarEstruturaTarifaria:
Public Sub BuscarEstruturaTarifaria(strResolucao As String, intConsumoAguaPorEconomia As Integer)
''
'' Estando preenchidos strCategoria, e dados os parâmetros strResolucao e intConsumoAguaPorEconomia,
''   busca a estrutura tarifária em cfTarifas e preenche as faixas e valores.
''
    Dim plan As Worksheet
    Dim rngCont As Excel.Range
    Dim fxCont As Faixa
    Dim intQtRes As Integer, intCont As Integer
    
    If Trim(strCategoria) = "" Then
        strMensagemErro = "Categoria tarifária em branco"
        Exit Sub
    End If
    
    ' Vai para o início da planilha e procura a tarifa adequada
    Set plan = ThisWorkbook.Sheets("cfTarifas")
    
    intQtRes = plan.Cells().Find(what:="Quantidade de resoluções:", lookat:=xlWhole).Offset(0, 1).Formula
    
    Set rngCont = plan.Range("A1")
    Set rngCont = plan.Cells().Find(what:=strCategoria, After:=rngCont, lookat:=xlWhole)
    
    If rngCont Is Nothing Then
        strMensagemErro = "Categoria tarifária não encontrada"
        Exit Sub
    End If
    
    For intCont = 1 To intQtRes Step 1
        If plan.Cells(rngCont.Row - 1, 1).Formula = strResolucao Then Exit For 'Se achou a categoria certa da resolução certa, prossegue
        Set rngCont = plan.Cells().FindNext(After:=rngCont) 'Se não, procura a próxima
    Next intCont
    
    If rngCont Is Nothing Then
        strMensagemErro = "Categoria tarifária " & strCategoria & " não encontrada sob a Resolução " & strResolucao
        Exit Sub
    End If
    
    'Encontrada a classificação. rngCont aponta para a célula da categoria.
    btQtdFaixas = rngCont.Offset(0, 1).Formula
    
    'Monta as faixas
    For intCont = 1 To btQtdFaixas Step 1
        Set fxCont = New Faixa
        fxCont.AbrangenciaFaixa = rngCont.Offset(intCont, 0).Value
        fxCont.ConsMaxAcumulado = rngCont.Offset(intCont, 1).Value
        fxCont.ConsumoNaFaixa = rngCont.Offset(intCont, 2).Value 'Pega o valor máximo da faixa e testa pra ver se é aceitável
        
        If intConsumoAguaPorEconomia >= fxCont.ConsMaxAcumulado Then
            ' Consumo na faixa é o total da faixa, portanto o valor já está certo.
        ElseIf intConsumoAguaPorEconomia > fxCont.ConsMaxAcumulado - fxCont.ConsumoNaFaixa Then
            ' Consumo na faixa é parte do total da faixa. Calcular a diferença.
            fxCont.ConsumoNaFaixa = intConsumoAguaPorEconomia - (fxCont.ConsMaxAcumulado - fxCont.ConsumoNaFaixa)
        Else
            ' Consumo não chegou na faixa.
            fxCont.ConsumoNaFaixa = 0
        End If
        
        fxCont.TipoTarifa = rngCont.Offset(intCont, 3).Value
        fxCont.Tarifa = rngCont.Offset(intCont, 4).Value
        
        listaFaixas.Add fxCont
    
    Next intCont
        
End Sub

'Método CalcularSubtotais:
Public Sub CalcularSubtotais()
''
'' Estando preenchido intQtdEconomias, calcula o subtotal desta categoria (calculando as faixas) e armazena em curAguaPorEconomia e curSubtotalAgua.
''
    
    Dim btCont As Byte
    Dim curCont As Currency
    
    If Trim(intQtdEconomias) = "" Then
        strMensagemErro = "Quantidade de economias não informada"
        Exit Sub
    End If
    
    curCont = 0
    
    For btCont = 1 To listaFaixas.Count Step 1 ' Calcula os subtotais de cada faixa e os soma
        listaFaixas(btCont).CalcularSubtotal
        curCont = curCont + listaFaixas(btCont).SubTotal
    Next btCont
    
    curAguaPorEconomia = curCont
    curSubtotalAgua = curAguaPorEconomia * intQtdEconomias
    
End Sub

