VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Fatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Esta Classe possiu muitas propriedades, representadas pelas seguintes variáveis:
Private intConsumoTotal As Integer, intConsumoPorEconomia As Integer, intQtdTotalEconomias As Integer
Private btPorcentEsgoto As Byte
Private strMesReferencia As String, strMensagemErro As String
Private curTotalAgua As Currency, curTotalEsgoto As Currency, curTotalAguaEsgoto As Currency
Private listaCategorias As Collection

Private Sub Class_Initialize()
    'Seta as coleções
    Set listaCategorias = New Collection
    
End Sub

'Propriedade Categorias:
Public Property Get Categorias() As Collection
Set Categorias = listaCategorias
End Property

'Propriedade MesReferencia:
Public Property Get MesReferencia() As String
MesReferencia = strMesReferencia
End Property
Public Property Let MesReferencia(str As String)
strMesReferencia = str
End Property

'Propriedade PorcentEsgoto:
Public Property Get PorcentEsgoto() As Byte
PorcentEsgoto = btPorcentEsgoto
End Property
Public Property Let PorcentEsgoto(num As Byte)
btPorcentEsgoto = num
End Property

'Propriedade ConsumoTotal:
Public Property Get ConsumoTotal() As Integer
ConsumoTotal = intConsumoTotal
End Property
Public Property Let ConsumoTotal(num As Integer)
intConsumoTotal = num
End Property

'Propriedade ConsumoPorEconomia:
Public Property Get ConsumoPorEconomia() As Integer
ConsumoPorEconomia = intConsumoPorEconomia
End Property
Public Property Let ConsumoPorEconomia(num As Integer)
intConsumoPorEconomia = num
End Property

'Propriedade QtdTotalEconomias:
Public Property Get QtdTotalEconomias() As Integer
QtdTotalEconomias = intQtdTotalEconomias
End Property
Public Property Let QtdTotalEconomias(num As Integer)
intQtdTotalEconomias = num
End Property

'Propriedade TotalAgua:
Public Property Get TotalAgua() As Currency
TotalAgua = curTotalAgua
End Property
Public Property Let TotalAgua(valor As Currency)
curTotalAgua = valor
End Property

'Propriedade TotalEsgoto:
Public Property Get TotalEsgoto() As Currency
TotalEsgoto = curTotalEsgoto
End Property
Public Property Let TotalEsgoto(valor As Currency)
curTotalEsgoto = valor
End Property

'Propriedade TotalAguaEsgoto:
Public Property Get TotalAguaEsgoto() As Currency
TotalAguaEsgoto = curTotalAguaEsgoto
End Property
Public Property Let TotalAguaEsgoto(valor As Currency)
curTotalAguaEsgoto = valor
End Property

'Propriedade MensagemErro:
Public Property Get MensagemErro() As String
MensagemErro = strMensagemErro
End Property
Public Property Let MensagemErro(str As String)
strMensagemErro = str
End Property

'Método BuscarEstruturaTarifaria:
Public Sub BuscarEstruturaTarifaria(strMesFat As String, intConsumoFat As Integer, btPercentualEsgoto As Byte, strClassificacao1 As String, intQtdEconomias1 As Integer, Optional strClassificacao2 As String, Optional intQtdEconomias2 As Integer, Optional strClassificacao3 As String, Optional intQtdEconomias3 As Integer, Optional strClassificacao4 As String, Optional intQtdEconomias4 As Integer)
''
'' Passados como parâmetros strMesReferencia, intConsumoTotal, btPorcentEsgoto e as classificações e quantidade de economias
''   de cada tipo (pelo menos 1 e até 4), monta as categorias e busca a estrutura tarifária em cfTarifas.
''
    Dim catCont As Categoria
    Dim strResolucao As String
    Dim bolReduzirAno As Boolean
    Dim intCont As Integer
    
    If Trim(strClassificacao1) = "" Or intQtdEconomias1 = 0 Then
        strMensagemErro = "Classificação tarifária 1 faltando ou com 0 economias"
        Exit Sub
    End If
    
    strMesReferencia = strMesFat
    intConsumoTotal = intConsumoFat
    btPorcentEsgoto = btPercentualEsgoto
    
    'Descobrir resolução pelo mês da fatura (não é o mais seguro, eu sei)
    bolReduzirAno = False
    
    Select Case Left(strMesReferencia, 2)
    Case "01", "02", "03", "04", "05", "06"
        bolReduzirAno = True
    
    Case "07", "08"
        strMensagemErro = "Atualmente, não estamos prontos para calcular faturas de julho e agosto."
        Exit Sub
    
    Case "09", "10", "11", "12"
        bolReduzirAno = False
    
    Case Else
        strMensagemErro = "Período de referência da fatura inválido. A referência da fatura deve ser passada no formato ""05/2019""."
        Exit Sub
    
    End Select
    
    Select Case Right(strMesReferencia, 4)
    Case "2013"
        If bolReduzirAno = True Then
            strMensagemErro = "Resolução anterior à 05/2013."
            Exit Sub
        Else
            strResolucao = "05/2013"
        End If
    
    Case "2014"
        strResolucao = IIf(bolReduzirAno, "05/2013", "02/2014")
    
    Case "2015"
        strResolucao = IIf(bolReduzirAno, "02/2014", "01/2015")
    
    Case "2016"
        strResolucao = IIf(bolReduzirAno, "01/2015", "02/2016")
    
    Case "2017"
        strResolucao = IIf(bolReduzirAno, "02/2016", "01/2017")
    
    Case "2018"
        strResolucao = IIf(bolReduzirAno, "01/2017", "04/2018")
    
    Case "2019"
        strResolucao = IIf(bolReduzirAno, "04/2018", "01/2019")
    
    Case "2020"
        strResolucao = IIf(bolReduzirAno, "01/2019", "")
    
    Case Else
        strMensagemErro = "Resolução não encontrada ou período de referência da fatura inválido. A referência da fatura deve ser passada no formato ""05/2019""."
        Exit Sub
    
    End Select
    
    'Descobrir quantidade total de economias
    intCont = intQtdEconomias1
    If intQtdEconomias2 <> 0 Then intCont = intCont + intQtdEconomias2
    If intQtdEconomias3 <> 0 Then intCont = intCont + intQtdEconomias3
    If intQtdEconomias4 <> 0 Then intCont = intCont + intQtdEconomias4
    
    intQtdTotalEconomias = intCont
    intConsumoPorEconomia = intConsumoTotal / intQtdTotalEconomias
    
    'Alimentar as categorias
    Set catCont = New Categoria
    catCont.Categoria = strClassificacao1
    catCont.QtdEconomias = intQtdEconomias1
    catCont.BuscarEstruturaTarifaria strResolucao, intConsumoPorEconomia
    listaCategorias.Add catCont
    
    If strClassificacao2 <> "" And intQtdEconomias2 <> 0 Then
        Set catCont = New Categoria
        catCont.Categoria = strClassificacao2
        catCont.QtdEconomias = intQtdEconomias2
        catCont.BuscarEstruturaTarifaria strResolucao, intConsumoPorEconomia
        listaCategorias.Add catCont
    End If
    
    If strClassificacao3 <> "" And intQtdEconomias3 <> 0 Then
        Set catCont = New Categoria
        catCont.Categoria = strClassificacao3
        catCont.QtdEconomias = intQtdEconomias3
        catCont.BuscarEstruturaTarifaria strResolucao, intConsumoPorEconomia
        listaCategorias.Add catCont
    End If
    
    If strClassificacao4 <> "" And intQtdEconomias4 <> 0 Then
        Set catCont = New Categoria
        catCont.Categoria = strClassificacao4
        catCont.QtdEconomias = intQtdEconomias4
        catCont.BuscarEstruturaTarifaria strResolucao, intConsumoPorEconomia
        listaCategorias.Add catCont
    End If

End Sub

'Método ExportaLinhasEspaider:
Public Sub CalcularTotal()
''
'' Estando preenchido btPorcentEsgoto, calcula o total da fatura (calculando as categorias e faixas) e armazena em
''   curTotalAgua, curTotalEsgoto e curTotalAguaEsgoto.
''
    
    Dim btCont As Byte
    Dim curCont As Currency
    
    curCont = 0
    
    For btCont = 1 To listaCategorias.Count Step 1  ' Calcula os subtotais de cada faixa e os soma
        listaCategorias(btCont).CalcularSubtotais
        curCont = curCont + listaCategorias(btCont).SubtotalAgua
    Next btCont
    
    curTotalAgua = curCont
    curTotalEsgoto = curTotalAgua * btPorcentEsgoto / 100
    curTotalAguaEsgoto = curTotalAgua + curTotalEsgoto
    
End Sub
