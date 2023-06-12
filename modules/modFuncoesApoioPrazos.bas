Attribute VB_Name = "modFuncoesApoioPrazos"
Option Explicit

Public bolSsfPrazosBotaoPdfPressionado As Boolean

Sub FechaConfigPrazosVisivel(ByVal controle As IRibbonControl, Optional ByRef returnedVal)
    
    FechaConfigVisivel "btFechaConfigPrazos", controle, returnedVal
        
End Sub

Private Sub AoCarregarRibbonPrazos(Ribbon As IRibbonUI)
    ' Chama a função geral AoCarregarRibbon com os parâmetros corretos.
    Dim plPlan As Worksheet
    Set plPlan = ThisWorkbook.Sheets("cfConfigurações")
    
    AoCarregarRibbon plPlan, Ribbon
    
End Sub

Sub LiberarEdicaoPrazos(ByVal controle As IRibbonControl)
' Chama a função geral LiberarEdicao
    LiberarEdicao ThisWorkbook
    
End Sub

Sub RestringirEdicaoRibbonPrazos(ByVal controle As IRibbonControl)
' Chama a função geral RestringirEdicaoRibbon
    RestringirEdicaoRibbon ThisWorkbook, controle
    
End Sub

Public Sub ssfPrazosTglGerarPDF_getPressed(control As IRibbonControl, ByRef returnedVal)
' Botão de gerar PDF
    returnedVal = bolSsfPrazosBotaoPdfPressionado
End Sub

Public Sub ssfPrazosTglGerarPDF_onAction(control As IRibbonControl, ByRef cancelDefault)
' Botão de gerar PDF
    bolSsfPrazosBotaoPdfPressionado = Not bolSsfPrazosBotaoPdfPressionado
End Sub

Function BuscarCaminhoPrograma() As String
    Dim rngCaminho As Range
    
    Set rngCaminho = ThisWorkbook.Sheets("cfConfigurações").Cells.Find(what:="Caminho da pasta do Sisifo", lookat:=xlWhole)
    
    If rngCaminho Is Nothing Then
        BuscarCaminhoPrograma = ""
    Else
        BuscarCaminhoPrograma = rngCaminho.Offset(1, 0).Formula
    End If
    
End Function

Function BuscaJuizo(strJuizoEspaider As String) As String
    Dim rngContJuizo As Range
    
    Set rngContJuizo = ThisWorkbook.Sheets("cfJuizos").Cells.Find(what:=strJuizoEspaider, lookat:=xlWhole)
    
    If rngContJuizo Is Nothing Then
        BuscaJuizo = ""
    Else
        BuscaJuizo = rngContJuizo.Offset(0, 1).Formula
    End If
    
End Function

Sub AjustarRelatorioProvidencias(control As IRibbonControl)

    Dim plan As Worksheet
    Dim lnUltimaLinha As Long

    Set plan = ActiveSheet
    
    With plan
        If .Shapes.Count <> 0 Then .Shapes(1).Delete
        If .Cells(3, 1).Formula = "" Then .Rows(3).Delete
        If .Cells(2, 1).Formula = "" Then .Rows(2).Delete
        If .Cells(1, 1).Formula = "" Then .Rows(1).Delete
        
        lnUltimaLinha = .UsedRange.Rows.Count
        
        If .Cells(lnUltimaLinha, 1).Formula = "" Then .Rows(lnUltimaLinha).Delete
        
        .Rows("2:" & lnUltimaLinha).WrapText = False
        .Rows(1).RowHeight = 30
        
        If .Cells(1, 2).Formula = "Adverso" Then .Columns(2).ColumnWidth = 40
        If LCase(.Cells(1, 3).Formula) = "matrícula principal" Then .Columns(3).ColumnWidth = 14
        If .Cells(1, 7).Formula = "Observações Prov." Then .Columns(7).ColumnWidth = 40
        If .Cells(1, 8).Formula = "Andamento" Then .Columns(8).ColumnWidth = 35
        If .Cells(1, 10).Formula = "Obs. do Andamento" Then .Columns(10).ColumnWidth = 30
        If .Cells(1, 11).Formula = "Juízo" Then .Columns(11).ColumnWidth = 40
        If .AutoFilter Is Nothing Then .UsedRange.AutoFilter
        
    End With
    
End Sub

Sub RecortarHistoricoCCR(control As IRibbonControl)
''
'' Em um histórico de matrícula do Espaider, apaga Categorias Tarifárias, Dados da Ligação de Esgoto, HPAG posterior a 05/2015,
''    HCON exceto 2014 a 2016, Consulta Notificações de Débito, Serviços posteriores a 05/2015, Informações de parcelamento,
''    COOB posterior a 05/2015.
''
    
    'Dim plan As Worksheet
    Dim rngInicio As Range, rngFim As Range
    Dim strUltimoCorte As String ', strUltimaReligacao As String
    
    'Dim rngTeste As Range
    
    ' Apaga os seguintes cabeçalhos:
    ApagarSubtitulo ActiveSheet, "CATEGORIAS TARIFÁRIAS"
    ApagarSubtitulo ActiveSheet, "DADOS DA LIGAÇÃO DE ESGOTO"
    ApagarSubtitulo ActiveSheet, "CONSULTA NOTIFICAÇÕES DE  DÉBITO (CNOT)"
    ApagarSubtitulo ActiveSheet, "INFORMAÇÕES DO PARCELAMENTO - HISTÓRICO (IPAR)"
    
    ' Se o último corte foi antes ou depois da CCR, apaga os outros cabeçalhos.
    strUltimoCorte = ActiveSheet.Cells().Find(what:="Último corte:", lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False).Offset(0, 1).Text
    strUltimaReligacao = ActiveSheet.Cells().Find(what:="Última religação:", lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False).Offset(0, 1).Text
    If strUltimoCorte = "31/12/9999" Or _
        (CDate(strUltimaReligacao) >= CDate(strUltimoCorte) And CDate(strUltimaReligacao) < CDate("01/04/2015")) Then
        ApagarSubtitulo ActiveSheet, "DADOS DA LIGAÇÃO DE ÁGUA (DLIG)"
        ApagarSubtitulo ActiveSheet, "HISTÓRICO DE PAGAMENTOS (HPAG)"
        ApagarSubtitulo ActiveSheet, "SERVIÇOS POR USUÁRIO (CSUS/CTSS)"
        Set rngInicio = ActiveSheet.Cells().Find(what:="OBSERVAÇÕES SOBRE O USUÁRIO (COOB)", lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False)
        ActiveSheet.Rows(rngInicio.Row & ":" & ActiveSheet.UsedRange.Rows.Count).Delete
        
        ' Apaga o HCON após Jan/2017
        Set rngInicio = ActiveSheet.Cells().Find(what:="HISTÓRICO CONSUMOS E LEITURAS (HCON)", lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False)
        Set rngInicio = ActiveSheet.Cells().Find(what:="Mês Referência", After:=rngInicio, lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False)
        Set rngFim = rngInicio
        
        Do
            Set rngFim = rngFim.Offset(1, 0)
        Loop Until rngFim <= 201701 Or rngFim = ""
        
        ActiveSheet.Rows(rngInicio.Row + 1 & ":" & rngFim.Row).Delete
    
        ' Apaga o HCON até Dez/2013
        Do
            Set rngInicio = rngInicio.Offset(1, 0)
        Loop Until rngInicio <= 201401 Or rngInicio = ""
        
        If rngInicio <> "" Then
            Set rngFim = rngInicio
            Set rngFim = rngFim.End(xlDown)
            ActiveSheet.Rows(rngInicio.Row + 1 & ":" & rngFim.Row).Delete
        End If
    
        
    Else
        
        
        ' Apaga o HPAG após 06/2015
        Set rngInicio = ActiveSheet.Cells().Find(what:="HISTÓRICO DE PAGAMENTOS (HPAG)", lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False)
        If rngInicio Is Nothing Then Exit Sub
        Set rngInicio = ActiveSheet.Cells().Find(what:="Referência", After:=rngInicio, lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False)
        Set rngFim = rngInicio
        
        Do
            Set rngFim = rngFim.Offset(1, 0)
        Loop Until rngFim <= 201506
        
        ActiveSheet.Rows(rngInicio.Row + 1 & ":" & rngFim.Row).Delete
        
        ' Apaga o HCON após Jan/2017
        Set rngInicio = ActiveSheet.Cells().Find(what:="HISTÓRICO CONSUMOS E LEITURAS (HCON)", lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False)
        Set rngInicio = ActiveSheet.Cells().Find(what:="Mês Referência", After:=rngInicio, lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False)
        Set rngFim = rngInicio
        
        Do
            Set rngFim = rngFim.Offset(1, 0)
        Loop Until rngFim <= 201701 Or rngFim = ""
        
        ActiveSheet.Rows(rngInicio.Row + 1 & ":" & rngFim.Row).Delete
    
        ' Apaga o HCON até Dez/2013
        Do
            Set rngInicio = rngInicio.Offset(1, 0)
        Loop Until rngInicio <= 201401 Or rngInicio = ""
        
        If rngInicio <> "" Then
            Set rngFim = rngInicio
            Set rngFim = rngFim.End(xlDown)
            ActiveSheet.Rows(rngInicio.Row + 1 & ":" & rngFim.Row).Delete
        End If
    
        ' Apaga o COOB até Mar/2015
        Set rngInicio = ActiveSheet.Cells().Find(what:="OBSERVAÇÕES SOBRE O USUÁRIO (COOB)", lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False)
        Set rngInicio = ActiveSheet.Cells().Find(what:="Data:", After:=rngInicio, lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False)
        Set rngFim = rngInicio.Offset(1, 0)
        
        Do Until CDate(rngFim.Text) <= CDate("31/03/2015") Or (rngFim = "" And rngFim.Offset(1, 0) = "")
            Set rngFim = rngFim.Offset(3, 0)
        Loop
        
        If rngFim = "" Then
            ActiveSheet.Rows(rngInicio.Row - 1 & ":" & rngFim.Row).Delete
        Else
            Set rngInicio = rngFim.Offset(-2, 0)
            ActiveSheet.Rows(rngInicio.Row & ":" & ActiveSheet.UsedRange.Rows.Count).Delete
            
        End If
    End If


End Sub

Sub FazerGrafico(control As IRibbonControl)
''
'' Encontra as colunas, copia-as, cola-as na planilha Plan2, inverte mês e ano e faz o gráfico.
''
    ' Encontra as colunas
    Dim strMatricula As String
    Dim plan As Worksheet
    Dim rngTeste As Range
    Dim strEspaiderOuSCI As String
    
    strEspaiderOuSCI = IIf(ActiveSheet.Range("A1").Formula = "EMBASA - Empresa Baiana de Águas e Saneamento", "SCI", "Espaider")
    
    If strEspaiderOuSCI = "SCI" Then
        Set rngTeste = EncontrarColunas("Histórico de Consumos e Leituras", "Referencia", "Consumo")
        strMatricula = Trim(Application.InputBox(DeterminarTratamento & ", informe o número da matrícula", "Informe o número matrícula", Type:=2))
        
    Else
        Set rngTeste = EncontrarColunas("HISTÓRICO CONSUMOS E LEITURAS (HCON)", "Mês Referência", "Consumo")
        strMatricula = ActiveSheet.Cells().Find(what:="Matrícula:", After:=ActiveCell, lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False).Offset(0, 1).Formula
        strMatricula = Trim(strMatricula)
        
    End If
    
    ' Cria planilha e coloca os valores
    Set plan = ActiveWorkbook.Sheets.Add(Before:=Worksheets(1))
    rngTeste.Copy plan.Cells(1, 1)
    Set rngTeste = plan.UsedRange
    
    ' Inverte ano e mês
    If strEspaiderOuSCI = "Espaider" Then InverteAnoeMes plan.Cells(2, 1)
    
    ' Cria o gráfico
    plan.Cells(2, 4).Activate
    plan.Shapes.AddChart.Select
    With ActiveChart
        .SetSourceData source:=rngTeste
        .ChartTitle.Text = "Gráfico de Consumo - Matrícula " & strMatricula
        .ChartType = xlLine
        .HasLegend = False
        .ChartArea.Width = 500
        .ChartArea.Height = 175
        .Axes(xlValue).MajorUnit = 10
        If .Axes(xlValue).MaximumScale < 20 Then .Axes(xlValue).MaximumScale = 20
    End With

End Sub

Sub ApagarSubtitulo(plan As Worksheet, strSubtitulo As String)
''
'' Procura o subtítulo passado como parâmetro e o apaga, com todas as informações.
''
    Dim rngInicio As Range, rngFim As Range
    
    'Buscar o subtítulo
    Set rngInicio = ActiveSheet.Cells().Find(what:=strSubtitulo, lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False)
    If rngInicio Is Nothing Then Exit Sub
    Set rngFim = rngInicio
    'Selecionar até o final do subtítulo
    Do
        Set rngFim = rngFim.End(xlDown)
    Loop Until rngFim.Interior.Color = 14922894
    
    plan.Rows(rngInicio.Row & ":" & rngFim.Row - 1).Delete
    
End Sub

Function EncontrarColunas(strSubtitulo As String, strColuna1 As String, strColuna2 As String) As Range
''
'' Retorna 2 colunas especificadas abaixo de determinado Subtítulo.
''
    Dim Coluna1 As Range, Coluna2 As Range
    Dim lnUltimaLinha As Long
    
    'Buscar o subtítulo e o cabeçalho da coluna 1 da coluna 2
    Set Coluna1 = ActiveSheet.Cells().Find(what:=strSubtitulo, After:=ActiveCell, lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False)
    Set Coluna1 = ActiveSheet.Cells().Find(what:=strColuna1, After:=Coluna1, lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False)
    Set Coluna2 = ActiveSheet.Cells().Find(what:=strColuna2, After:=Coluna1, lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False)
    
    'Selecionar até o final da coluna (coluna 2 pode ter linhas em branco, portanto seleciona até a última linha de coluna 1)
    lnUltimaLinha = Coluna1.End(xlDown).Row
    Set Coluna1 = Range(Coluna1, Cells(lnUltimaLinha, Coluna1.Column))
    Set Coluna2 = Range(Coluna2, Cells(lnUltimaLinha, Coluna2.Column))
    
    Set EncontrarColunas = Union(Coluna1, Coluna2)

End Function

Sub InverteAnoeMes(rngInicio As Range)
''
'' Inverte o ano-e-mês para mês e ano da coluna inteira abaixo da célula especificada. Para na primeira
'' célula vazia. Ex.: de 201508, passa para 08/2015.
''

    Dim strFormula
    Dim rngCont As Range
    Set rngCont = rngInicio
    
    Do While rngCont.Formula <> ""
        strFormula = Right(rngCont.Formula, 2) & "/" & Left(rngCont.Formula, 4)
        rngCont.Formula = CDate(strFormula)
        rngCont.NumberFormat = "mmm/yyyy"
        Set rngCont = rngCont.Offset(1, 0)
    Loop
    
End Sub

Function RemoverDuplicadosArray(strValores As String, strSeparador As String) As String
''
'' Remove valores duplicados de uma string.
''

    Dim d As Object, i As Integer, strArray() As String
    Set d = CreateObject("Scripting.Dictionary")
    strArray() = Split(strValores, strSeparador)
    
    For i = LBound(strArray) To UBound(strArray)
        d(strArray(i)) = i
    Next i
    
    RemoverDuplicadosArray = Join(d.Keys(), ",")

End Function

Sub DiaUtilAnterior(control As IRibbonControl)
''
'' Retorna o dia útil anterior
''
    Dim dtAudiencia As Date
    
    dtAudiencia = InputBox("Digite o Dia da Audiência:", "Calcular dia útil anterior")
    
    ActiveCell.Formula = WorksheetFunction.WorkDay(dtAudiencia, -1)
    
End Sub

Function RetornaColuna(plan As Worksheet, strNome As String) As Range
''
'' Retorna a coluna de sob um cabeçalho.
''
    Dim rngCont As Range
    
    Set rngCont = plan.Cells().Find(what:=strNome, lookat:=xlWhole, searchorder:=xlByRows)
    Set rngCont = Range(rngCont, plan.Cells(plan.UsedRange.Rows.Count, rngCont.Column))
    Set RetornaColuna = rngCont
    
End Function

Function SeparaPrimeirosNomes(strNome As String, btAteQuantosNomes As Byte) As String
''
'' Separa os primeiros nomes do adverso, retornando com inicial maiúscula.
'' A quantidade de primeiros nomes retornada é btAteQuantosNomes.
''

    Dim arrCont() As String, strCont As String
    Dim btCont As Byte
    
    arrCont = Split(strNome, " ")
    
    ' Se a array tiver menos nomes, eu quero todos (+1 é porque a array começa do 0)
    If UBound(arrCont) + 1 < btAteQuantosNomes Then btAteQuantosNomes = UBound(arrCont) + 1
    
    ' Se o último nome a incluir for partícula "de", "do", "da", "dos", "das", pega mais um nome.
    If (LCase(arrCont(btAteQuantosNomes - 1)) = "de" Or _
        LCase(arrCont(btAteQuantosNomes - 1)) = "do" Or _
        LCase(arrCont(btAteQuantosNomes - 1)) = "da" Or _
        LCase(arrCont(btAteQuantosNomes - 1)) = "dos" Or _
        LCase(arrCont(btAteQuantosNomes - 1)) = "das") And _
        UBound(arrCont) > btAteQuantosNomes - 1 Then _
        btAteQuantosNomes = btAteQuantosNomes + 1
    
    ' Coloca cada nome com iniciais em maiúsculas, exceto "de", "do", "da", "dos", "das", que ficam em minúsculas
    For btCont = 0 To btAteQuantosNomes - 1 Step 1
        If LCase(arrCont(btCont)) = "de" Or LCase(arrCont(btCont)) = "do" Or LCase(arrCont(btCont)) = "da" Or LCase(arrCont(btCont)) = "dos" Or LCase(arrCont(btCont)) = "das" Then
            arrCont(btCont) = LCase(arrCont(btCont))
        Else
            arrCont(btCont) = StrConv(arrCont(btCont), vbProperCase)
        End If
        strCont = strCont & " " & arrCont(btCont)
    Next btCont
    
    SeparaPrimeirosNomes = Trim(strCont)

End Function

Sub InserirArquivo(strCaminhoDocOrigem As String, wdDocDestino As Word.Document)
    Dim finalDoArquivo As Word.Range
    
    On Error GoTo Erros
    
Começo:
    Set finalDoArquivo = wdDocDestino.bookmarks("\EndOfDoc").Range
    finalDoArquivo.InsertFile strCaminhoDocOrigem
    
    On Error GoTo 0
    
    Exit Sub
    
Erros:
    If Err.Number = 4605 Then
        Err.Clear
        GoTo Começo
    Else
        MsgBox DeterminarTratamento & ", ocorreu um erro não esperado: """ & Err.Description & """." & vbCrLf & "Descarte a petição gerada.", vbOKOnly, "Sísifo - Erro na montagem da petição"
        Exit Sub
    End If
    
End Sub

Function ValidaNumeros(ChaveAscii As MSForms.ReturnInteger, Optional strPermitir1 As String, Optional strPermitir2 As String) As Boolean
''
'' Faz uma validação front end, só permitindo números, ponto e barra.
''
    Select Case ChaveAscii
    Case Asc("0") To Asc("9") 'Números são sempre permitidos
        ValidaNumeros = True
    Case Else
        ValidaNumeros = False
    End Select
    
    If strPermitir1 <> "" Then
        If ChaveAscii = Asc(strPermitir1) Then ValidaNumeros = True
    End If
    
    If strPermitir2 <> "" Then
        If ChaveAscii = Asc(strPermitir2) Then ValidaNumeros = True
    End If

End Function

Function InserePontos(strTexto As String) As Boolean
''
'' Fazer uma classificação front end, inserindo pontos entre as categorias.
''
    Dim btInicioUltimaCateg As Byte

    btInicioUltimaCateg = InStrRev(strTexto, "/")
    
    Select Case Len(strTexto) - btInicioUltimaCateg ' Quantidade de caracteres da categoria atual.
    Case 1, 3 ' Se for só 1 ou 3, é hora de colocar um ponto, pois estamos após o número da categoria ou subcategoria (exemplo: "1", "1.7").
        InserePontos = True
    Case Else
        InserePontos = False
    End Select

End Function

Function CodClassificacaoFatura(strClassificacao As String) As String
''
'' Recebe uma classificação tarifária no formato com pontos (1.2.11, etc) e retorna com hífen (12-0011). O parâmetro de entrada pode ter mais de
''   uma categoria, separadas por barras (exemplo: 1.2.3/2.1.1)
''

    Dim strCont() As String, strCont2 As String
    Dim btCont As Byte
    
    strCont = Split(strClassificacao, "/")
    
    For btCont = 0 To UBound(strCont)
        If Trim(strCont(btCont)) = "" Or Trim(strCont(btCont)) = "." Then
            strCont(btCont) = ""
        Else
            strCont2 = strCont(btCont)
            strCont2 = Left(strCont2, 1) & Mid(strCont2, 3, 1) & "-" & Format(Right(strCont2, Len(strCont2) - 4), "0000")
            strCont(btCont) = strCont2
        End If
    Next btCont
    
    CodClassificacaoFatura = Join(strCont, "/")

End Function

Function ClassificacaoExtenso(strClassificacao As String) As String
''
'' Recebe uma classificação tarifária no formato com pontos (1.2.11, etc) e retorna uma string com o significado da classificação tarifária por extenso.
''   O parâmetro de entrada pode ter mais de uma categoria, separadas por barras (exemplo: 1.2.3/2.1.1), nesse caso retornará separado por vírgulas.
''

    Dim strCont() As String, strCont2 As String
    Dim intCont As Integer
    
    strCont = Split(strClassificacao, "/")
    
    For intCont = 0 To UBound(strCont)
        If Trim(strCont(btCont)) = "" Or Trim(strCont(btCont)) = "." Then
            strCont(btCont) = ""
        Else
            strCont2 = CByte(Right(strCont(intCont), Len(strCont(intCont)) - 4))
            
            strCont2 = strCont2 & IIf(strCont2 = 1, " unidade/economia", " unidades/economias")
            
            Select Case Left(strCont(intCont), 3)
            Case "1.1"
                strCont2 = strCont2 & " do tipo residencial intermediária"
            Case "1.2"
                strCont2 = strCont2 & " do tipo residencial normal"
            Case "1.3"
                strCont2 = strCont2 & " do tipo residencial veraneio"
            Case "1.7"
                strCont2 = strCont2 & " do tipo residencial social"
            Case "2.1"
                strCont2 = strCont2 & " do tipo comercial e serviços normal"
            Case "2.2"
                strCont2 = strCont2 & " do tipo comercial e serviços reduzida"
            Case "2.3"
                strCont2 = strCont2 & " do tipo comercial e serviços água bruta"
            Case "2.4"
                strCont2 = strCont2 & " do tipo filantrópica"
            Case "2.5"
                strCont2 = strCont2 & " do tipo derivação rural tratada"
            Case "2.6"
                strCont2 = strCont2 & " do tipo derivação rural bruta"
            Case "3.1"
                strCont2 = strCont2 & " do tipo construção"
            Case "3.2"
                strCont2 = strCont2 & " do tipo industrial"
            Case "4.1"
                strCont2 = strCont2 & " do tipo pública"
            Case "4.2"
                strCont2 = strCont2 & " do tipo contrato demanda"
            End Select
            
            strCont(intCont) = strCont2
        End If
    Next intCont
    
    strCont2 = Join(strCont, ", ")
    intCont = InStrRev(strCont2, ", ")
    If intCont <> 0 Then strCont2 = Left(strCont2, intCont - 1) & " e " & Right(strCont2, intCont + 2)
    
    ClassificacaoExtenso = strCont2

End Function

Function CategoriaExtenso(strClassificacao As String) As String
''
'' Recebe uma classificação tarifária no formato com pontos (1.2.11, etc) e retorna uma string com o significado da classificação tarifária.
''   O parâmetro de entrada pode ter mais de uma categoria, separadas por barras (exemplo: 1.2.3/2.1.1), nesse caso retornará separado por vírgulas.
''

    Dim strCont() As String, strCont2 As String
    Dim btCont As Byte
    
    strCont = Split(strClassificacao, "/")
    
    For btCont = 0 To UBound(strCont)
        If Trim(strCont(btCont)) = "" Or Trim(strCont(btCont)) = "." Then
            strCont(btCont) = ""
        Else
            strCont2 = Left(strCont(btCont), 3)
            Select Case strCont2
            Case "1.1"
                strCont2 = "1.1 = residencial intermediária"
            Case "1.2"
                strCont2 = "1.2 = residencial normal"
            Case "1.3"
                strCont2 = "1.3 = residencial veraneio"
            Case "1.7"
                strCont2 = "1.7 = residencial social"
            Case "2.1"
                strCont2 = "2.1 = comercial e serviços normal"
            Case "2.2"
                strCont2 = "2.2 = comercial e serviços reduzida"
            Case "2.3"
                strCont2 = "2.3 = comercial e serviços água bruta"
            Case "2.4"
                strCont2 = "2.4 = filantrópica"
            Case "2.5"
                strCont2 = "2.5 = derivação rural tratada"
            Case "2.6"
                strCont2 = "2.6 = derivação rural bruta"
            Case "3.1"
                strCont2 = "3.1 = construção"
            Case "3.2"
                strCont2 = "3.2 = industrial"
            Case "4.1"
                strCont2 = "4.1 = pública"
            Case "4.2"
                strCont2 = "4.2 = contrato demanda"
            End Select
            
            strCont(btCont) = strCont2
        End If
    Next btCont
    
    CategoriaExtenso = Join(strCont, ", ")

End Function

Function EconomiasExtenso(strClassificacao As String) As String
''
'' Recebe uma classificação tarifária no formato com pontos (1.2.11, etc) e retorna uma string com a quantidade de economias.
''   O parâmetro de entrada pode ter mais de uma categoria, separadas por barras (exemplo: 1.2.3/2.1.1), nesse caso retornará separado por vírgulas.
''

    Dim strCont() As String, strCont2 As String
    Dim btCont As Byte
    
    strCont = Split(strClassificacao, "/")
    
    For btCont = 0 To UBound(strCont)
        If Trim(strCont(btCont)) = "" Or Trim(strCont(btCont)) = "." Then
            strCont(btCont) = ""
        Else
            strCont2 = Right(strCont(btCont), Len(strCont(btCont)) - 4)
            strCont2 = CByte(strCont2)
            strCont(btCont) = strCont2
        End If
    Next btCont
    
    EconomiasExtenso = "no caso, " & Join(strCont, " e ")

End Function

Function PegarDiretorioPPJCM() As String
''
'' Pega o endereço da pasta PPJCM cadastrada no Sísifo.
''
    
    Dim strPasta As String
    Dim plan As Worksheet
    
    Set plan = ThisWorkbook.Sheets("cfConfigurações")
    strPasta = plan.Cells().Find(what:="Caminho da pasta do Sisifo", lookat:=xlWhole, searchorder:=xlByRows, MatchCase:=False).Offset(1, 0).Text
    PegarDiretorioPPJCM = strPasta
    
End Function

Function SalvarDiretorioPPJCM(strPasta As String) As Boolean
''
'' Salva o endereço da pasta PPJCM no Sísifo.
''
    
    Dim plan As Worksheet
    Dim rngCont As Range
    
    Set plan = ThisWorkbook.Sheets("cfConfigurações")
    Set rngCont = plan.Cells().Find(what:="Caminho da pasta do Sisifo", lookat:=xlWhole, searchorder:=xlByRows, MatchCase:=False).Offset(1, 0)
    rngCont.Formula = strPasta
    
    If rngCont.Formula = strPasta Then
        SalvarDiretorioPPJCM = True
    Else
        SalvarDiretorioPPJCM = False
    End If
    
    Set plan = Nothing
    Set rngCont = Nothing
    
End Function

Function PerguntarDiretorio(strTitulo As String, bolMultiselecao As Boolean) As String
    Dim objPopup As FileDialog
    Dim strPasta As String
    
    Set objPopup = Application.FileDialog(msoFileDialogFolderPicker)
    With objPopup
        .Title = strTitulo
        .AllowMultiSelect = bolMultiselecao
        .InitialFileName = CaminhoDesktop
        If .Show <> -1 Then GoTo AtribuirValor
        strPasta = .SelectedItems(1)
    End With
    
AtribuirValor:
    PerguntarDiretorio = strPasta
    Set objPopup = Nothing
    
End Function

Function PerguntarNomeArquivo(strTitulo As String, strNomePadraoComCaminho As String) As String
    Dim objPopup As New APICxDialogo
    Dim lngJanelaOrigemHwnd As Long, lngAplicacaohWnd As Long, lngResultado As Long
    Dim strNomeArquivo As String, strCaminho As String, strFiltro As String
    
    lngJanelaOrigemHwnd = ActiveWorkbook.Windows(1).hWnd
    lngAplicacaohWnd = Application.hWnd
    strFiltro = "XML do Sísifo (*.ssfx)" & Chr(0) & "*.ssfx"
    
    lngResult = objPopup.SaveFileDialog(lngJanelaOrigemHwnd, lngAplicacaohWnd, strNomePadraoComCaminho, strTitulo, strFiltro)
    
    If objPopup.Status = True Then
        strNomeArquivo = objPopup.Name
        strNomeArquivo = Replace(strNomeArquivo, ChrW(0), "")
        If Right(strNomeArquivo, 5) <> ".ssfx" Then strNomeArquivo = strNomeArquivo & ".ssfx"
        PerguntarNomeArquivo = strNomeArquivo
    Else
        PerguntarNomeArquivo = ""
    End If
    
End Function
