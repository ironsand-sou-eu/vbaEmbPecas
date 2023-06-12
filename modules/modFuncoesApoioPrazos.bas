Attribute VB_Name = "modFuncoesApoioPrazos"
Option Explicit

Public bolSsfPrazosBotaoPdfPressionado As Boolean

Sub FechaConfigPrazosVisivel(ByVal controle As IRibbonControl, Optional ByRef returnedVal)
    
    FechaConfigVisivel "btFechaConfigPrazos", controle, returnedVal
        
End Sub

Private Sub AoCarregarRibbonPrazos(Ribbon As IRibbonUI)
    ' Chama a fun��o geral AoCarregarRibbon com os par�metros corretos.
    Dim plPlan As Worksheet
    Set plPlan = ThisWorkbook.Sheets("cfConfigura��es")
    
    AoCarregarRibbon plPlan, Ribbon
    
End Sub

Sub LiberarEdicaoPrazos(ByVal controle As IRibbonControl)
' Chama a fun��o geral LiberarEdicao
    LiberarEdicao ThisWorkbook
    
End Sub

Sub RestringirEdicaoRibbonPrazos(ByVal controle As IRibbonControl)
' Chama a fun��o geral RestringirEdicaoRibbon
    RestringirEdicaoRibbon ThisWorkbook, controle
    
End Sub

Public Sub ssfPrazosTglGerarPDF_getPressed(control As IRibbonControl, ByRef returnedVal)
' Bot�o de gerar PDF
    returnedVal = bolSsfPrazosBotaoPdfPressionado
End Sub

Public Sub ssfPrazosTglGerarPDF_onAction(control As IRibbonControl, ByRef cancelDefault)
' Bot�o de gerar PDF
    bolSsfPrazosBotaoPdfPressionado = Not bolSsfPrazosBotaoPdfPressionado
End Sub

Function BuscarCaminhoPrograma() As String
    Dim rngCaminho As Range
    
    Set rngCaminho = ThisWorkbook.Sheets("cfConfigura��es").Cells.Find(what:="Caminho da pasta do Sisifo", lookat:=xlWhole)
    
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
        If LCase(.Cells(1, 3).Formula) = "matr�cula principal" Then .Columns(3).ColumnWidth = 14
        If .Cells(1, 7).Formula = "Observa��es Prov." Then .Columns(7).ColumnWidth = 40
        If .Cells(1, 8).Formula = "Andamento" Then .Columns(8).ColumnWidth = 35
        If .Cells(1, 10).Formula = "Obs. do Andamento" Then .Columns(10).ColumnWidth = 30
        If .Cells(1, 11).Formula = "Ju�zo" Then .Columns(11).ColumnWidth = 40
        If .AutoFilter Is Nothing Then .UsedRange.AutoFilter
        
    End With
    
End Sub

Sub RecortarHistoricoCCR(control As IRibbonControl)
''
'' Em um hist�rico de matr�cula do Espaider, apaga Categorias Tarif�rias, Dados da Liga��o de Esgoto, HPAG posterior a 05/2015,
''    HCON exceto 2014 a 2016, Consulta Notifica��es de D�bito, Servi�os posteriores a 05/2015, Informa��es de parcelamento,
''    COOB posterior a 05/2015.
''
    
    'Dim plan As Worksheet
    Dim rngInicio As Range, rngFim As Range
    Dim strUltimoCorte As String ', strUltimaReligacao As String
    
    'Dim rngTeste As Range
    
    ' Apaga os seguintes cabe�alhos:
    ApagarSubtitulo ActiveSheet, "CATEGORIAS TARIF�RIAS"
    ApagarSubtitulo ActiveSheet, "DADOS DA LIGA��O DE ESGOTO"
    ApagarSubtitulo ActiveSheet, "CONSULTA NOTIFICA��ES DE  D�BITO (CNOT)"
    ApagarSubtitulo ActiveSheet, "INFORMA��ES DO PARCELAMENTO - HIST�RICO (IPAR)"
    
    ' Se o �ltimo corte foi antes ou depois da CCR, apaga os outros cabe�alhos.
    strUltimoCorte = ActiveSheet.Cells().Find(what:="�ltimo corte:", lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False).Offset(0, 1).Text
    strUltimaReligacao = ActiveSheet.Cells().Find(what:="�ltima religa��o:", lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False).Offset(0, 1).Text
    If strUltimoCorte = "31/12/9999" Or _
        (CDate(strUltimaReligacao) >= CDate(strUltimoCorte) And CDate(strUltimaReligacao) < CDate("01/04/2015")) Then
        ApagarSubtitulo ActiveSheet, "DADOS DA LIGA��O DE �GUA (DLIG)"
        ApagarSubtitulo ActiveSheet, "HIST�RICO DE PAGAMENTOS (HPAG)"
        ApagarSubtitulo ActiveSheet, "SERVI�OS POR USU�RIO (CSUS/CTSS)"
        Set rngInicio = ActiveSheet.Cells().Find(what:="OBSERVA��ES SOBRE O USU�RIO (COOB)", lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False)
        ActiveSheet.Rows(rngInicio.Row & ":" & ActiveSheet.UsedRange.Rows.Count).Delete
        
        ' Apaga o HCON ap�s Jan/2017
        Set rngInicio = ActiveSheet.Cells().Find(what:="HIST�RICO CONSUMOS E LEITURAS (HCON)", lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False)
        Set rngInicio = ActiveSheet.Cells().Find(what:="M�s Refer�ncia", After:=rngInicio, lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False)
        Set rngFim = rngInicio
        
        Do
            Set rngFim = rngFim.Offset(1, 0)
        Loop Until rngFim <= 201701 Or rngFim = ""
        
        ActiveSheet.Rows(rngInicio.Row + 1 & ":" & rngFim.Row).Delete
    
        ' Apaga o HCON at� Dez/2013
        Do
            Set rngInicio = rngInicio.Offset(1, 0)
        Loop Until rngInicio <= 201401 Or rngInicio = ""
        
        If rngInicio <> "" Then
            Set rngFim = rngInicio
            Set rngFim = rngFim.End(xlDown)
            ActiveSheet.Rows(rngInicio.Row + 1 & ":" & rngFim.Row).Delete
        End If
    
        
    Else
        
        
        ' Apaga o HPAG ap�s 06/2015
        Set rngInicio = ActiveSheet.Cells().Find(what:="HIST�RICO DE PAGAMENTOS (HPAG)", lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False)
        If rngInicio Is Nothing Then Exit Sub
        Set rngInicio = ActiveSheet.Cells().Find(what:="Refer�ncia", After:=rngInicio, lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False)
        Set rngFim = rngInicio
        
        Do
            Set rngFim = rngFim.Offset(1, 0)
        Loop Until rngFim <= 201506
        
        ActiveSheet.Rows(rngInicio.Row + 1 & ":" & rngFim.Row).Delete
        
        ' Apaga o HCON ap�s Jan/2017
        Set rngInicio = ActiveSheet.Cells().Find(what:="HIST�RICO CONSUMOS E LEITURAS (HCON)", lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False)
        Set rngInicio = ActiveSheet.Cells().Find(what:="M�s Refer�ncia", After:=rngInicio, lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False)
        Set rngFim = rngInicio
        
        Do
            Set rngFim = rngFim.Offset(1, 0)
        Loop Until rngFim <= 201701 Or rngFim = ""
        
        ActiveSheet.Rows(rngInicio.Row + 1 & ":" & rngFim.Row).Delete
    
        ' Apaga o HCON at� Dez/2013
        Do
            Set rngInicio = rngInicio.Offset(1, 0)
        Loop Until rngInicio <= 201401 Or rngInicio = ""
        
        If rngInicio <> "" Then
            Set rngFim = rngInicio
            Set rngFim = rngFim.End(xlDown)
            ActiveSheet.Rows(rngInicio.Row + 1 & ":" & rngFim.Row).Delete
        End If
    
        ' Apaga o COOB at� Mar/2015
        Set rngInicio = ActiveSheet.Cells().Find(what:="OBSERVA��ES SOBRE O USU�RIO (COOB)", lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False)
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
'' Encontra as colunas, copia-as, cola-as na planilha Plan2, inverte m�s e ano e faz o gr�fico.
''
    ' Encontra as colunas
    Dim strMatricula As String
    Dim plan As Worksheet
    Dim rngTeste As Range
    Dim strEspaiderOuSCI As String
    
    strEspaiderOuSCI = IIf(ActiveSheet.Range("A1").Formula = "EMBASA - Empresa Baiana de �guas e Saneamento", "SCI", "Espaider")
    
    If strEspaiderOuSCI = "SCI" Then
        Set rngTeste = EncontrarColunas("Hist�rico de Consumos e Leituras", "Referencia", "Consumo")
        strMatricula = Trim(Application.InputBox(DeterminarTratamento & ", informe o n�mero da matr�cula", "Informe o n�mero matr�cula", Type:=2))
        
    Else
        Set rngTeste = EncontrarColunas("HIST�RICO CONSUMOS E LEITURAS (HCON)", "M�s Refer�ncia", "Consumo")
        strMatricula = ActiveSheet.Cells().Find(what:="Matr�cula:", After:=ActiveCell, lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False).Offset(0, 1).Formula
        strMatricula = Trim(strMatricula)
        
    End If
    
    ' Cria planilha e coloca os valores
    Set plan = ActiveWorkbook.Sheets.Add(Before:=Worksheets(1))
    rngTeste.Copy plan.Cells(1, 1)
    Set rngTeste = plan.UsedRange
    
    ' Inverte ano e m�s
    If strEspaiderOuSCI = "Espaider" Then InverteAnoeMes plan.Cells(2, 1)
    
    ' Cria o gr�fico
    plan.Cells(2, 4).Activate
    plan.Shapes.AddChart.Select
    With ActiveChart
        .SetSourceData source:=rngTeste
        .ChartTitle.Text = "Gr�fico de Consumo - Matr�cula " & strMatricula
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
'' Procura o subt�tulo passado como par�metro e o apaga, com todas as informa��es.
''
    Dim rngInicio As Range, rngFim As Range
    
    'Buscar o subt�tulo
    Set rngInicio = ActiveSheet.Cells().Find(what:=strSubtitulo, lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False)
    If rngInicio Is Nothing Then Exit Sub
    Set rngFim = rngInicio
    'Selecionar at� o final do subt�tulo
    Do
        Set rngFim = rngFim.End(xlDown)
    Loop Until rngFim.Interior.Color = 14922894
    
    plan.Rows(rngInicio.Row & ":" & rngFim.Row - 1).Delete
    
End Sub

Function EncontrarColunas(strSubtitulo As String, strColuna1 As String, strColuna2 As String) As Range
''
'' Retorna 2 colunas especificadas abaixo de determinado Subt�tulo.
''
    Dim Coluna1 As Range, Coluna2 As Range
    Dim lnUltimaLinha As Long
    
    'Buscar o subt�tulo e o cabe�alho da coluna 1 da coluna 2
    Set Coluna1 = ActiveSheet.Cells().Find(what:=strSubtitulo, After:=ActiveCell, lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False)
    Set Coluna1 = ActiveSheet.Cells().Find(what:=strColuna1, After:=Coluna1, lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False)
    Set Coluna2 = ActiveSheet.Cells().Find(what:=strColuna2, After:=Coluna1, lookat:=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:=False)
    
    'Selecionar at� o final da coluna (coluna 2 pode ter linhas em branco, portanto seleciona at� a �ltima linha de coluna 1)
    lnUltimaLinha = Coluna1.End(xlDown).Row
    Set Coluna1 = Range(Coluna1, Cells(lnUltimaLinha, Coluna1.Column))
    Set Coluna2 = Range(Coluna2, Cells(lnUltimaLinha, Coluna2.Column))
    
    Set EncontrarColunas = Union(Coluna1, Coluna2)

End Function

Sub InverteAnoeMes(rngInicio As Range)
''
'' Inverte o ano-e-m�s para m�s e ano da coluna inteira abaixo da c�lula especificada. Para na primeira
'' c�lula vazia. Ex.: de 201508, passa para 08/2015.
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
'' Retorna o dia �til anterior
''
    Dim dtAudiencia As Date
    
    dtAudiencia = InputBox("Digite o Dia da Audi�ncia:", "Calcular dia �til anterior")
    
    ActiveCell.Formula = WorksheetFunction.WorkDay(dtAudiencia, -1)
    
End Sub

Function RetornaColuna(plan As Worksheet, strNome As String) As Range
''
'' Retorna a coluna de sob um cabe�alho.
''
    Dim rngCont As Range
    
    Set rngCont = plan.Cells().Find(what:=strNome, lookat:=xlWhole, searchorder:=xlByRows)
    Set rngCont = Range(rngCont, plan.Cells(plan.UsedRange.Rows.Count, rngCont.Column))
    Set RetornaColuna = rngCont
    
End Function

Function SeparaPrimeirosNomes(strNome As String, btAteQuantosNomes As Byte) As String
''
'' Separa os primeiros nomes do adverso, retornando com inicial mai�scula.
'' A quantidade de primeiros nomes retornada � btAteQuantosNomes.
''

    Dim arrCont() As String, strCont As String
    Dim btCont As Byte
    
    arrCont = Split(strNome, " ")
    
    ' Se a array tiver menos nomes, eu quero todos (+1 � porque a array come�a do 0)
    If UBound(arrCont) + 1 < btAteQuantosNomes Then btAteQuantosNomes = UBound(arrCont) + 1
    
    ' Se o �ltimo nome a incluir for part�cula "de", "do", "da", "dos", "das", pega mais um nome.
    If (LCase(arrCont(btAteQuantosNomes - 1)) = "de" Or _
        LCase(arrCont(btAteQuantosNomes - 1)) = "do" Or _
        LCase(arrCont(btAteQuantosNomes - 1)) = "da" Or _
        LCase(arrCont(btAteQuantosNomes - 1)) = "dos" Or _
        LCase(arrCont(btAteQuantosNomes - 1)) = "das") And _
        UBound(arrCont) > btAteQuantosNomes - 1 Then _
        btAteQuantosNomes = btAteQuantosNomes + 1
    
    ' Coloca cada nome com iniciais em mai�sculas, exceto "de", "do", "da", "dos", "das", que ficam em min�sculas
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
    
Come�o:
    Set finalDoArquivo = wdDocDestino.bookmarks("\EndOfDoc").Range
    finalDoArquivo.InsertFile strCaminhoDocOrigem
    
    On Error GoTo 0
    
    Exit Sub
    
Erros:
    If Err.Number = 4605 Then
        Err.Clear
        GoTo Come�o
    Else
        MsgBox DeterminarTratamento & ", ocorreu um erro n�o esperado: """ & Err.Description & """." & vbCrLf & "Descarte a peti��o gerada.", vbOKOnly, "S�sifo - Erro na montagem da peti��o"
        Exit Sub
    End If
    
End Sub

Function ValidaNumeros(ChaveAscii As MSForms.ReturnInteger, Optional strPermitir1 As String, Optional strPermitir2 As String) As Boolean
''
'' Faz uma valida��o front end, s� permitindo n�meros, ponto e barra.
''
    Select Case ChaveAscii
    Case Asc("0") To Asc("9") 'N�meros s�o sempre permitidos
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
'' Fazer uma classifica��o front end, inserindo pontos entre as categorias.
''
    Dim btInicioUltimaCateg As Byte

    btInicioUltimaCateg = InStrRev(strTexto, "/")
    
    Select Case Len(strTexto) - btInicioUltimaCateg ' Quantidade de caracteres da categoria atual.
    Case 1, 3 ' Se for s� 1 ou 3, � hora de colocar um ponto, pois estamos ap�s o n�mero da categoria ou subcategoria (exemplo: "1", "1.7").
        InserePontos = True
    Case Else
        InserePontos = False
    End Select

End Function

Function CodClassificacaoFatura(strClassificacao As String) As String
''
'' Recebe uma classifica��o tarif�ria no formato com pontos (1.2.11, etc) e retorna com h�fen (12-0011). O par�metro de entrada pode ter mais de
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
'' Recebe uma classifica��o tarif�ria no formato com pontos (1.2.11, etc) e retorna uma string com o significado da classifica��o tarif�ria por extenso.
''   O par�metro de entrada pode ter mais de uma categoria, separadas por barras (exemplo: 1.2.3/2.1.1), nesse caso retornar� separado por v�rgulas.
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
                strCont2 = strCont2 & " do tipo residencial intermedi�ria"
            Case "1.2"
                strCont2 = strCont2 & " do tipo residencial normal"
            Case "1.3"
                strCont2 = strCont2 & " do tipo residencial veraneio"
            Case "1.7"
                strCont2 = strCont2 & " do tipo residencial social"
            Case "2.1"
                strCont2 = strCont2 & " do tipo comercial e servi�os normal"
            Case "2.2"
                strCont2 = strCont2 & " do tipo comercial e servi�os reduzida"
            Case "2.3"
                strCont2 = strCont2 & " do tipo comercial e servi�os �gua bruta"
            Case "2.4"
                strCont2 = strCont2 & " do tipo filantr�pica"
            Case "2.5"
                strCont2 = strCont2 & " do tipo deriva��o rural tratada"
            Case "2.6"
                strCont2 = strCont2 & " do tipo deriva��o rural bruta"
            Case "3.1"
                strCont2 = strCont2 & " do tipo constru��o"
            Case "3.2"
                strCont2 = strCont2 & " do tipo industrial"
            Case "4.1"
                strCont2 = strCont2 & " do tipo p�blica"
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
'' Recebe uma classifica��o tarif�ria no formato com pontos (1.2.11, etc) e retorna uma string com o significado da classifica��o tarif�ria.
''   O par�metro de entrada pode ter mais de uma categoria, separadas por barras (exemplo: 1.2.3/2.1.1), nesse caso retornar� separado por v�rgulas.
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
                strCont2 = "1.1 = residencial intermedi�ria"
            Case "1.2"
                strCont2 = "1.2 = residencial normal"
            Case "1.3"
                strCont2 = "1.3 = residencial veraneio"
            Case "1.7"
                strCont2 = "1.7 = residencial social"
            Case "2.1"
                strCont2 = "2.1 = comercial e servi�os normal"
            Case "2.2"
                strCont2 = "2.2 = comercial e servi�os reduzida"
            Case "2.3"
                strCont2 = "2.3 = comercial e servi�os �gua bruta"
            Case "2.4"
                strCont2 = "2.4 = filantr�pica"
            Case "2.5"
                strCont2 = "2.5 = deriva��o rural tratada"
            Case "2.6"
                strCont2 = "2.6 = deriva��o rural bruta"
            Case "3.1"
                strCont2 = "3.1 = constru��o"
            Case "3.2"
                strCont2 = "3.2 = industrial"
            Case "4.1"
                strCont2 = "4.1 = p�blica"
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
'' Recebe uma classifica��o tarif�ria no formato com pontos (1.2.11, etc) e retorna uma string com a quantidade de economias.
''   O par�metro de entrada pode ter mais de uma categoria, separadas por barras (exemplo: 1.2.3/2.1.1), nesse caso retornar� separado por v�rgulas.
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
'' Pega o endere�o da pasta PPJCM cadastrada no S�sifo.
''
    
    Dim strPasta As String
    Dim plan As Worksheet
    
    Set plan = ThisWorkbook.Sheets("cfConfigura��es")
    strPasta = plan.Cells().Find(what:="Caminho da pasta do Sisifo", lookat:=xlWhole, searchorder:=xlByRows, MatchCase:=False).Offset(1, 0).Text
    PegarDiretorioPPJCM = strPasta
    
End Function

Function SalvarDiretorioPPJCM(strPasta As String) As Boolean
''
'' Salva o endere�o da pasta PPJCM no S�sifo.
''
    
    Dim plan As Worksheet
    Dim rngCont As Range
    
    Set plan = ThisWorkbook.Sheets("cfConfigura��es")
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
    strFiltro = "XML do S�sifo (*.ssfx)" & Chr(0) & "*.ssfx"
    
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
