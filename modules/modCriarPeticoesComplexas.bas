Attribute VB_Name = "modCriarPeticoesComplexas"
Option Explicit

Sub MontarPeticaoComplexa(strCausaPedir As String, strPeticao As String, strJuizoEspaider As String, strTermoInicialPrazo As String)
''
'' Monta uma petição complexa. strCausaPedir aceita qualquer valor das causas de pedir do Espaider.
''  strPeticao aceita os valores "Contestação", "Recurso Inominado" ou "Contrarrazões de RI".
''  strJuizoEspaider aceita valores que contenham as expressões "Juizado" ou ???
''

    Dim appword As Object
    Dim wdDocPeticao As Word.Document
    Dim plan As Worksheet
    Dim strTopicos As String, strVariaveis As String, arrVariaveis() As String, strJuizo As String, strJuizoResumido As String
    Dim strOrgao As String, strCausaPedirParadigma As String, strCaminhoDocOrigem As String
    Dim intColunaTopicos As Long
    Dim rngCont As Excel.Range
    Dim wdrngCont As Word.Range
    Dim form As Variant, X As Variant, z As Variant
    Dim byteCont As Byte, btTabela As Byte
    Dim bolGrafico As Boolean, bolGerarPDF As Boolean
    Dim varCont As Variant
    
    ' Estabelece qual o órgão em que tramita o processo. Se houver a expressão "Juizado", considera que é juizado. Se houver "procuradoria",
    '  "coordenadoria", "coordenação", "Procon" ou "Codecon", considera que é Procon. Em qualquer outro caso, considera que é Vara cível.
    '   Não diferencia maiúsculas de minúsculas.
    If InStr(1, LCase(strJuizoEspaider), "juizado") <> 0 Or InStr(1, LCase(strJuizoEspaider), "sje") <> 0 Then
        strOrgao = "JEC"
    
    ElseIf InStr(1, LCase(strJuizoEspaider), "procuradoria") <> 0 Or InStr(1, LCase(strJuizoEspaider), "coordenadoria") <> 0 Or _
            InStr(1, LCase(strJuizoEspaider), "coordenação") <> 0 Or InStr(1, LCase(strJuizoEspaider), "Procon") <> 0 Or _
            InStr(1, LCase(strJuizoEspaider), "Codecon") <> 0 Then
        strOrgao = "Procon"
        
    Else
        strOrgao = "VC"
    
    End If
    
    
    '''''''''''''''''''''''''''''''''''''''''
    '' Define as causas de pedir-paradigma ''
    '''''''''''''''''''''''''''''''''''''''''
    
    Select Case strCausaPedir
    ' Primeiro, os casos em que não muda, porque há descrição específica da peça
    Case "Revisão de consumo elevado", "Corte no fornecimento", "Negativação no SPC", "Realizar ligação de água", _
        "Cobrança de esgoto em imóvel não ligado à rede", "Cobrança de esgoto com água cortada", "Classificação tarifa ou qtd. de economias", _
        "Suspeita de by-pass", "Débito de terceiro", "Desabastecimentos por período e causa determinados", "Desabastecimento CCR 04/2015", _
        "Desabastecimento Uruguai 09/2016", "Desabastecimento Liberdade 10/2017", "Desabastecimento Apagão Xingu 03/2018", _
        "Vaz. água ou extravas. esgoto com danos a patrimônio/morais", "Obra da Embasa com danos a patrimônio/morais", _
        "Acidente com pessoa/veículo em buraco", "Acidente com veículo (colisão ou atropelamento)", "Fixo de esgoto"
        strCausaPedirParadigma = strCausaPedir
        
    Case "Consumo elevado com corte"
        strCausaPedirParadigma = "Revisão de consumo elevado"
    
    Case "Desmembramento de ligações"
        strCausaPedirParadigma = "Realizar ligação de água"
    
    Case "Multa por infração"
        strCausaPedirParadigma = "Suspeita de by-pass"
    
    Case "Desabastecimento CCR SEGUNDO ACIDENTE 04/2016", "Desabastecimento Faz. Grande do Retiro Verão 2019", _
        "Desabastecimento Nova Brasília de Itapuã 10/2018", "Desabastecimento Novo Horizonte vários períodos", "Desabastecimento Pernambués 11/2018", _
        "Desabastecimento Subúrbio Ferroviário 02/2017", "Desabastecimento Sussuarana 12/2018", "Desabastecimento Res. Bosque das Bromélias 10/2019", "Irregularidade no abastecimento de água"
        strCausaPedirParadigma = "Desabastecimentos por período e causa determinados"
    
    Case Else
        strCausaPedirParadigma = "Outros vícios"
        
    End Select

    
    
    '''''''''''''''''''''''''
    '' Mostra o formulário ''
    '''''''''''''''''''''''''
    
    Set form = ConfigurarFormulario(strCausaPedirParadigma, strOrgao, strPeticao)
    If form Is Nothing Then Exit Sub
    
    If strCausaPedir = "Consumo elevado com corte" Then
        form.chbDanMorCorte.Value = True
    End If
    
    form.Show
    
    If form.chbDeveGerar.Value = False Then Exit Sub
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' Colhe os tópicos, na variável strTopicos; também ajusta bolGrafico e btTabela, se houve gráficos ou tabelas ''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    strTopicos = ColherTopicosGeral(strCausaPedir, strCausaPedirParadigma, strPeticao, form, bolGrafico, btTabela)
    
    
    ''''''''''''''''''''''''''''''
    '' Colhe outras informações ''
    ''''''''''''''''''''''''''''''
    
    ' Pega o Juízo na redação do Espaider, depois pega o juízo na redação para ficar na planilha.
    strJuizo = BuscaJuizo(strJuizoEspaider) ' strJuizo assume a redação longa do juízo
    If strPeticao = "RI/Apelação" Or strPeticao = "Contrarrazões de RI/Apelação" Then strJuizoResumido = BuscaJuizo(strJuizo) 'strJuizoResumido assume a redação curta do juízo
    
    ' Criar o documento a partir do modelo
    Set appword = New Word.Application
    appword.Visible = True
    Set wdDocPeticao = appword.Documents.Add(BuscarCaminhoPrograma & "modelos-automaticos\PPJCM Modelo.dotx")
    
    '''''''''''''''''''''''''''''
    '' Pega a planilha correta ''
    '''''''''''''''''''''''''''''
    
    Set plan = DescobrirPlanilhaDeEstrutura(strPeticao, strOrgao)
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' Copiar os tópicos selecionados para o documento-destino ''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ' Na planilha adequada, busca o cabeçalho de tópicos da petição e descobre qual a primeira linha da coluna núm-tópico
    
    Set rngCont = plan.Cells().Find(what:=strCausaPedirParadigma, lookat:=xlWhole, searchorder:=xlByRows, MatchCase:=False).Offset(1, 0).Offset(0, 3)
    
    ' Desce de célula em célula, copiando e colando cada vez que o valor da célula for zero ou estiver contido em strTopicos (cercado por duas vírgulas)
    
    Do Until rngCont.Offset(1, 0).Text = ""
        Set rngCont = rngCont.Offset(1, 0)
        If rngCont.Formula = "0" Or InStr(1, strTopicos, ",," & rngCont.Formula & ",,") Then
            
            ' Pega o nome à direita, procura na pasta de tópicos personalizados e, não havendo, na Frankenstein normal.
            If Dir(BuscarCaminhoPrograma & "Frankenstein\Tópicos personalizados\" & rngCont.Offset(0, 1).Formula) <> "" Then
                strCaminhoDocOrigem = BuscarCaminhoPrograma & "Frankenstein\Tópicos personalizados\" & rngCont.Offset(0, 1).Formula
            Else
                strCaminhoDocOrigem = BuscarCaminhoPrograma & "Frankenstein\" & rngCont.Offset(0, 1).Formula
            End If
            
            ' Adiciona o documento respectivo ao arquivo Word principal (na ordem das linhas da planilha "plan").
            InserirArquivo strCaminhoDocOrigem, wdDocPeticao
            
            ' Se houver conteúdo duas colunas à direita, armazenar numa array de variáveis.
            If Trim(rngCont.Offset(0, 2).Formula) <> "" Then
                strVariaveis = strVariaveis & rngCont.Offset(0, 2).Formula & ","
            End If
        End If
    Loop
    
    wdDocPeticao.Activate
    wdDocPeticao.Paragraphs.Last.Range.Delete 'Apaga o parágrafo vazio que fica no final.
    
    ' Substituir as variáveis Comarca, Número e Adverso.
    
    With appword.Selection.Find
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .Text = "<Juízo>"
        .Replacement.Text = strJuizo
        .Execute Replace:=wdReplaceAll
        'Se o juízo não foi encontrado no Sísifo, avisa.
        If strJuizo = "" Then MsgBox "Juízo não encontrado na base de dados no Sísifo. Lembre-se de acrescentá-lo manualmente no endereçamento da petição.", _
            vbCritical + vbOKOnly, "Alerta - juízo não encontrado"

        If strPeticao = "RI/Apelação" Or strPeticao = "Contrarrazões de RI/Apelação" Then
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .Text = "<Juízo-Resumido>"
            .Replacement.Text = strJuizoResumido
            .Execute Replace:=wdReplaceAll
            'Se o juízo não foi encontrado no Sísifo, avisa.
            If strJuizoResumido = "" Then MsgBox "Nome resumido do juízo não encontrado na base de dados no Sísifo. Lembre-se de acrescentá-lo manualmente no endereçamento da petição.", _
                vbCritical + vbOKOnly, "Alerta - juízo resumido não encontrado"
        End If

        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .Text = "<Número>"
        .Replacement.Text = ActiveSheet.Cells(ActiveCell.Row, 1).Formula
        .Execute Replace:=wdReplaceAll

        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .Text = "<Adverso>"
        .Replacement.Text = ActiveSheet.Cells(ActiveCell.Row, 2).Formula
        .Execute Replace:=wdReplaceAll

        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .Text = "<data>"
        .Replacement.Text = Format(Date, "dd") & " de " & Format(Date, "mmmm") & " de " & Format(Date, "yyyy")
        .Execute Replace:=wdReplaceAll
    End With

    ''''''''''''''''''''''''''''''''''''
    '' Substituir as demais variáveis.''
    ''''''''''''''''''''''''''''''''''''
    
    strVariaveis = RemoverDuplicadosArray(strVariaveis, ",")
    arrVariaveis = Split(strVariaveis, ",")
    
    For Each z In arrVariaveis
        If z = "data-audiência-conciliação" Then
            X = ObterVariaveis(CStr(z), strCausaPedirParadigma, form, strTermoInicialPrazo)
        Else
            X = ObterVariaveis(CStr(z), strCausaPedirParadigma, form)
        End If
            
        With appword.Selection.Find
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .Text = "<" & z & ">"
            .Replacement.Text = X
            .Execute Replace:=wdReplaceAll
        End With
    Next z
    
    ' Substituir os gráficos, se houver
    If bolGrafico Then
        On Error Resume Next
        Set varCont = Application.InputBox(Prompt:=DeterminarTratamento & ", no Excel, clique em qualquer célula na planilha que contém o gráfico de consumo, o qual será inserido na petição.", Title:="Sísifo - Selecione o gráfico!", Type:=8)
        
        If Err.Number <> 0 Or varCont.Worksheet.ChartObjects.Count = 0 Then
            MsgBox DeterminarTratamento & ", não escolheste uma celula numa planilha com gráfico. A petição será gerada sem o gráfico; lembre-se de Adicionar " & _
                "um gráfico ou tela.", vbCritical + vbOKOnly, "Sísifo - Petição gerada sem gráfico"
        Else
            On Error GoTo 0
            varCont.Worksheet.ChartObjects(1).Copy
            
            wdDocPeticao.Activate
            With appword.Selection.Find
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .Text = "<gráfico-de-consumo>^p"
                .Replacement.Text = ""
                .Execute
            End With
            
            appword.Selection.Paste
            wdDocPeticao.InlineShapes(1).ConvertToShape
            wdDocPeticao.Shapes(wdDocPeticao.Shapes.Count).WrapFormat.Type = wdWrapTopBottom
            wdDocPeticao.Shapes(wdDocPeticao.Shapes.Count).RelativeHorizontalPosition = wdRelativeHorizontalPositionMargin
            wdDocPeticao.Shapes(wdDocPeticao.Shapes.Count).Left = wdShapeCenter
            wdDocPeticao.Shapes(wdDocPeticao.Shapes.Count).IncrementTop 7
            wdDocPeticao.Shapes(wdDocPeticao.Shapes.Count).ZOrder msoSendToBack
        End If
    End If
    
    ' Substituir as tabelas de média, se houver
    If btTabela <> 0 Then
        Set rngCont = Application.InputBox(Prompt:=DeterminarTratamento & ", selecione as células para a tabela de médias (12 linhas e 2 colunas, SEM cabeçalho).", Title:="Sísifo - Selecione a tabela de média!", Type:=8)
        rngCont.Copy
        
        For byteCont = 1 To btTabela
        
            With wdDocPeticao.Tables(byteCont)
                .Rows(3).Select
                wdDocPeticao.Activate
                appword.Selection.PasteAndFormat wdTableInsertAsRows 'wdUseDestinationStylesRecovery
                .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
                .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
                .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
                .Borders(wdBorderRight).LineStyle = wdLineStyleNone
                .Borders(wdBorderTop).LineStyle = wdLineStyleNone
                .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
                .Range.Style = "PPJCM Tabelas"
                .Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
                .Rows(1).Range.Bold = True
                .Rows(2).Range.Bold = True
                .Rows(.Rows.Count).Range.Bold = True
                .Rows.Alignment = wdAlignRowCenter
            End With
        Next byteCont
    End If
    
    
    ' Simplifica o nome para não ficar muito grande
    Select Case strCausaPedir
    Case "Desabastecimentos por período e causa determinados"
        strCausaPedir = "Desabastecimentos"
    Case "Cobrança de esgoto em imóvel não ligado à rede"
        strCausaPedir = "Esgoto sem rede"
    Case "Cobrança de esgoto com água cortada"
        strCausaPedir = "Esgoto com água cortada"
    Case "Classificação tarifa ou qtd. de economias"
        strCausaPedir = "Economias"
    Case "Suspeita de by-pass", "Multa por infração"
        strCausaPedir = "Gato"
    Case "Vaz. água ou extravas. esgoto com danos a patrimônio/morais"
        strCausaPedir = "Extravasamento"
    Case "Obra da Embasa com danos a patrimônio/morais"
        strCausaPedir = "Obra"
    Case "Acidente com pessoa/veículo em buraco"
        strCausaPedir = "Acidente em buraco"
    Case "Acidente com veículo (colisão ou atropelamento)"
        strCausaPedir = "Acidente com veículo"
    End Select
    
    ' Se tiver /, ajusta o nome (por exemplo, na CCR e outros desabastecimentos que têm data, pois o windows não aceita nome de arquivo com "/")
    strCausaPedir = Replace(strCausaPedir, "/", ".")
    
    ' Ajusta o nome pelo tipo de petição e pelo órgão em que tramita
    Select Case strPeticao
    Case "Contestação"
        If strOrgao = "JEC" Or strOrgao = "VC" Then strPeticao = "Contestacao"
        If strOrgao = "Procon" Then strPeticao = "Impugnacao"
        
    Case "RI/Apelação"
        Select Case strOrgao
        Case "JEC"
            strPeticao = "Recurso Inominado"
        
        Case "VC"
            strPeticao = "Apelacao"
        
        'Case "Procon"
        '    strPeticao = "Recurso Administrativo"
        
        End Select
    
    Case "Contrarrazões de RI/Apelação"
        Select Case strOrgao
        Case "JEC"
            strPeticao = "Contrarrazoes RI"
        
        Case "VC"
            strPeticao = "Contrarrazoes Apelacao"
        
        'Case "Procon"
        '    strPeticao = "Recurso Administrativo"
        
        End Select
        
    End Select
    
    ' Salvar o documento, ir para o início e exibir
    If bolSsfPrazosBotaoPdfPressionado Then 'Gerar como PDF
        wdDocPeticao.ExportAsFixedFormat OutputFilename:=BuscarCaminhoPrograma & "01 " & strPeticao & " " & strCausaPedir & " - " & SeparaPrimeirosNomes(ActiveSheet.Cells(ActiveCell.Row, 2).Formula, 2) & ".pdf", _
            ExportFormat:=wdExportFormatPDF, OptimizeFor:=wdExportOptimizeForOnScreen, CreateBookmarks:=wdExportCreateHeadingBookmarks, BitmapMissingFonts:=False
        MsgBox DeterminarTratamento & ", o PDF foi gerado com sucesso.", vbInformation + vbOKOnly, "Sísifo - Documento gerado"
        wdDocPeticao.Close wdDoNotSaveChanges
        If appword.Documents.Count = 0 Then appword.Quit
        
    Else ' Gerar como Word
        wdDocPeticao.SaveAs BuscarCaminhoPrograma & Format(Date, "yyyy.mm.dd") & " - " & strPeticao & " " & strCausaPedir & " - " & SeparaPrimeirosNomes(ActiveSheet.Cells(ActiveCell.Row, 2).Formula, 2) & ".docx"
        wdDocPeticao.GoTo wdGoToPage, wdGoToFirst
        'wdDocPeticao.Activate
        'appword.Activate
        MsgBox DeterminarTratamento & ", o documento Word foi gerado com sucesso.", vbInformation + vbOKOnly, "Sísifo - Documento gerado"
    End If
    
End Sub

Function ConfigurarFormulario(strCausaPedir As String, strOrgao As String, strPeticao As String) As Variant
''
'' Configura os formulários a serem exibidos.
''

Dim form As Variant
    
Select Case strCausaPedir
Case "Revisão de consumo elevado"
    Select Case strPeticao
    Case "Contestação"
        Set form = New frmContRevConsElevado
        With form
            .cmbAferHidrometro.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Aferiçãodehidrômetro").Address
            .cmbPadraoConsumo.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Padrãodeconsumo").Address
            .cmbVazInterno.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Vazamentointerno").Address
            If strOrgao = "VC" Then .chbRequererPericia.Visible = True
            If strOrgao = "Procon" Then
                .chbDanMorCorte.Value = False
                .chbDanMorCorte.Visible = False
                .chbDanoMoral.Value = False
                .chbDanoMoral.Visible = False
            End If
        End With
    
    Case "RI/Apelação"
        GoTo NãoFaz
        'Set form = New
        
    Case "Contrarrazões de RI/Apelação"
        GoTo NãoFaz
        'Set form = New
        
    End Select
    
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Devoluçãoemdobro").Address
    
Case "Corte no fornecimento"
    Select Case strPeticao
    Case "Contestação"
        Set form = New frmContCorte
        
    Case "RI/Apelação"
        GoTo NãoFaz
        'Set form = New frmRICorte
        
    Case "Contrarrazões de RI/Apelação"
        GoTo NãoFaz
        'Set form = New
        
    End Select
    
    form.cmbAvisoCorte.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("AvisoCorte").Address
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Devoluçãoemdobro").Address
    
Case "Realizar ligação de água"
    Select Case strPeticao
    Case "Contestação"
        Set form = New frmContFazerLigacao
    
    Case "RI/Apelação"
        GoTo NãoFaz
        'Set form = New
        
    Case "Contrarrazões de RI/Apelação"
        GoTo NãoFaz
        'Set form = New
        
    End Select
    
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Devoluçãoemdobro").Address
    
Case "Negativação no SPC"
    Select Case strPeticao
    Case "Contestação"
        Set form = New frmContNegativacao
        form.cmbAtitudeAutor.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("AtitudeAutorNegativação").Address
        form.cmbPerfilContrato.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("PerfilContrato").Address
    
    Case "RI/Apelação"
        GoTo NãoFaz
        'Set form = New
        
    Case "Contrarrazões de RI/Apelação"
        GoTo NãoFaz
        'Set form = New
        
    End Select
    
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Devoluçãoemdobro").Address
    
Case "Classificação tarifa ou qtd. de economias"
    Select Case strPeticao
    Case "Contestação"
        Set form = New frmContClasTarif
        form.cmbPorcentEsgoto.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("TarifaEsgotoPercent").Address
    
    Case "RI/Apelação"
        GoTo NãoFaz
        'Set form = New frmRIDesabGenerico
        
    Case "Contrarrazões de RI/Apelação"
        GoTo NãoFaz
        'Set form = New frmCRRIDesabGenerico
        
    End Select
    
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Devoluçãoemdobro").Address
    
Case "Suspeita de by-pass", "Multa por infração"
    Select Case strPeticao
    Case "Contestação"
        Set form = New frmContGato
    
    Case "RI/Apelação"
        GoTo NãoFaz
        'Set form = New frmRIDesabGenerico
        
    Case "Contrarrazões de RI/Apelação"
        GoTo NãoFaz
        'Set form = New frmCRRIDesabGenerico
        
    End Select
    
    form.cmbTipoGato.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("TipoGato").Address
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Devoluçãoemdobro").Address
    
Case "Débito de terceiro"
    Select Case strPeticao
    Case "Contestação"
        Set form = New frmContDebitoTerceiro
    
    Case "RI/Apelação"
        GoTo NãoFaz
        'Set form = New frmRIDesabGenerico
        
    Case "Contrarrazões de RI/Apelação"
        GoTo NãoFaz
        'Set form = New frmCRRIDesabGenerico
        
    End Select
    
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Devoluçãoemdobro").Address
    
Case "Cobrança de esgoto em imóvel não ligado à rede"
    Select Case strPeticao
    Case "Contestação"
        Set form = New frmContEsgotoSemRede
        
    Case "RI/Apelação"
        GoTo NãoFaz
        'Set form = New frmRIDesabGenerico
        
    Case "Contrarrazões de RI/Apelação"
        GoTo NãoFaz
        'Set form = New frmCRRIDesabGenerico
        
    End Select
        
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Devoluçãoemdobro").Address
    
Case "Cobrança de esgoto com água cortada"
    Select Case strPeticao
    Case "Contestação"
        Set form = New frmContEsgotoAguaCortada
        
    Case "RI/Apelação"
        GoTo NãoFaz
        'Set form = New frmRIDesabGenerico
        
    Case "Contrarrazões de RI/Apelação"
        GoTo NãoFaz
        'Set form = New frmCRRIDesabGenerico
        
    End Select
    
    form.cmbAtitudeAutor.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("AtitudeAutorEsgotoÁguaCortada").Address
    form.cmbProvaUsoImovel.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("ProvaUsoImóvel").Address
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Devoluçãoemdobro").Address
    
Case "Vaz. água ou extravas. esgoto com danos a patrimônio/morais"
    Select Case strPeticao
    Case "Contestação"
        Set form = New frmContRespCivil
        
    Case "RI/Apelação"
        GoTo NãoFaz
        'Set form = New frmRICorte
        
    Case "Contrarrazões de RI/Apelação"
        GoTo NãoFaz
        'Set form = New frmCRRIDesabGenerico
        
    End Select
    
    form.cmbOcorrencia.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Recorrência").Address
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Devoluçãoemdobro").Address
    
Case "Obra da Embasa com danos a patrimônio/morais"
    Select Case strPeticao
    Case "Contestação"
        Set form = New frmContRespCivil
        
    Case "RI/Apelação"
        GoTo NãoFaz
        'Set form = New frmRICorte
        
    Case "Contrarrazões de RI/Apelação"
        GoTo NãoFaz
        'Set form = New frmCRRIDesabGenerico
        
    End Select
    
    form.cmbOcorrencia.Enabled = False
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Devoluçãoemdobro").Address
    
Case "Acidente com pessoa/veículo em buraco"
    Select Case strPeticao
    Case "Contestação"
        Set form = New frmContRespCivil
        
    Case "RI/Apelação"
        GoTo NãoFaz
        'Set form = New frmRICorte
        
    Case "Contrarrazões de RI/Apelação"
        GoTo NãoFaz
        'Set form = New frmCRRIDesabGenerico
        
    End Select
    
    form.cmbOcorrencia.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("AcidenteBuraco").Address
    form.cmbOcorrencia.Text = "Veículo"
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Devoluçãoemdobro").Address
    
Case "Acidente com veículo (colisão ou atropelamento)"
    Select Case strPeticao
    Case "Contestação"
        Set form = New frmContRespCivil
        
    Case "RI/Apelação"
        GoTo NãoFaz
        'Set form = New frmRICorte
        
    Case "Contrarrazões de RI/Apelação"
        GoTo NãoFaz
        'Set form = New frmCRRIDesabGenerico
        
    End Select
    
    form.cmbOcorrencia.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("ColisaoAtropelamento").Address
    form.cmbOcorrencia.Text = "Veículo"
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Devoluçãoemdobro").Address
    
Case "Desabastecimentos por período e causa determinados"
    Select Case strPeticao
    Case "Contestação"
        Set form = New frmContDesabastecimento
        
    Case "RI/Apelação"
        Set form = New frmRIDesabastecimento
        form.cmbTestemunhas.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Testemunhas").Address
        
    Case "Contrarrazões de RI/Apelação"
        GoTo NãoFaz
        'Set form = New frmCRRIDesabGenerico
        'form.cmbTestemunhas.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Testemunhas").Address
        
    End Select
    
    With form
        .frmDesabastecimento.Caption = "Desabastecimentos por período e causa determinados - Genérico"
        .chbPrescricao.Caption = "Ajuizamento posterior a (prescrição trienal)"
        .chbSemFatura.Caption = "Autor não juntou conta do mês em questão"
        .cmbAlteracaoConsumo.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("AlteraçãoConsumoCCR").Address
        .cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Devoluçãoemdobro").Address
        .cmbCorresponsavel.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Corresponsáveis").Address
        .chbDanMorCorte.Value = False
        .chbDanMorCorte.Enabled = False
        .chbDanMorMeraCobranca.Value = False
        .chbDanMorMeraCobranca.Enabled = False
    End With
    
Case "Desabastecimento CCR 04/2015"
    Select Case strPeticao
    Case "Contestação"
        Set form = New frmContDesabastecimento
        
    Case "RI/Apelação"
        Set form = New frmRIDesabastecimento
        form.cmbTestemunhas.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Testemunhas").Address
        
    Case "Contrarrazões de RI/Apelação"
        GoTo NãoFaz
        'Set form = New frmCRRIDesabGenerico
        'form.cmbTestemunhas.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Testemunhas").Address
        
    End Select
    
    With form
        .frmDesabastecimento.Caption = "Desabastecimento - CCR - Abril/2015"
        .chbPrescricao.Caption = "Ajuizamento posterior a 02/04/2018"
        .chbSemFatura.Caption = "Autor não juntou conta do mês em questão (normalmente, Maio/2015)"
        .cmbAlteracaoConsumo.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("AlteraçãoConsumoCCR").Address
        .cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Devoluçãoemdobro").Address
        .cmbCorresponsavel.Text = "CCR"
        .cmbCorresponsavel.Enabled = False
        .txtDataInicio.Text = "01/04/2015"
        .txtDuracao.Text = "uma semana"
        .chbDanMorCorte.Value = False
        .chbDanMorCorte.Enabled = False
        .chbDanMorMeraCobranca.Value = False
        .chbDanMorMeraCobranca.Enabled = False
    End With
    
Case "Desabastecimento Uruguai 09/2016"
    Select Case strPeticao
    Case "Contestação"
        Set form = New frmContDesabastecimento
        
    Case "RI/Apelação"
        Set form = New frmRIDesabastecimento
        form.cmbTestemunhas.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Testemunhas").Address
        
    Case "Contrarrazões de RI/Apelação"
        GoTo NãoFaz
        'Set form = New frmCRRIDesabGenerico
        'form.cmbTestemunhas.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Testemunhas").Address
        
    End Select
    
    With form
        .frmDesabastecimento.Caption = "Desabastecimento - Uruguai, Mares - Set/2016"
        .chbPrescricao.Caption = "Ajuizamento posterior a 14/09/2019"
        .chbSemFatura.Caption = "Autor não juntou conta do mês em questão (normalmente, Out e Nov/2016)"
        .cmbAlteracaoConsumo.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("AlteraçãoConsumoCCR").Address
        .cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Devoluçãoemdobro").Address
        .cmbCorresponsavel.Text = ""
        .cmbCorresponsavel.Enabled = False
        .txtDataInicio.Text = "03/09/2016"
        .txtDuracao.Text = "quinze dias"
        .chbDanMorCorte.Value = False
        .chbDanMorCorte.Enabled = False
        .chbDanMorMeraCobranca.Value = False
        .chbDanMorMeraCobranca.Enabled = False
    End With
    
    
Case "Desabastecimento Liberdade 10/2017"
    Select Case strPeticao
    Case "Contestação"
        Set form = New frmContDesabastecimento
        
    Case "RI/Apelação"
        Set form = New frmRIDesabastecimento
        form.cmbTestemunhas.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Testemunhas").Address
        
    Case "Contrarrazões de RI/Apelação"
        GoTo NãoFaz
        'Set form = New frmCRRIDesabGenerico
        'form.cmbTestemunhas.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Testemunhas").Address
        
    End Select
    
    With form
        .frmDesabastecimento.Caption = "Desabastecimento - Liberdade, IAPI, Pero Vaz, Curuzu, Santa Mônica - Out/2017"
        .chbPrescricao.Caption = "Ajuizamento posterior a 02/11/2020"
        .chbSemFatura.Caption = "Autor não juntou conta do mês em questão (normalmente, Dezembro/2017)"
        .cmbAlteracaoConsumo.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("AlteraçãoConsumoCCR").Address
        .cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Devoluçãoemdobro").Address
        .cmbCorresponsavel.Text = ""
        .cmbCorresponsavel.Enabled = False
        .txtDataInicio.Text = "30/10/2017"
        .txtDuracao.Text = "quinze dias"
        .chbDanMorCorte.Value = False
        .chbDanMorCorte.Enabled = False
        .chbDanMorMeraCobranca.Value = False
        .chbDanMorMeraCobranca.Enabled = False
    End With
    
    
Case "Desabastecimento Apagão Xingu 03/2018"
    Select Case strPeticao
    Case "Contestação"
        Set form = New frmContDesabastecimento
        
    Case "RI/Apelação"
        Set form = New frmRIDesabastecimento
        form.cmbTestemunhas.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Testemunhas").Address
        
    Case "Contrarrazões de RI/Apelação"
        GoTo NãoFaz
        'Set form = New frmCRRIDesabGenerico
        'form.cmbTestemunhas.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Testemunhas").Address
        
    End Select
    
    With form
        .frmDesabastecimento.Caption = "Desabastecimento - Apagão Xingu 03/2018"
        .chbPrescricao.Caption = "Ajuizamento posterior a 21/03/2021"
        .chbSemFatura.Caption = "Autor não juntou conta do mês em questão (normalmente, Abr ou Mai/2018)"
        .cmbAlteracaoConsumo.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("AlteraçãoConsumoCCR").Address
        .cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Devoluçãoemdobro").Address
        .cmbCorresponsavel.Text = "Operador Nacional do Sistema Elétrico - ONS"
        .cmbCorresponsavel.Enabled = False
        .txtDataInicio.Text = "21/03/2018"
        .txtDuracao.Text = "10 dias"
        .chbDanMorCorte.Value = False
        .chbDanMorCorte.Enabled = False
        .chbDanMorMeraCobranca.Value = False
        .chbDanMorMeraCobranca.Enabled = False
    End With
    
    
Case "Fixo de esgoto"
    Select Case strPeticao
    Case "Contestação"
        Set form = New frmContFixoEsgoto
        
    Case "RI/Apelação"
        GoTo NãoFaz
        
    Case "Contrarrazões de RI/Apelação"
        GoTo NãoFaz
        
    End Select
    
    With form
        .cmbConfessaPoco.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("PosturaAutorPoço").Address
        .cmbVolumeParadigma.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("VolumeParadigmaEsgoto").Address
        .cmbPorcentEsgoto.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("TarifaEsgotoPercent").Address
        .cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Devoluçãoemdobro").Address
        .cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Devoluçãoemdobro").Address
    End With
    
Case "Outros vícios"
    Select Case strPeticao
    Case "Contestação"
        Set form = New frmContVicio
        
    Case "RI/Apelação"
        GoTo NãoFaz
        
    Case "Contrarrazões de RI/Apelação"
        GoTo NãoFaz
        
    End Select
    
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigurações'!" & ThisWorkbook.Sheets("cfConfigurações").Range("Devoluçãoemdobro").Address
    
Case Else
NãoFaz:
    MsgBox "Sinto muito, " & DeterminarTratamento & "! Eu ainda não sei fazer " & strPeticao & " de " & strCausaPedir & ". Aguarde uma nova versão -- eu juro que vou tentar aprender!", vbInformation + vbOKOnly, "Sísifo em treinamento"
    Set ConfigurarFormulario = Nothing
    Exit Function
    
End Select

Set ConfigurarFormulario = form

End Function

Function ObterVariaveis(strVariavel As String, strCausaPedir As String, form As Variant, Optional varCont As Variant)
''
'' Retorna o valor de uma variável, conforme a strCausaPedir.
''
Dim X As Variant
Dim btCont As Byte
Dim strCont As String
Dim fatImpugnada As Fatura, fatPretendida As Fatura
        

' Variáveis de mais de uma causa de pedir
Select Case strVariavel
    Case "data-audiência-conciliação"
        X = Trim(Application.InputBox(DeterminarTratamento & ", informe a data da audiência de conciliação", "Sísifo - Dados de tempestividade", varCont, Type:=2))
    Case "termo-decadência"
        X = CDate(InputBox(DeterminarTratamento & ", qual foi a data da reclamação administrativa ou ajuizamento da ação?", "Informações para decadência", Format(Date - 30, "dd/mm/yyyy"))) - 30
        GoTo AtribuiVariavel
    Case "pediu-devolução"
        X = IIf(form.chbDevolDobro.Value = True, ", com devolução em dobro dos valores pagos de forma alegadamente indevida no período impugnado", "")
        GoTo AtribuiVariavel
    Case "pediu-danos-materiais"
        X = IIf(form.chbDanMat.Value = True, " e indenização por danos materiais", "")
        GoTo AtribuiVariavel
    Case "pediu-danos-morais"
        X = IIf(form.chbDanoMoral.Value = True, ", além de indenização por danos morais", "")
        GoTo AtribuiVariavel
    Case "comarca-competente"
        X = Trim(form.txtComarcaCompetente.Value)
        GoTo AtribuiVariavel
    Case "mês-inicial"
        X = InputBox(DeterminarTratamento & ", informe o mês de referência da fatura inicial do período impugnado pela parte Adversa", "Informações sobre pedido", "abril/2019")
        GoTo AtribuiVariavel
    Case "resumo-da-má-fé"
        X = InputBox(DeterminarTratamento & ", faça um resumo da conduta da parte Adversa que constitui má-fé, completando a seguinte frase:" & vbCrLf & "A parte Autora ...", "Informações sobre má-fé", "sonegou a informação de que esta empresa verificara o vazamento e a cientificara previamente")
        GoTo AtribuiVariavel
End Select

Select Case strCausaPedir
Case "Revisão de consumo elevado"
    ' Variáveis de consumo elevado
    Select Case strVariavel
    Case "período"
        X = InputBox("Informe o período impugnado:", "Informe o período", "de Setembro a Novembro/2019")
    Case "mês-início-medições"
        X = InputBox("Informe o mês de referência da fatura em que se iniciaram as medições de consumo por hidrômetro:", "Informações sobre consumo", "08/2019")
    Case "média-real"
        X = Trim(Application.InputBox(DeterminarTratamento & ", informe a média de consumo real dos 12 meses anteriores ao período impugnado, em m3:", "Informe a média real", Type:=2))
    Case "média-afirmada"
        X = InputBox(DeterminarTratamento & ", informe o valor que a parte alega ser sua média de consumo, em m3:", "Informe a média alegada", "10")
    Case "tempo-sem-medição"
        X = InputBox(DeterminarTratamento & ", informe o tempo que o imóvel passou sem medições reais de consumo antes de ser instalado hidrômetro, completando a frase abaixo:" & vbCrLf & """Há ..., a parte Demandante não tem consumo mensal no valor que afirma ser sua média.""", "Informações sobre consumo", "anos")
    Case "consumo-impugnado"
        X = Trim(Application.InputBox(DeterminarTratamento & ", informe o valor do(s) consumo(s) impugnado(s), em m3:", "Informe o consumo impugnado", "13,  15", Type:=2))
    Case "consumo-fictício-utilizado"
        X = Trim(Application.InputBox(DeterminarTratamento & ", informe o valor do(s) consumo(s) fictício(s) utilizados antes da instalação do hidrômetro, em m³:", "Informações sobre consumo", "06", Type:=2))
    Case "número-hidrômetro"
        X = Trim(Application.InputBox(DeterminarTratamento & ", informe o número do hidrômetro", "Informe o número do hidrômetro", Type:=2))
    Case "substituição-hidrômetro"
        X = Trim(Application.InputBox(DeterminarTratamento & ", informe a data em que o hidrômetro foi instalado ou substituído", "Informe data da substituição", Type:=2))
    Case "números-processos"
        X = InputBox(DeterminarTratamento & ", informe os números dos processos anteriores que resultaram na média viciada", "Informe números de processos anteriores", "XX, YY e ZZ")
    Case "tempo-medições-fictícias"
        X = InputBox(DeterminarTratamento & ", informe por quanto tempo, aproximadamente, o Autor pagou por mínimo ou média de consumo", "Informe tempo pagando média", "cerca de um ano")
    Case "médias-estabelecidas-judicialmente"
        X = InputBox(DeterminarTratamento & ", informe o valor das médias estabelecidas judicialmente nos processos anteriores", "Informe valor das médias judiciais", "10 ou 13")
    Case "medição-maior"
        X = InputBox(DeterminarTratamento & ", informe o valor da maior medição nos últimos tempos", "Informe valor mais alto")
    Case "quantidade-habitantes"
        X = InputBox(DeterminarTratamento & ", informe a quantidade de habitantes do imóvel, segundo informações dos autos", "Quantidade de habitantes", "03")
        btCont = CByte(X)
    Case "projeção-média"
        X = btCont * 5.4
    Case "exemplos-consumo"
        X = Trim(Application.InputBox(DeterminarTratamento & ", qual é a faixa de consumo do período impugnado, para comparação com a quantidade de habitantes?", "Informe a faixa de consumo impugnada em m3", "10 a 13", Type:=2))
    Case "devolução-em-dobro"
        X = IIf(form.chbDevolDobro.Value = True, " Pede a devolução em dobro dos valores pagos em excesso.", "")
    Case "dano-moral"
        X = IIf(form.chbDanoMoral.Value = True, " Requer, ainda, indenização por danos morais.", "")
    End Select

Case "Corte no fornecimento"
    ' Variáveis de corte
    Select Case strVariavel
    Case "data-corte"
        X = InputBox(DeterminarTratamento & ", informe a data do corte:", "Informações da síntese dos fatos", Format(Date - 60, "dd/mm/yyyy"))
    Case "conta-razão-corte"
        X = InputBox(DeterminarTratamento & ", informe o período que motivou o corte:", "Informações da síntese dos fatos", "de Setembro e Outubro/2017")
    Case "data-pagamento-alegado"
        X = InputBox(DeterminarTratamento & ", informe a data em que o Autor alega haver pago:", "Informações da síntese dos fatos", Format(Date - 15, "dd/mm/yyyy"))
    Case "vencimentos-débitos"
        X = InputBox(DeterminarTratamento & ", digite as datas de vencimentos das faturas que estavam em aberto", "Informações sobre existência de débito", "11/05/2017, 11/06/2017")
    Case "maior-período-atraso"
        X = InputBox(DeterminarTratamento & ", qual foi o maior período que a parte Autora atrasou os pagamentos na época dos fatos?", "Informações sobre existência de débito", "quase três meses")
    Case "quantas-equipes-corte"
        X = InputBox(DeterminarTratamento & ", quantas SSs de corte foram iniciadas no histórico da parte Autora?", "Informações sobre inadimplência contumaz", "cinco")
    Case "atraso-pagamento-véspera"
        X = InputBox(DeterminarTratamento & ", o pagamento feito na véspera do corte ocorreu com quanto tempo de atraso?", "Informações sobre pagamento na véspera", "um mês e meio")
    Case "duração-do-corte"
        X = InputBox(DeterminarTratamento & ", qual foi a duração do corte?", "Informações sobre o corte de curta duração", "algumas poucas horas")
    End Select

Case "Negativação no SPC"
    ' Variáveis de Negativação
    Select Case strVariavel
    Case "matrícula"
        X = Trim(form.txtMatricula1)
    Case "vencim1"
        X = Trim(form.txtVencim1)
    Case "val1"
        X = Trim(form.txtVal1)
    Case "aaaamm1"
        X = Trim(form.txtaaaamm1)
    Case "vencim2"
        X = Trim(form.txtVencim2)
    Case "val2"
        X = Trim(form.txtVal2)
    Case "aaaamm2"
        X = Trim(form.txtaaaamm2)
    Case "vencim3"
        X = Trim(form.txtVencim3)
    Case "val3"
        X = Trim(form.txtVal3)
    Case "aaaamm3"
        X = Trim(form.txtaaaamm3)
    'Case "data-negativação"
    '    X = Trim(Application.InputBox(DeterminarTratamento & ", qual é a data da negativação realizada pela Embasa?", "Informe data da negativação", Type:=2))
    Case "mês-final-do-uso-regular"
        X = Trim(form.txtMesFinalUsoRegular.Text)
    Case "empresas-negativadoras-anteriores"
        X = Trim(InputBox(DeterminarTratamento & ", quais foram as empresas que negativaram a parte Autora ANTES da negativação realizada pela Embasa?", "Informe as empresas negativadoras anteriores"))
    Case "data-parcelamento"
        X = CDate(Trim(Application.InputBox(DeterminarTratamento & ", qual é a data da realização do parcelamento?", "Informe data do parcelamento", Type:=2)))
    Case "período-parcelado"
        X = Trim(InputBox(DeterminarTratamento & ", quais foram os meses das contas parceladas?", "Informe o período abrangido pelo parcelamento", "01/2016 até 06/2017"))
    Case "números-processos-negativação"
        X = Trim(InputBox(DeterminarTratamento & ", quais foram os números dos outros processos sobre as outras linhas da mesma negativação?", "Informe os outros processos"))
    Case "endereço-matrícula"
        X = Trim(InputBox(DeterminarTratamento & ", qual é o endereço da ligação registrado no SCI?", "Informe o endereço da matrícula"))
    End Select

Case "Realizar ligação de água", "Desmembramento de ligações"
    ' Variáveis de Realizar ligação
    Select Case strVariavel
    Case "data-solicitação"
        X = CDate(Trim(Application.InputBox(DeterminarTratamento & ", informe a data da solicitação de ligação:", "Informe data da solicitação", Type:=2)))
    Case "número-SS"
        X = Trim(Application.InputBox(DeterminarTratamento & ", informe o número da Solicitação de Serviço de ligação:", "Informe SS", Type:=2))
    Case "data-ligação"
        X = CDate(Trim(Application.InputBox(DeterminarTratamento & ", informe a data da ligação:", "Informe data da ligação", Type:=2)))
    Case "exigencia-alegada"
        If form.chbSemReservatorioBomba.Value = True Then
            X = " (instalação de reservatório ao nível do solo com bomba)"
        ElseIf form.chbSepararInstalacoesInternas.Value = True Then
            X = " (separação das instalações hidráulicas)"
        ElseIf form.chbAltitudeInsuficiente.Value = True Then
            X = " (distância superior à regulamentar da extremidade da rede e altitude demasiada do seu terreno)"
        End If
    Case "alega-comparação-com-vizinhos"
        X = IIf(form.chbComparacaoVizinhos.Value = True, " Afirma equivocadamente que tal tratamento seria uma arbitrariedade contra sua pessoa, e que não foi exigido de nenhum de seus vizinhos.", "")
    Case "defesa-comparação-com-vizinhos"
        X = IIf(form.chbComparacaoVizinhos.Value = True, "3) Não houve tratamento diferenciado entre Demandante e vizinhos. Se houve alguma exceção em imóvel vizinho - no que não acreditamos, e não há provas nos autos -, certamente o foi por ordem judicial ou anterior à vigência do regulamento.", "")
    Case "defesa-comparação-com-vizinhos2"
        X = IIf(form.chbComparacaoVizinhos.Value = True, " Ademais, eventual erro ocorrido com vizinho não gera para a parte Autora direito subjetivo a um erro similar! Esse pensamento é absurdo: um erro é um erro; ele deve ser corrigido, e não gera direito a novos erros!", "")
    Case "pedidos-comparação-com-vizinhos"
        X = IIf(form.chbComparacaoVizinhos.Value = True, "; bem como por não ter havido discriminação para com vizinhos - e mesmo que, por lapso, um vizinho tenha conseguido ligação de água infringindo as normas técnicas, um erro anterior deve ser consertado, e não gera direito a novos erros", "")
    End Select

Case "Cobrança de esgoto em imóvel não ligado à rede"
    ' Variáveis de Esgoto sem ligação à rede
    Select Case strVariavel
    Case "motivo-dano-moral"
        X = "cobrança de tarifa de esgoto em período no qual alega não ter utilizado o serviço"
    Case "diâmetro-rede"
        X = InputBox(DeterminarTratamento & ", informe o diâmetro da rede que serve o imóvel da parte Autora", "Informações sobre cobrança de esgoto sem ligação à rede", "150 mm")
    Case "data-implantação-esgoto"
        X = Trim(form.txtDataImplantacao.Text)
    Case "mês-implantação-esgoto"
        X = Trim(Format(form.txtDataImplantacao, "mmmm/yyyy"))
    Case "quantidade-ligações-no-mês"
        X = InputBox(DeterminarTratamento & ", informe quantas ligações foram realizadas pela unidade no mesmo mês que a do imóvel do Autor", "Informações sobre cobrança de esgoto sem ligação à rede", "30")
    Case "escritório-local"
        X = InputBox(DeterminarTratamento & ", informe qual o escritório local do imóvel da parte Autora, conforme agrupamento no gráfico de ligações de esgoto", "Informações sobre cobrança de esgoto sem ligação à rede", "Escritório de Serviços de Itapuã")
    End Select

Case "Cobrança de esgoto com água cortada"
    ' Variáveis de Esgoto sem ligação à rede
    Select Case strVariavel
    Case "água-cortada"
        X = IIf(form.cmbAtitudeAutor.Value = "Água estava cortada, não pode ser cobrado esgoto", "seu imóvel está com o abastecimento de água cortado, mas recebe cobranças do serviço de coleta de esgoto", "")
    Case "imóvel-desabitado"
        X = IIf(form.cmbAtitudeAutor.Value = "Imóvel estava desabitado no período impugnado", "seu imóvel está desabitado, sem uso do serviço de abastecimento, mas recebe cobranças do serviço de coleta de esgoto", "")
    Case "solicitou-suspensão"
        X = IIf(form.cmbAtitudeAutor.Value = "Afirma apenas que solicitou cancelamento", "solicitou suspensão do abastecimento de água de seu imóvel, mas recebe cobranças do serviço de coleta de esgoto", "")
    Case "motivo-dano-moral"
        X = "cobrança de tarifa de esgoto em período no qual não houve abastecimento de água"
    Case "pede-devolução-em-dobro"
        X = IIf(form.chbDevolDobro.Value = True, ", a devolução em dobro dos valores pagos", "")
    End Select

Case "Classificação tarifa ou qtd. de economias"
    ' Variáveis de Classificação tarifária
    Select Case strVariavel
    Case "motivo-dano-moral"
        X = "cobrança de tarifa em desconformidade com a classificação do imóvel"
    Case "mês-final"
        X = Trim(InputBox(DeterminarTratamento & ", informe o mês de referência da fatura final do período pleiteado pela parte Adversa", "Informações sobre pedido", Trim(form.txtRefAlteracao2.Text)))
    Case "classificação-original"
        X = Trim(form.txtClassifOriginal.Text)
    Case "classificação-original-sem-ponto"
        X = CodClassificacaoFatura(Trim(form.txtClassifOriginal.Text))
    Case "categoria-original"
        X = CategoriaExtenso(Trim(form.txtClassifOriginal.Text))
    Case "qtd-economias-original"
        X = EconomiasExtenso(Trim(form.txtClassifOriginal.Text))
    Case "data-segunda-classificação"
        X = Trim(form.txtDataAlteracao1.Text)
    Case "classificação-segunda"
        X = Trim(form.txtClassifAlteracao1.Text)
    Case "classificação-segunda-sem-ponto"
        X = CodClassificacaoFatura(Trim(form.txtClassifAlteracao1.Text))
    Case "categoria-segunda"
        X = CategoriaExtenso(Trim(form.txtClassifAlteracao1.Text))
    Case "qtd-economias-segunda"
        X = EconomiasExtenso(Trim(form.txtClassifAlteracao1.Text))
    Case "referência-segunda-classificação"
        X = Trim(form.txtRefAlteracao1.Text)
    Case "motivo-terceira-classificação"
        X = InputBox(DeterminarTratamento & ", informe o motivo da segunda reclassificação, completando a frase abaixo" & vbCrLf & vbCrLf & "fiscalização na qual se percebeu que...", "Informações sobre classificação tarifária", "o imóvel efetivamente tinha composição diferente")
    Case "data-terceira-classificação"
        X = Trim(form.txtDataAlteracao2.Text)
    Case "classificação-terceira"
        X = Trim(form.txtClassifAlteracao2.Text)
    Case "classificação-terceira-sem-ponto"
        X = CodClassificacaoFatura(Trim(form.txtClassifAlteracao2.Text))
    Case "categoria-terceira"
        X = CategoriaExtenso(Trim(form.txtClassifAlteracao2.Text))
    Case "qtd-economias-terceira"
        X = EconomiasExtenso(Trim(form.txtClassifAlteracao2.Text))
    Case "referência-terceira-classificação"
        X = Trim(form.txtRefAlteracao2.Text)
    Case "composição-alegada-embasa-extenso"
        X = InputBox(DeterminarTratamento & ", informe a classificação tarifária defendida pela Embasa, por extenso -- lembrando que o Manual Comercial da Embasa permite que o consumidor opte por ser faturado como uma única economia", "Informações sobre classificação pretendida", ClassificacaoExtenso(Trim(form.txtClassifAlteracao1.Text)))
    Case "classificação-pretendida-extenso"
        X = InputBox(DeterminarTratamento & ", informe a classificação tarifária pretendida pela parte Autora, por extenso", "Informações sobre classificação pretendida", ClassificacaoExtenso(Trim(form.txtExPretCat.Text) & "." & Trim(form.txtExPretEconomias.Text)))
    Case "mês-inicial-contraposto"
        X = InputBox(DeterminarTratamento & ", informe o mês de referência da fatura inicial do período calculado a cobrar da parte Adversa", "Informações sobre reconvenção ou pedido contraposto", Trim(form.txtRefAlteracao1.Text))
    Case "mês-final-contraposto"
        X = InputBox(DeterminarTratamento & ", informe o mês de referência da fatura final do período calculado a cobrar da parte Adversa", "Informações sobre reconvenção ou pedido contraposto", Trim(form.txtRefAlteracao2.Text))
    Case "valor-diferença"
        X = InputBox(DeterminarTratamento & ", informe o valor da diferença total a ser demandado da parte Autora", "Informações sobre reconvenção ou pedido contraposto")
        X = Format(X, "#,##0.00")
    Case "inexistência-de-solicitação"
        X = IIf(form.chbRetroativoSemSolicitacao.Value = True, ", e nunca houve solicitação do consumidor para a mudança", "")
    Case "pede-reclassificação-retroativa"
        X = IIf(form.chbRetroativoSemSolicitacao.Value = True, "especialmente o de reclassificação retroativa, ", "")
    End Select
    
    If form.chbApresentarCalcExemplo.Value = True Then
        ''Configura fatura e faz os cálculos
        Set fatImpugnada = New Fatura
        Set fatPretendida = New Fatura
        
        If form.chbTemDoisTipos.Value = False Then '1 categoria só
            fatImpugnada.BuscarEstruturaTarifaria form.txtMesRefExemplo.Text, form.txtConsExemplo.Text, form.cmbPorcentEsgoto.Text, form.txtExFaturado1Cat.Text, form.txtExFaturado1Economias.Text
        Else '2 categorias
            fatImpugnada.BuscarEstruturaTarifaria form.txtMesRefExemplo.Text, form.txtConsExemplo.Text, form.cmbPorcentEsgoto.Text, form.txtExFaturado1Cat.Text, form.txtExFaturado1Economias.Text, form.txtExFaturado2Cat.Text, form.txtExFaturado2Economias.Text
        End If
        
        fatPretendida.BuscarEstruturaTarifaria form.txtMesRefExemplo.Text, form.txtConsExemplo.Text, form.cmbPorcentEsgoto.Text, form.txtExPretCat.Text, form.txtExPretEconomias.Text
        
        fatImpugnada.CalcularTotal
        fatPretendida.CalcularTotal
        
        ''Atribui as variáveis (valores de faixas primeiro, porque são eliminadas mais fácil, enquanto o Select testa uma por uma).
        If (Left(strVariavel, 5) = "faixa" And IsNumeric(Mid(strVariavel, 6, 1)) And Mid(strVariavel, 7, 1) = "-") _
        Or (Left(strVariavel, 9) = "consfaixa" And IsNumeric(Mid(strVariavel, 10, 1)) And Mid(strVariavel, 11, 1) = "-") _
        Or (Left(strVariavel, 8) = "tarfaixa" And IsNumeric(Mid(strVariavel, 9, 1)) And Mid(strVariavel, 10, 1) = "-") _
        Or (Left(strVariavel, 9) = "subtfaixa" And IsNumeric(Mid(strVariavel, 10, 1)) And Mid(strVariavel, 11, 1) = "-") Then  ' Se for faixa...
        
            For btCont = 1 To 9 Step 1
                If Right(strVariavel, 5) = "-imp1" And fatImpugnada.Categorias(1).Faixas.Count >= btCont Then
                    With fatImpugnada.Categorias(1).Faixas(btCont)
                        Select Case strVariavel
                        Case "faixa" & btCont & "-imp1"
                            X = .AbrangenciaFaixa
                            Exit For
                        Case "consfaixa" & btCont & "-imp1"
                            X = .ConsumoNaFaixa
                            Exit For
                        Case "tarfaixa" & btCont & "-imp1"
                            X = Format(.Tarifa, "#,##0.00") & IIf(.TipoTarifa = "fixo", "", "/m3")
                            Exit For
                        Case "subtfaixa" & btCont & "-imp1"
                            X = .SubTotal
                            X = Format(X, "#,##0.00")
                            Exit For
                        End Select
                    End With
                End If
            
                If Right(strVariavel, 5) = "-imp2" Then
                    If fatImpugnada.Categorias(2).Faixas.Count >= btCont Then
                        With fatImpugnada.Categorias(2).Faixas(btCont)
                            Select Case strVariavel
                            Case "faixa" & btCont & "-imp2"
                                X = .AbrangenciaFaixa
                                Exit For
                            Case "consfaixa" & btCont & "-imp2"
                                X = .ConsumoNaFaixa
                                Exit For
                            Case "tarfaixa" & btCont & "-imp2"
                                X = Format(.Tarifa, "#,##0.00") & IIf(.TipoTarifa = "fixo", "", "/m3")
                                Exit For
                            Case "subtfaixa" & btCont & "-imp2"
                                X = .SubTotal
                                X = Format(X, "#,##0.00")
                                Exit For
                            End Select
                        End With
                    End If
                End If
            
                If Right(strVariavel, 6) = "-pret1" And fatPretendida.Categorias(1).Faixas.Count >= btCont Then
                    With fatPretendida.Categorias(1).Faixas(btCont)
                        Select Case strVariavel
                        Case "faixa" & btCont & "-pret1"
                            X = .AbrangenciaFaixa
                            Exit For
                        Case "consfaixa" & btCont & "-pret1"
                            X = .ConsumoNaFaixa
                            Exit For
                        Case "tarfaixa" & btCont & "-pret1"
                            X = Format(.Tarifa, "#,##0.00") & IIf(.TipoTarifa = "fixo", "", "/m3")
                            Exit For
                        Case "subtfaixa" & btCont & "-pret1"
                            X = .SubTotal
                            X = Format(X, "#,##0.00")
                            Exit For
                        End Select
                    End With
                End If
            Next btCont
        End If
        
        ' Demais variáveis que não são valores de faixas
        Select Case strVariavel
        Case "referência-conta-exemplo"
            X = fatImpugnada.MesReferencia
        Case "consumo-conta-exemplo"
            X = fatImpugnada.ConsumoTotal
        Case "consumo-por-economia-imp"
            X = fatImpugnada.ConsumoPorEconomia
        Case "total-água-imp"
            X = fatImpugnada.TotalAgua
            X = Format(X, "#,##0.00")
        Case "porcentagem-esgoto"
            X = fatImpugnada.PorcentEsgoto
        Case "total-água-esgoto-imp"
            X = fatImpugnada.TotalAguaEsgoto
            X = Format(X, "#,##0.00")
        Case "qt-economias-imp"
            X = fatImpugnada.QtdTotalEconomias
        Case "total-esgoto-imp"
            X = fatImpugnada.TotalEsgoto
            X = Format(X, "#,##0.00")
        Case "qt-economias-imp1"
            X = fatImpugnada.Categorias(1).QtdEconomias
        Case "classificação-imp1"
            X = CategoriaExtenso(fatImpugnada.Categorias(1).Categoria)
        Case "água-por-econ-imp1"
            X = fatImpugnada.Categorias(1).AguaPorEconomia
            X = Format(X, "#,##0.00")
        Case "subt-água-imp1"
            X = fatImpugnada.Categorias(1).SubtotalAgua
            X = Format(X, "#,##0.00")
        Case "qt-economias-imp2"
            X = fatImpugnada.Categorias(2).QtdEconomias
        Case "classificação-imp2"
            X = CategoriaExtenso(fatImpugnada.Categorias(2).Categoria)
        Case "água-por-econ-imp2"
            X = fatImpugnada.Categorias(2).AguaPorEconomia
            X = Format(X, "#,##0.00")
        Case "subt-água-imp2"
            X = fatImpugnada.Categorias(2).SubtotalAgua
            X = Format(X, "#,##0.00")
        Case "classificação-pretendida"
            X = ClassificacaoExtenso(form.txtExPretCat.Text & "." & form.txtExPretEconomias.Text)
        Case "consumo-por-economia-pret"
            X = fatPretendida.ConsumoPorEconomia
        Case "total-água-pret"
            X = fatPretendida.TotalAgua
            X = Format(X, "#,##0.00")
        Case "porcentagem-esgoto"
            X = fatPretendida.PorcentEsgoto
        Case "total-água-esgoto-pret"
            X = fatPretendida.TotalAguaEsgoto
            X = Format(X, "#,##0.00")
        Case "qt-economias-pret"
            X = fatPretendida.QtdTotalEconomias
        Case "total-esgoto-pret"
            X = fatPretendida.TotalEsgoto
            X = Format(X, "#,##0.00")
        Case "qt-economias-pret1"
            X = fatPretendida.Categorias(1).QtdEconomias
        Case "classificação-pret1"
            X = CategoriaExtenso(fatPretendida.Categorias(1).Categoria)
        Case "água-por-econ-pret1"
            X = fatPretendida.Categorias(1).AguaPorEconomia
            X = Format(X, "#,##0.00")
        Case "subt-água-pret1"
            X = fatPretendida.Categorias(1).SubtotalAgua
            X = Format(X, "#,##0.00")
        End Select
            
    End If

Case "Suspeita de by-pass"
    'Variáveis de gato
    Select Case strVariavel
    Case "nega-autoria"
        X = IIf(form.optNegaAutoria.Value = True, "; alega que não foi o responsável pelo ilícito", "")
    Case "nega-gato"
        X = IIf(form.optNegaExistencia.Value = True, "; alega que não houve gato", "")
    Case "fez-reclamações-não-atendidas"
        X = IIf(form.optReclamacoesNaoAtendidas.Value = True, "; alega que realizou reclamações, as quais não foram atendidas", "")
    Case "pede-devolução-em-dobro"
        X = IIf(form.chbDevolDobro.Value = True, " e que os valores pagos sejam devolvidos em dobro", "")
    Case "data-fiscalização"
        X = Trim(form.txtDataRetiradaGato.Text)
    Case "fatura-regularização-consumo"
        X = Trim(form.txtMesRefRegulaConsumo.Text)
    Case "fatura-multa"
        X = Trim(form.txtMesRefMulta.Text)
    Case "total-sanções"
        X = Trim(form.txtTotalSancoes.Text)
    Case "valor-multa"
        X = IIf(Trim(form.txtValorMulta.Text) <> "", "; multa sancionatória pela prática do ilícito, no importe de R$ " & form.txtValorMulta.Text, "")
    Case "valor-recuperação-consumo"
        X = IIf(Trim(form.txtValorRecCons.Text) <> "", "; recuperação de consumo pela fruição do ilícito, calculada pela média de consumo anterior ao período do ilícito, no importe de R$ " & form.txtValorRecCons.Text, "")
    Case "valor-custos-reparo"
        X = IIf(Trim(form.txtValorReparo.Text) <> "", "; ressarcimento dos custos pela reparação do ilícito, no importe de R$ " & form.txtValorReparo.Text, "")
    
    End Select

Case "Débito de terceiro"
    'Variáveis de gato
    Select Case strVariavel
    Case "motivo-dano-moral"
        X = "cobrança contra si de débito de período em que alega que o serviço era utilizado por outro usuário"
    
    End Select

Case "Fixo de esgoto"
    'Variáveis de gato
    Select Case strVariavel
    Case "motivo-dano-moral"
        X = "cobrança de tarifa de esgoto em período no qual alega não ter utilizado o serviço"
    Case "confessa-hidrômetro"
        X = IIf(form.cmbConfessaPoco.Value = "Confessa que foi instalado hidrômetro no poço", " Afirma que a Embasa instalou hidrômetro no seu poço artesiano para medição do consumo deste, a fim de mensurar o valor da cobrança de esgoto.", "")
    Case "tipo-estabelecimento"
        X = Trim(InputBox(DeterminarTratamento & ", informe o tipo de estabelecimento da parte Adversa", "Informações sobre pedido", "um salão de beleza"))
        
    End Select
    
    ''Dimensiona objetos, configura fatura e faz os cálculos
    Set fatImpugnada = New Fatura
    
    If form.chbTemDoisTipos.Value = False Then '1 categoria só
        fatImpugnada.BuscarEstruturaTarifaria form.txtMesRefExemplo.Text, form.txtConsExemplo.Text, form.cmbPorcentEsgoto.Text, form.txtExCat1.Text, form.txtExEconomias1.Text
    Else '2 categorias
        fatImpugnada.BuscarEstruturaTarifaria form.txtMesRefExemplo.Text, form.txtConsExemplo.Text, form.cmbPorcentEsgoto.Text, form.txtExCat1.Text, form.txtExEconomias1.Text, form.txtExCat2.Text, form.txtExEconomias2.Text
    End If
    
    fatImpugnada.CalcularTotal
    
    ''Atribui as variáveis (valores de faixas primeiro, porque são eliminadas mais fácil, enquanto o Select testa uma por uma).
    If (Left(strVariavel, 5) = "faixa" And IsNumeric(Mid(strVariavel, 6, 1)) And Mid(strVariavel, 7, 1) = "-") _
    Or (Left(strVariavel, 9) = "consfaixa" And IsNumeric(Mid(strVariavel, 10, 1)) And Mid(strVariavel, 11, 1) = "-") _
    Or (Left(strVariavel, 8) = "tarfaixa" And IsNumeric(Mid(strVariavel, 9, 1)) And Mid(strVariavel, 10, 1) = "-") _
    Or (Left(strVariavel, 9) = "subtfaixa" And IsNumeric(Mid(strVariavel, 10, 1)) And Mid(strVariavel, 11, 1) = "-") Then  ' Se for faixa...
    
        For btCont = 1 To 9 Step 1
            If Right(strVariavel, 2) = "-1" And fatImpugnada.Categorias(1).Faixas.Count >= btCont Then
                With fatImpugnada.Categorias(1).Faixas(btCont)
                    Select Case strVariavel
                    Case "faixa" & btCont & "-1"
                        X = .AbrangenciaFaixa
                        Exit For
                    Case "consfaixa" & btCont & "-1"
                        X = .ConsumoNaFaixa
                        Exit For
                    Case "tarfaixa" & btCont & "-1"
                        X = Format(.Tarifa, "#,##0.00") & IIf(.TipoTarifa = "fixo", "", "/m3")
                        Exit For
                    Case "subtfaixa" & btCont & "-1"
                        X = .SubTotal
                        X = Format(X, "#,##0.00")
                        Exit For
                    End Select
                End With
            End If
        
            If Right(strVariavel, 5) = "-2" Then
                If fatImpugnada.Categorias(2).Faixas.Count >= btCont Then
                    With fatImpugnada.Categorias(2).Faixas(btCont)
                        Select Case strVariavel
                        Case "faixa" & btCont & "-2"
                            X = .AbrangenciaFaixa
                            Exit For
                        Case "consfaixa" & btCont & "-2"
                            X = .ConsumoNaFaixa
                            Exit For
                        Case "tarfaixa" & btCont & "-2"
                            X = Format(.Tarifa, "#,##0.00") & IIf(.TipoTarifa = "fixo", "", "/m3")
                            Exit For
                        Case "subtfaixa" & btCont & "-2"
                            X = .SubTotal
                            X = Format(X, "#,##0.00")
                            Exit For
                        End Select
                    End With
                End If
            End If
        Next btCont
    End If
        
    ' Demais variáveis que não são valores de faixas
    Select Case strVariavel
    Case "referência-conta-exemplo"
        X = fatImpugnada.MesReferencia
    Case "consumo-conta-exemplo"
        X = fatImpugnada.ConsumoTotal
    Case "volume-paradigma"
        X = fatImpugnada.ConsumoTotal
    Case "consumo-por-economia"
        X = fatImpugnada.ConsumoPorEconomia
    Case "total-água"
        X = fatImpugnada.TotalAgua
        X = Format(X, "#,##0.00")
    Case "porcentagem-esgoto"
        X = fatImpugnada.PorcentEsgoto
    Case "total-água-esgoto"
        X = fatImpugnada.TotalAguaEsgoto
        X = Format(X, "#,##0.00")
    Case "qt-economias"
        X = fatImpugnada.QtdTotalEconomias
    Case "total-esgoto"
        X = fatImpugnada.TotalEsgoto
        X = Format(X, "#,##0.00")
    Case "qt-economias-1"
        X = fatImpugnada.Categorias(1).QtdEconomias
    Case "classificação-1"
        X = CategoriaExtenso(fatImpugnada.Categorias(1).Categoria)
    Case "água-por-econ-1"
        X = fatImpugnada.Categorias(1).AguaPorEconomia
        X = Format(X, "#,##0.00")
    Case "subt-água-1"
        X = fatImpugnada.Categorias(1).SubtotalAgua
        X = Format(X, "#,##0.00")
    Case "qt-economias-2"
        X = fatImpugnada.Categorias(2).QtdEconomias
    Case "classificação-2"
        X = CategoriaExtenso(fatImpugnada.Categorias(2).Categoria)
    Case "água-por-econ-2"
        X = fatImpugnada.Categorias(2).AguaPorEconomia
        X = Format(X, "#,##0.00")
    Case "subt-água-2"
        X = fatImpugnada.Categorias(2).SubtotalAgua
        X = Format(X, "#,##0.00")
    End Select
    
Case "Vaz. água ou extravas. esgoto com danos a patrimônio/morais", "Obra da Embasa com danos a patrimônio/morais", _
    "Acidente com pessoa/veículo em buraco", "Acidente com veículo (colisão ou  atropelamento)"
    ' Variáveis de Responsabilidade civil
    Select Case strVariavel
    Case "data-incidente"
        X = Trim(form.txtDataFato.Text)
    Case "local-incidente"
        X = InputBox(DeterminarTratamento & ", onde ocorreu o incidente narrado pelo Adverso?", "Informações sobre responsabilidade civil", "em Plataforma, bairro onde reside")
    Case "pediu-dano-material-RC"
        X = IIf(form.chbDanMat.Value = True, "indenização por danos materiais no valor pleiteado na Inicial, além de ", "")
    Case "descrição-dano-material-RC"
        X = IIf(form.chbDanMat.Value = True, " Alega ter sofrido prejuízos referentes " & InputBox(DeterminarTratamento & _
        ", descreva o dano material sofrido pela parte Autora: ""Alega ter sofrido prejuízos materiais referentes...""", _
        "Informações sobre responsabilidade civil", IIf(form.cmbOcorrencia.Text = "Pessoa", "a um aparelho de telefone celular quebrado", "aos reparos que se fizeram necessários")) & ".", "")
    Case "termo-ad-quem-prescrição"
        X = DateAdd("yyyy", 3, CDate(form.txtDataFato))
    Case "ato-apontado-ilícito"
        If strCausaPedir = "Vaz. água ou extravas. esgoto com danos a patrimônio/morais" Then
            X = InputBox(DeterminarTratamento & ", qual é a conduta que a parte Adversa aponta como ilícito?", "Informações sobre responsabilidade civil", "ocorreu extravasamento da rede pública para dentro de seu imóvel")
        
        ElseIf strCausaPedir = "Obra da Embasa com danos a patrimônio/morais" Then
            X = InputBox(DeterminarTratamento & ", qual é a conduta que a parte Adversa aponta como ilícito?", "Informações sobre responsabilidade civil", "sofreu prejuízos em decorrência de obra realizada pela Embasa")
        
        ElseIf strCausaPedir = "Acidente com pessoa/veículo em buraco" And form.cmbOcorrencia.Value = "Veículo" Then
            X = InputBox(DeterminarTratamento & ", qual é a conduta que a parte Adversa aponta como ilícito?", "Informações sobre responsabilidade civil", "um buraco realizado pela Embasa teria causado prejuízos ao seu veículo")
        
        ElseIf strCausaPedir = "Acidente com pessoa/veículo em buraco" And form.cmbOcorrencia.Value = "Pessoa" Then
            X = InputBox(DeterminarTratamento & ", qual é a conduta que a parte Adversa aponta como ilícito?", "Informações sobre responsabilidade civil", "sofreu prejuízos por ter caído em buraco realizado pela Embasa")
        
        ElseIf strCausaPedir = "Acidente com veículo (colisão ou atropelamento)" And form.cmbOcorrencia.Value = "Veículo" Then
            X = InputBox(DeterminarTratamento & ", qual é a conduta que a parte Adversa aponta como ilícito?", "Informações sobre responsabilidade civil", "um veículo a serviço da Embasa teria sido culpado por colisão com o seu veículo")
        
        ElseIf strCausaPedir = "Acidente com veículo (colisão ou atropelamento)" And form.cmbOcorrencia.Value = "Pessoa" Then
            X = InputBox(DeterminarTratamento & ", qual é a conduta que a parte Adversa aponta como ilícito?", "Informações sobre responsabilidade civil", "um veículo a serviço da Embasa teria sido culpado por atropelamento")
        End If
        
    Case "motivo-culpa-exclusiva"
        X = InputBox(DeterminarTratamento & ", qual é o motivo da culpa exclusiva do consumidor?", "Informações sobre responsabilidade civil", "não havia ninguém para franquear acesso ao imóvel")
    Case "descrição-lucros-cessantes"
        X = InputBox(DeterminarTratamento & ", favor descrever qual é o motivo aceitável dos lucros cessantes, se houver", "Informações sobre responsabilidade civil", "devem restringir-se a ")
    Case "valor-lucros-cessantes"
        X = InputBox(DeterminarTratamento & ", informe o valor aceitável dos lucros cessantes, se houver", "Informações sobre responsabilidade civil", "1.000,00")
    
    End Select

Case "Desabastecimentos por período e causa determinados"
    ' Variáveis de Desabastecimento
    Select Case strVariavel
    Case "bairro"
        X = InputBox(DeterminarTratamento & ", informe o bairro do imóvel da parte Autora", "Informações sobre desabastecimento")
    Case "corresponsável"
        X = form.cmbCorresponsavel.Value
    Case "data-desabastecimento"
        X = Trim(form.txtDataInicio.Text)
    Case "duração-alegada-desabastecimento"
        X = Trim(form.txtDuracao.Text)
    Case "mês-referência-fatura"
        X = InputBox(DeterminarTratamento & ", informe a fatura em que está o período impugnado", "Informações específicas de caso de Desabastecimento", "fatura de referência de Abril/2018")
    Case "prazo-final-prescrição"
        X = Day(CDate(form.txtDataFim.Text)) & "/" & Month(CDate(form.txtDataFim.Text)) & "/" & (Year(CDate(form.txtDataFim.Text)) + 3)
    Case "lista-de-processos-matrícula"
        X = Trim(form.txtProcessosSimilares.Text)
    Case "valor-condenação"
        X = form.txtValorCondenacao
        X = Format(X, "#,##0.00")
    Case "pediu-dano-material"
        X = IIf(form.chbDanMat.Value = True, "materiais e ", "")
    Case "pediu-devolução-em-dobro"
        X = IIf(form.chbDevolDobro.Value = True, ", bem como devolução em dobro das faturas do período", "")
    Case "exclusão-corresponsável"
        X = IIf(form.chbExcluiuCorresp.Value = True, " Entendeu que a " & form.cmbCorresponsavel.Value & " não é parte legítima para figurar no polo passivo da presente " & _
            "demanda.", "")
    Case "exclusão-corresponsável"
        If form.chbExcluiuCorresp.Visible = True Then X = IIf(form.chbExcluiuCorresp.Value = True, " Entendeu que a corré não é parte legítima para figurar no polo passivo da presente demanda", "")
    Case "condenou-dano-material"
        X = IIf(form.chbCondenouDanosMateriais.Value = True, "indenização por danos materiais, ", "")
    Case "condenou-devolução-em-dobro"
        X = IIf(form.chbCondenouDevDobro.Value = True, ", bem como devolução em dobro das faturas do período", "")
    Case "bairro-não-atingido"
        X = IIf(form.chbBairroNaoAfetado.Value = True, ", e o bairro da parte Autora não foi atingido", "")
    Case "sem-alteração-consumo"
        X = IIf(form.cmbAlteracaoConsumo.Value = "Não houve alteração relevante [gráfico]" Or _
            form.cmbAlteracaoConsumo.Value = "Não houve alteração [print de contas juntadas pelo Autor]", _
            ", além de que não houve alteração no consumo do imóvel", "")
    Case "motivo-incidente"
        X = InputBox(DeterminarTratamento & ", informe o motivo do desabastecimento alegado pela parte Autora, de forma resumida, para ser usado na introdução de alguns tópicos", "Informações específicas de caso de Desabastecimento", "em decorrência de conserto realizado na rede pública de abastecimento na sua região")
    Case "motivo-culpa-exclusiva-terceiro"
        X = InputBox(DeterminarTratamento & ", informe o motivo da culpa exclusiva de terceiro, de forma resumida, para constar na Contestação", "Informações específicas de caso de Desabastecimento", _
            "ignorou deliberadamente as múltiplas orientações recebidas desta empresa e danificou a adutora, com a intenção egoísta de cumprir seus prazos, " & _
            "mesmo que fosse às custas de destruir o abastecimento de água da coletividade.")
    End Select

Case "Desabastecimento Apagão Xingu 03/2018"
    ' Variáveis de Desabastecimento
    Select Case strVariavel
    Case "bairro"
        X = InputBox(DeterminarTratamento & ", informe o bairro do imóvel da parte Autora", "Informações específicas de caso CCR")
    Case "corresponsável"
        X = "Operadora Nacional do Sistema Elétrico - ONS, associação civil de CNPJ 02.831.210/0002-38"
    Case "motivo-culpa-exclusiva-terceiro"
        X = "suspendeu, por força maior, o fornecimento de energia elétrica necessário à operação das bombas hidráulicas, essenciais para a distribuição de água"
    Case "data-desabastecimento"
        X = Trim(form.txtDataInicio.Text)
    Case "duração-alegada-desabastecimento"
        X = Trim(form.txtDuracao.Text)
    Case "mês-referência-fatura"
        X = "faturas de referência de Abril/2018"
    Case "prazo-final-prescrição"
        X = "21/03/2021"
    Case "lista-de-processos-matrícula"
        X = Trim(form.txtProcessosSimilares.Text)
    Case "valor-condenação"
        X = form.txtValorCondenacao
        X = Format(X, "#,##0.00")
    Case "pediu-dano-material"
        X = IIf(form.chbDanMat.Value = True, "materiais e ", "")
    Case "pediu-devolução-em-dobro"
        X = IIf(form.chbDevolDobro.Value = True, ", bem como devolução em dobro das faturas do período", "")
    Case "exclusão-corresponsável"
        X = IIf(form.chbExcluiuCorresp.Value = True, " Entendeu que a " & form.cmbCorresponsavel.Value & " não é parte legítima para figurar no polo passivo da presente " & _
            "demanda.", "")
    Case "exclusão-corresponsável"
        If form.chbExcluiuCorresp.Visible = True Then X = IIf(form.chbExcluiuCorresp.Value = True, " Entendeu que a corré não é parte legítima para figurar no polo passivo da presente demanda", "")
    Case "condenou-dano-material"
        X = IIf(form.chbCondenouDanosMateriais.Value = True, "indenização por danos materiais, ", "")
    Case "condenou-devolução-em-dobro"
        X = IIf(form.chbCondenouDevDobro.Value = True, ", bem como devolução em dobro das faturas do período", "")
    Case "bairro-não-atingido"
        X = IIf(form.chbBairroNaoAfetado.Value = True, ", e o bairro da parte Autora não foi atingido", "")
    Case "sem-alteração-consumo"
        X = IIf(form.cmbAlteracaoConsumo.Value = "Não houve alteração relevante [gráfico]" Or _
            form.cmbAlteracaoConsumo.Value = "Não houve alteração [print de contas juntadas pelo Autor]", _
            ", além de que não houve alteração no consumo do imóvel", "")
    Case "motivo-incidente"
        X = ", em razão da falta de fornecimento de energia elétrica, decorrente de um apagão de energia elétrica que atingiu as regiões Norte, Nordeste e (parcialmente) Sudeste"
    
    End Select

Case "Desabastecimento Uruguai 09/2016"
    ' Variáveis de Desabastecimento
    Select Case strVariavel
    Case "bairro"
        X = InputBox(DeterminarTratamento & ", informe o bairro do imóvel da parte Autora", "Informações específicas de caso CCR")
    Case "data-desabastecimento"
        X = Trim(form.txtDataInicio.Text)
    Case "duração-alegada-desabastecimento"
        X = Trim(form.txtDuracao.Text)
    Case "mês-referência-fatura"
        X = "faturas de referência de Outubro e Novembro/2016"
    Case "prazo-final-prescrição"
        X = "14/09/2019"
    Case "lista-de-processos-matrícula"
        X = Trim(form.txtProcessosSimilares.Text)
    Case "valor-condenação"
        X = form.txtValorCondenacao
        X = Format(X, "#,##0.00")
    Case "pediu-dano-material"
        X = IIf(form.chbDanMat.Value = True, "materiais e ", "")
    Case "pediu-devolução-em-dobro"
        X = IIf(form.chbDevolDobro.Value = True, ", bem como devolução em dobro das faturas do período", "")
    Case "exclusão-corresponsável"
        X = IIf(form.chbExcluiuCorresp.Value = True, " Entendeu que a " & form.cmbCorresponsavel.Value & " não é parte legítima para figurar no polo passivo da presente " & _
            "demanda.", "")
    Case "exclusão-corresponsável"
        If form.chbExcluiuCorresp.Visible = True Then X = IIf(form.chbExcluiuCorresp.Value = True, " Entendeu que a corré não é parte legítima para figurar no polo passivo da presente demanda", "")
    Case "condenou-dano-material"
        X = IIf(form.chbCondenouDanosMateriais.Value = True, "indenização por danos materiais, ", "")
    Case "condenou-devolução-em-dobro"
        X = IIf(form.chbCondenouDevDobro.Value = True, ", bem como devolução em dobro das faturas do período", "")
    Case "bairro-não-atingido"
        X = IIf(form.chbBairroNaoAfetado.Value = True, ", e o bairro da parte Autora não foi atingido", "")
    Case "sem-alteração-consumo"
        X = IIf(form.cmbAlteracaoConsumo.Value = "Não houve alteração relevante [gráfico]" Or _
            form.cmbAlteracaoConsumo.Value = "Não houve alteração [print de contas juntadas pelo Autor]", _
            ", além de que não houve alteração no consumo do imóvel", "")
    Case "motivo-incidente"
        X = ", em decorrência de conserto realizado na rede pública de abastecimento na sua região"
    
    End Select

Case "Desabastecimento Liberdade 10/2017"
    ' Variáveis de Desabastecimento
    Select Case strVariavel
    Case "bairro"
        X = InputBox(DeterminarTratamento & ", informe o bairro do imóvel da parte Autora", "Informações específicas de caso CCR")
    Case "data-desabastecimento"
        X = Trim(form.txtDataInicio.Text)
    Case "duração-alegada-desabastecimento"
        X = Trim(form.txtDuracao.Text)
    Case "mês-referência-fatura"
        X = "fatura de referência de Dezembro/2017"
    Case "prazo-final-prescrição"
        X = "02/11/2020"
    Case "lista-de-processos-matrícula"
        X = Trim(form.txtProcessosSimilares.Text)
    Case "valor-condenação"
        X = form.txtValorCondenacao
        X = Format(X, "#,##0.00")
    Case "pediu-dano-material"
        X = IIf(form.chbDanMat.Value = True, "materiais e ", "")
    Case "pediu-devolução-em-dobro"
        X = IIf(form.chbDevolDobro.Value = True, ", bem como devolução em dobro das faturas do período", "")
    Case "exclusão-corresponsável"
        X = IIf(form.chbExcluiuCorresp.Value = True, " Entendeu que a " & form.cmbCorresponsavel.Value & " não é parte legítima para figurar no polo passivo da presente " & _
            "demanda.", "")
    Case "exclusão-corresponsável"
        If form.chbExcluiuCorresp.Visible = True Then X = IIf(form.chbExcluiuCorresp.Value = True, " Entendeu que a corré não é parte legítima para figurar no polo passivo da presente demanda", "")
    Case "condenou-dano-material"
        X = IIf(form.chbCondenouDanosMateriais.Value = True, "indenização por danos materiais, ", "")
    Case "condenou-devolução-em-dobro"
        X = IIf(form.chbCondenouDevDobro.Value = True, ", bem como devolução em dobro das faturas do período", "")
    Case "bairro-não-atingido"
        X = IIf(form.chbBairroNaoAfetado.Value = True, ", e o bairro da parte Autora não foi atingido", "")
    Case "sem-alteração-consumo"
        X = IIf(form.cmbAlteracaoConsumo.Value = "Não houve alteração relevante [gráfico]" Or _
            form.cmbAlteracaoConsumo.Value = "Não houve alteração [print de contas juntadas pelo Autor]", _
            ", além de que não houve alteração no consumo do imóvel", "")
    Case "motivo-incidente"
        X = ", em decorrência de conserto realizado na rede pública de abastecimento na sua região"
    
    End Select

Case "Desabastecimento CCR 04/2015"
    ' Variáveis de Desabastecimento
    Select Case strVariavel
    Case "bairro"
        X = InputBox(DeterminarTratamento & ", informe o bairro do imóvel da parte Autora", "Informações específicas de caso CCR")
    Case "corresponsável"
        X = "CCR Metrô"
    Case "data-desabastecimento"
        X = Trim(form.txtDataInicio.Text)
    Case "duração-alegada-desabastecimento"
        X = Trim(form.txtDuracao.Text)
    Case "mês-referência-fatura"
        X = "fatura de referência de Maio/2015"
    Case "prazo-final-prescrição"
        X = "08/04/2018"
    Case "lista-de-processos-matrícula"
        X = Trim(form.txtProcessosSimilares.Text)
    Case "valor-condenação"
        X = form.txtValorCondenacao
        X = Format(X, "#,##0.00")
    Case "pediu-dano-material"
        X = IIf(form.chbDanMat.Value = True, "materiais e ", "")
    Case "pediu-devolução-em-dobro"
        X = IIf(form.chbDevolDobro.Value = True, ", bem como devolução em dobro das faturas do período", "")
    Case "exclusão-corresponsável"
        X = IIf(form.chbExcluiuCorresp.Value = True, " Entendeu que a corré não é parte legítima para figurar no polo passivo da presente demanda", "")
    Case "condenou-dano-material"
        X = IIf(form.chbCondenouDanosMateriais.Value = True, "indenização por danos materiais, ", "")
    Case "condenou-devolução-em-dobro"
        X = IIf(form.chbCondenouDevDobro.Value = True, ", bem como devolução em dobro das faturas do período", "")
    Case "bairro-não-atingido"
        X = IIf(form.chbBairroNaoAfetado.Value = True, ", e o bairro da parte Autora não foi atingido", "")
    Case "sem-alteração-consumo"
        X = IIf(form.cmbAlteracaoConsumo.Value = "Não houve alteração relevante [gráfico]" Or _
            form.cmbAlteracaoConsumo.Value = "Não houve alteração [print de contas juntadas pelo Autor]", _
            ", além de que não houve alteração no consumo do imóvel", "")
    Case "motivo-incidente"
        X = ", em decorrência do notório rompimento de uma adutora, causado pela Companhia do Metrô da Bahia em 01/04/2015"
    Case "motivo-culpa-exclusiva-terceiro"
        X = "pois a CCR ignorou as orientações da Embasa e destruiu a adutora que abastece parte relevante da cidade com dolo eventual, pelo motivo espúrio de não pagar multa administrativa pelo atraso na obra"
    
    End Select

Case "Outros vícios"
    ' Variáveis de outros vícios
    Select Case strVariavel
    Case "sanado-administrativamente"
        X = IIf(form.chbVicioConsertado.Value = True, ", o qual foi sanado administrativamente", "")
    Case "mês-final"
        X = InputBox(DeterminarTratamento & ", informe o mês de referência da fatura final do período impugnado pela parte Adversa", "Informações sobre pedido", "agosto/2019")
    Case "consumo-afirmado-autor"
        X = Trim(InputBox(DeterminarTratamento & ", informe o consumo alegado pelo Autor, completando a frase abaixo:" & vbCrLf & """Alega que seu consumo efetivo é ... m3""", "Informações sobre pedido", "06"))
    Case "motivo-desligamento"
        strCont = Trim(InputBox(DeterminarTratamento & ", informe o motivo do pedido de suspensão do abastecimento, completando a frase abaixo:" & vbCrLf & """(...) requereu o desligamento do abastecimento de água, ...""", "Informações sobre pedido", "haja vista haver deixado de residir no imóvel"))
        X = IIf(strCont <> "", ", " & strCont, "")
    Case "motivo-inexistência-vício"
        strCont = Trim(InputBox(DeterminarTratamento & ", informe o motivo do pedido de inexistência de vício, completando a frase abaixo:" & vbCrLf & """(...) não houve nem defeito nem vício do serviço. ...""", "Informações sobre pedido", "As cobranças realizadas pela Embasa foram feitas nos estritos limites dos valores contratados"))
        X = IIf(strCont <> "", ". " & strCont, "")
    Case "pretensão-autoral"
        X = Trim(InputBox(DeterminarTratamento & ", informe a pretensão da parte Autora, completando a frase abaixo:" & vbCrLf & """(...) Caso a parte Autora, por questões particulares, pretenda ...""", "Informações sobre pedido", "suspender as cobranças de tarifa de esgoto"))
    Case "requisitos-pretensão"
        strCont = Trim(InputBox(DeterminarTratamento & ", informe os requisitos da pretensão da parte Autora, completando a frase abaixo:" & vbCrLf & """(...) deveria requerer à Embasa e demonstrar o cumprimento dos requisitos (...)""", "Informações sobre pedido", "desabitação do imóvel"))
        X = IIf(strCont <> "", " (" & strCont & ")", "")
    Case "data-requerimento-corte"
        X = Trim(InputBox(DeterminarTratamento & ", informe a data em que a parte Autora requereu o corte", "Informações sobre pedido"))
    
    End Select

End Select

AtribuiVariavel:

ObterVariaveis = X

End Function

Function DescobrirPlanilhaDeEstrutura(strPeticao As String, strOrgao As String) As Excel.Worksheet
''
'' Descobre qual planilha tem a lista correta de tópicos, conforme o tipo de petição e o órgão.
''
    Dim plan As Excel.Worksheet
    
    Select Case strPeticao
    Case "Contestação"
        Select Case strOrgao
        Case "JEC"
            Set plan = ThisWorkbook.Sheets("cfContestaçõesJEC")
        Case "VC"
            Set plan = ThisWorkbook.Sheets("cfContestaçõesVC")
        Case "Procon"
            Set plan = ThisWorkbook.Sheets("cfContestaçõesProcon")
        End Select
        
    Case "RI/Apelação"
        Select Case strOrgao
        Case "JEC"
            Set plan = ThisWorkbook.Sheets("cfRIsJEC")
        Case "VC"
            Set plan = ThisWorkbook.Sheets("cfApelaçõesVC")
        'Case "Procon"
        '    Set Plan = ThisWorkbook.Sheets("cfRIsProcon")
        End Select
                
    Case "Contrarrazões de RI/Apelação"
        Select Case strOrgao
        Case "JEC"
            Set plan = ThisWorkbook.Sheets("cfCRRIsJEC")
        Case "VC"
            Set plan = ThisWorkbook.Sheets("cfCRApelaçõesVC")
        'Case "Procon"
        '    Set Plan = ThisWorkbook.Sheets("cfContestaçõesProcon")
        End Select
        
    End Select
    
    Set DescobrirPlanilhaDeEstrutura = plan

End Function


Function ColherTopicosGeral(strCausaPedir As String, strCausaPedirParadigma As String, strTipoPeticao As String, form As Variant, ByRef bolGrafico As Boolean, ByRef btTabela As Byte) As String
''
'' Chama as funções de colher os tópicos conforme a causa de pedir. Também altera a variável "
''
    
    Dim strTopicos As String
    
    Select Case strCausaPedirParadigma
    Case "Revisão de consumo elevado"
        Select Case strTipoPeticao
        Case "Contestação"
            strTopicos = ColherTopicosContRevisaoConsumo(form)
            'Verifica se é um dos tópicos que contém gráficos
            If InStr(1, strTopicos, ",,50,,") <> 0 Or InStr(1, strTopicos, ",,55,,") <> 0 Then bolGrafico = True
            
            'Verifica se é um dos tópicos que contém tabelas de média, anotando quantas tabelas são.
            btTabela = 0
            If InStr(1, strTopicos, ",,95,,") <> 0 Then btTabela = btTabela + 1
            If InStr(1, strTopicos, ",,115,,") <> 0 Then btTabela = btTabela + 1
            If InStr(1, strTopicos, ",,120,,") <> 0 Then btTabela = btTabela + 1
        
        'Case "RI/Apelação"
        '    strTopicos = ColherTopicosRICCR(form)
        '    If InStr(1, strTopicos, ",,40,,") <> 0 Or InStr(1, strTopicos, ",,42,,") <> 0 Then bolGrafico = True 'Verifica se é um dos tópicos que contém gráficos
            
        'Case "Contrarrazões de RI/Apelação"
        '    strTopicos = ColherTopicosCRRICCR(form)
        '    If InStr(1, strTopicos, ",,40,,") <> 0 Or InStr(1, strTopicos, ",,42,,") <> 0 Then bolGrafico = True 'Verifica se é um dos tópicos que contém gráficos
            
        End Select
        
    Case "Corte no fornecimento"
        Select Case strTipoPeticao
        Case "Contestação"
            strTopicos = ColherTopicosContCorte(form)
            'Verifica se é um dos tópicos que contém gráficos
            If InStr(1, strTopicos, ",,50,,") <> 0 Or InStr(1, strTopicos, ",,55,,") <> 0 Then bolGrafico = True
                        
        Case "RI/Apelação"
            strTopicos = ColherTopicosRICorte(form)
            
        Case "Contrarrazões de RI/Apelação"
            'strTopicos = ColherTopicosCRRIDesabGenerico(form)
            
        End Select
    
    Case "Realizar ligação de água"
        strTopicos = ColherTopicosContRealizarLigacao(form)
        
    Case "Negativação no SPC"
        strTopicos = ColherTopicosContNegativacao(form)
    
    Case "Cobrança de esgoto em imóvel não ligado à rede"
        Select Case strTipoPeticao
        Case "Contestação"
            strTopicos = ColherTopicosContEsgotoSemRede(form)
            
        Case "RI/Apelação"
            'strTopicos = ColherTopicosRIDesabGenerico(form)
            
        Case "Contrarrazões de RI/Apelação"
            'strTopicos = ColherTopicosCRRIDesabGenerico(form)
            
        End Select
        
    Case "Cobrança de esgoto com água cortada"
        Select Case strTipoPeticao
        Case "Contestação"
            strTopicos = ColherTopicosContEsgotoAguaCortada(form)
            
        Case "RI/Apelação"
            'strTopicos = ColherTopicosRIDesabGenerico(form)
            
        Case "Contrarrazões de RI/Apelação"
            'strTopicos = ColherTopicosCRRIDesabGenerico(form)
            
        End Select
        
    Case "Classificação tarifa ou qtd. de economias"
        Select Case strTipoPeticao
        Case "Contestação"
            strTopicos = ColherTopicosContClassificacaoTarifaria(form)
            
        Case "RI/Apelação"
            'strTopicos = ColherTopicosRIDesabGenerico(form)
            
        Case "Contrarrazões de RI/Apelação"
            'strTopicos = ColherTopicosCRRIDesabGenerico(form)
            
        End Select
        
    Case "Suspeita de by-pass"
        Select Case strTipoPeticao
        Case "Contestação"
            strTopicos = ColherTopicosContGato(form)
            
            'O tópico de gráfico está sempre
            bolGrafico = True
            
        Case "RI/Apelação"
            'strTopicos = ColherTopicosRIDesabGenerico(form)
            
        Case "Contrarrazões de RI/Apelação"
            'strTopicos = ColherTopicosCRRIDesabGenerico(form)
            
        End Select
        
    Case "Débito de terceiro"
        Select Case strTipoPeticao
        Case "Contestação"
            strTopicos = ColherTopicosContDebitoTerceiro(form)
            
        Case "RI/Apelação"
            'strTopicos = ColherTopicosRIDesabGenerico(form)
            
        Case "Contrarrazões de RI/Apelação"
            'strTopicos = ColherTopicosCRRIDesabGenerico(form)
            
        End Select
        
    Case "Vaz. água ou extravas. esgoto com danos a patrimônio/morais", "Obra da Embasa com danos a patrimônio/morais", _
        "Acidente com pessoa/veículo em buraco", "Acidente com veículo (colisão ou atropelamento)"
        Select Case strTipoPeticao
        Case "Contestação"
            strTopicos = ColherTopicosContRespCivil(form, strCausaPedir)
            
        Case "RI/Apelação"
            'strTopicos = ColherTopicosRIRespCivil(form)
            
        Case "Contrarrazões de RI/Apelação"
            'strTopicos = ColherTopicosCRRIRespCivil(form)
            
        End Select
        
    Case "Desabastecimentos por período e causa determinados", "Desabastecimento CCR 04/2015", "Desabastecimento Uruguai 09/2016", _
        "Desabastecimento Liberdade 10/2017", "Desabastecimento Apagão Xingu 03/2018"
        Select Case strTipoPeticao
        Case "Contestação"
            strTopicos = ColherTopicosContDesabastecimento(form)
            
        Case "RI/Apelação"
            strTopicos = ColherTopicosRIDesabastecimento(form)
            
        Case "Contrarrazões de RI/Apelação"
            'GoTo NãoFaz
            'strTopicos = ColherTopicosCRRIDesabGenerico(form)
            
        End Select
        
        If InStr(1, strTopicos, ",,40,,") <> 0 Or InStr(1, strTopicos, ",,42,,") <> 0 Then bolGrafico = True 'Verifica se é um dos tópicos que contém gráficos
            
    Case "Fixo de esgoto"
        Select Case strTipoPeticao
        Case "Contestação"
            strTopicos = ColherTopicosContFixoEsgoto(form, strCausaPedir)
            
        Case "RI/Apelação"
            'strTopicos = ColherTopicosRIDesabApagXingu2018(form)
            
        Case "Contrarrazões de RI/Apelação"
            'strTopicos = ColherTopicosCRRIDesabApagXingu2018(form)
            
        End Select
        
    Case "Outros vícios"
        Select Case strTipoPeticao
        Case "Contestação"
            strTopicos = ColherTopicosContVicio(form, strCausaPedir)
            
        Case "RI/Apelação"
            'strTopicos = ColherTopicosRIDesabApagXingu2018(form)
            
        Case "Contrarrazões de RI/Apelação"
            'strTopicos = ColherTopicosCRRIDesabApagXingu2018(form)
            
        End Select
        
    End Select
    
    ColherTopicosGeral = strTopicos

End Function

Function ColherTopicosContPedidosIniciais(form As Variant) As String
''
'' Metafunção para ser usado em outras funções de colher tópicos. Colhe os tópicos exclusivamente da frame de pedidos iniciais.
''
    ' Agrupa os números dos tópicos numa string, cercados por vírgulas duplas.
    Dim strTopicos As String
    strTopicos = ",,"
    
    With form
        
        'Pedidos gerais
        ''Liminar cumprida
        If .chbLiminarCumprida.Value = True Then strTopicos = strTopicos & "200,,"
        If .chbIncompetenciaTerritorial = True Then strTopicos = strTopicos & "205,,220,,"
        If .chbIlegitimidade = True Then strTopicos = strTopicos & "210,,292,,"
        If .chbDecadencia = True Then strTopicos = strTopicos & "270,,296,,"

        'If .chbPrescricao = True Then strTopicos = strTopicos & ",,"
        
    End With
    
    ColherTopicosContPedidosIniciais = strTopicos

End Function

Function ColherTopicosContPedidosFinais(form As Variant) As String
''
'' Metafunção para ser usado em outras funções de colher tópicos. Colhe os tópicos exclusivamente da frame de pedidos gerais.
''
    ' Agrupa os números dos tópicos numa string, cercados por vírgulas duplas.
    Dim strTopicos As String
    
    With form
        ''Devolução em dobro
        If .chbDevolDobro.Value = True Then
            If .cmbPagamento.Value = "Houve pagamento" Then
                strTopicos = strTopicos & "800,,897,,"
            ElseIf .cmbPagamento.Value = "Não houve pagamento" Then
                strTopicos = strTopicos & "800,,805,,897,,"
            End If
        End If
        
        ''Dano material
        If .chbDanMat.Value = True Then
            If .chbDanMatSemProvas.Value = True Then
                strTopicos = strTopicos & "810,,"
            Else
                strTopicos = strTopicos & "815,,"
            End If
            If .chbValorLucroCessante.Value = True Then strTopicos = strTopicos & "820,,"
        End If
        
        ''Dano moral
        If .chbDanoMoral.Value = True Then
            strTopicos = strTopicos & "825,,"
            If .optAutorCondominio.Value = True Then strTopicos = strTopicos & "828,,"
            If .optAutorOutrosPJ.Value = True Then strTopicos = strTopicos & "827,,"
            If .chbDanMorMeraCobranca.Value = True Then strTopicos = strTopicos & "830,,"
            If .chbDanMorCorte.Value = True Then strTopicos = strTopicos & "835,,"
            If .chbDanMorNegativacao.Value = True Then strTopicos = strTopicos & "840,,"
            If .chbDanMorMeraCobranca.Value = False And .chbDanMorCorte.Value = False And .chbDanMorNegativacao.Value = False Then strTopicos = strTopicos & "845,,"
        End If
        
        ''Litigância de má-fé
        If .chbLitigMaFe.Value = True Then strTopicos = strTopicos & "880,,899,,"
        
    End With
    
    ColherTopicosContPedidosFinais = strTopicos

End Function

Function ColherTopicosContRevisaoConsumo(form As Variant) As String
''
'' Recolhe os tópicos para a Contestação de Revisão de Consumo. Aqui não importa a ordem (a ordem
''    dos tópicos será a que está na planilha de apoio).
''

    ' Agrupa os números dos tópicos numa string, cercados por vírgulas duplas.
    Dim strTopicos As String
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
       'Perfil de consumo
        If .cmbPadraoConsumo.Value = "Há padrão, mas o consumo impugnado é idêntico ou menor que a média de consumo" Then
            strTopicos = strTopicos & "25,,95,,155,,"
        ElseIf .cmbPadraoConsumo.Value = "Há padrão, mas o consumo impugnado é razoavelmente compatível com a média" Then
            strTopicos = strTopicos & "55,,"
        ElseIf .cmbPadraoConsumo.Value = "Não há padrão definido, consumo cheio de altos e baixos" Then
            strTopicos = strTopicos & "50,,"
        ElseIf .cmbPadraoConsumo.Value = "Sem padrão anterior, impugna consumos medidos após início do contrato pelo mínimo" Then
            strTopicos = strTopicos & "60,,"
        ElseIf .cmbPadraoConsumo.Value = "Sem padrão anterior, impugna consumos medidos após longo tempo sem hidrômetro" Then
            strTopicos = strTopicos & "61,,"
        End If
        
        'Há Medição individualizada, mas consumo rateado não foi relevante no aumento
        If .chbMIIrrelevante.Value = True Then strTopicos = strTopicos & "5,,"
        
        'Defender parcelamento
        If .chbParcelamento.Value = True Then strTopicos = strTopicos & "125,,"
        
        ' Defender cobrança de esgoto durante corte
        If .chbEsgotoAposCorte.Value = True Then strTopicos = strTopicos & "65,,"
        
        'Consumo de acordo com a Média do Procon/SP
        If .chbMediaProcon.Value = True Then strTopicos = strTopicos & "110,,"
        
        'Existência de aferição
        If .cmbAferHidrometro.Value = "Há, hidrômetro regular" Then
            strTopicos = strTopicos & "75,,180,,"
        ElseIf .cmbAferHidrometro.Value = "Há, irregular contra a fornecedora (medindo a menor)" Then
            strTopicos = strTopicos & "80,,185,,"
        End If
        
        'Requerer aferição
        If .chbRequerAfericao.Value = True Then
            If .chbMIIrrelevante.Value = True Then
                strTopicos = strTopicos & "70,,170,,"
            Else
                strTopicos = strTopicos & "65,,165,,"
            End If
        End If
        
        'Vazamento interno
        If .cmbVazInterno.Value = "Há SS no SCI (anexar SS!)" Then
            strTopicos = strTopicos & "15,,30,,90,,150,,"
        ElseIf .cmbVazInterno.Value = "Há confissão" Then
            strTopicos = strTopicos & "10,,23,,85,,150,,"
        End If
        
        'Hidrômetro já foi substituído, hidrômetro novo corrobora medições
        If .chbHidrTrocado.Value = True Then
            strTopicos = strTopicos & "100,,"
        End If
        
        'Média de consumo viciada por processos anteriores
        If .chbMediaViciada.Value = True Then strTopicos = strTopicos & "105,,"
        
        'Média de consumo correta
        If .chbMediaCorreta.Value = True Then
            If .chbMediaConsRetificado.Value = True Then
                strTopicos = strTopicos & "115,,190,,"
            Else
                strTopicos = strTopicos & "120,,190,,"
            End If
        End If
        
        'Se já não tem tópico de mérito com hidrômetro regular ou a menor, coloca o genérico
        If InStr(1, strTopicos, ",,180,,") = 0 And InStr(1, strTopicos, ",,185,,") = 0 Then strTopicos = strTopicos & "175,,"
        
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosFinais(form)
        
    End With
    
    ColherTopicosContRevisaoConsumo = strTopicos

End Function

Function ColherTopicosContCorte(form As Variant) As String
''
'' Recolhe os tópicos para a Contestação de Corte. Aqui não importa a ordem (a ordem
''    dos tópicos será a que está na planilha de apoio).
''

    ' Agrupa os números dos tópicos numa string, cercados por vírgulas duplas.
    Dim strTopicos As String
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
        'Particularidades do corte
        If .chbNaoHouveCorte.Value = True Then strTopicos = strTopicos & "10,,15,,"
        If .chbSemAlteracaoConsumo.Value = True Then strTopicos = strTopicos & "37,,"
        If .chbFaturasAberto.Value = True Then strTopicos = strTopicos & "20,,"
        If .chbInadimplente.Value = True Then strTopicos = strTopicos & "30,,"
        If .chbComprovanteIlegivel.Value = True Then strTopicos = strTopicos & "35,,"
        'If .chbContaErrada.Value = True Then strTopicos = strTopicos & "40,,"
        If .chbPagVespera.Value = True Then strTopicos = strTopicos & "55,,65,,"
        If .chbCorteBreve.Value = True Then strTopicos = strTopicos & "60,,"
        If .chbEsgotoAposCorte.Value = True Then strTopicos = strTopicos & "65,,"
            
        
        'Aviso de corte
        If .cmbAvisoCorte.Value = "Houve, em faturas anteriores / não houve" Then
            strTopicos = strTopicos & "50,,"
        ElseIf .cmbAvisoCorte.Value = "Houve, em correspondência específica" Then
            strTopicos = strTopicos & "45,,"
        End If
        
        ' Pedido (tópico especial para pagamento na véspera)
        If .chbPagVespera.Value = True Then strTopicos = strTopicos & "80,," Else strTopicos = strTopicos & "70,,"
        
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosFinais(form)
        
    End With
    
    ColherTopicosContCorte = strTopicos

End Function

Function ColherTopicosContNegativacao(form As Variant) As String
''
'' Recolhe os tópicos para a Contestação de Realizar ligação. Aqui não importa a ordem (a ordem
''    dos tópicos será a que está na planilha de apoio).
''

    ' Agrupa os números dos tópicos numa string, cercados por vírgulas duplas.
    Dim strTopicos As String
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
        'Atitude do Autor
        If .cmbAtitudeAutor.Value = "Autor afirma, de forma genérica, ""desconhecer"" dívida" Then
            strTopicos = strTopicos & "20,,80,,72,,90,,"
        ElseIf .cmbAtitudeAutor.Value = "Autor afirma claramente que não firmou contrato" Then
            strTopicos = strTopicos & "10,,72,,90,,"
        ElseIf .cmbAtitudeAutor.Value = "Autor reconhece contrato mas nega débitos" Then
            strTopicos = strTopicos & "15,,"
        ElseIf .cmbAtitudeAutor.Value = "Autor diz que se mudou e pediu suspensão do serviço" Then
            strTopicos = strTopicos & "17,,68,,72,,90,,"
        End If
        
        'Perfil do contrato
        If .cmbPerfilContrato.Value = "Sem uso nem pagamento, aparência de fraude (não comentar)" Then
            strTopicos = strTopicos & "35,,"
        ElseIf .cmbPerfilContrato.Value = "Há uso e pagamentos mais ou menos regulares" Then
            strTopicos = strTopicos & "40,,80,,"
        ElseIf .cmbPerfilContrato.Value = "Houve uso e pagamentos regulares até certa data" Then
            strTopicos = strTopicos & "42,,80,,"
        ElseIf .cmbPerfilContrato.Value = "A negativação foi de parcelamento firmado pelo Autor" Then
            strTopicos = strTopicos & "45,,"
        End If
        
        ' Prova da negativação
        If .chbProvaNegativacao.Value = True Then
            If .cmbAtitudeAutor.Value <> "Autor diz que se mudou e pediu suspensão do serviço" Then strTopicos = strTopicos & "27,,"
        Else
            strTopicos = strTopicos & "28,,67,,87,,"
        End If
        
        'Particularidades do caso
        If .chbSemComprovResidencia.Value = True Then strTopicos = strTopicos & "61,,"
        If .chbComprovResidenciaDeTerceiro.Value = True Then strTopicos = strTopicos & "62,,"
        If .chbNegativacaoPrevia.Value = True Then strTopicos = strTopicos & "65,,85,,"
        If .chbInadimplente.Value = True Then strTopicos = strTopicos & "70,,"
        If .chbProvaContrato.Value = True Then strTopicos = strTopicos & "30,,"
        If .chbMultiplosProcessos.Value = True Then strTopicos = strTopicos & "25,,75,,"
        
        'Endereço da qualificação e contrato com a Coelba
        If .chbEnderecoQualificacao.Value = True And .chbContratoCoelba.Value = True Then
            strTopicos = strTopicos & "60,,"
        ElseIf .chbEnderecoQualificacao.Value = True And .chbContratoCoelba.Value = False Then
            strTopicos = strTopicos & "50,,"
        ElseIf .chbEnderecoQualificacao.Value = False And .chbContratoCoelba.Value = True Then
            strTopicos = strTopicos & "55,,"
        End If
        
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosFinais(form)
        
    End With
    
    ColherTopicosContNegativacao = strTopicos

End Function

Function ColherTopicosContRealizarLigacao(form As Variant) As String
''
'' Recolhe os tópicos para a Contestação de Realizar ligação. Aqui não importa a ordem (a ordem
''    dos tópicos será a que está na planilha de apoio).
''

    ' Agrupa os números dos tópicos numa string, cercados por vírgulas duplas.
    Dim strTopicos As String
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
        'Particularidades do caso
        If (.chbNaoHouveSolicitacao.Value = True Or .chbNaoHouveRecusa.Value = True) Then
            strTopicos = strTopicos & "10,," 'Cabeçalho dos esclarecimentos simples
            If .chbNaoHouveSolicitacao.Value = True Then strTopicos = strTopicos & "13,,28,,110,,"
            If .chbNaoHouveRecusa.Value = True Then strTopicos = strTopicos & "15,,115,,"
            
        Else
            strTopicos = strTopicos & "20,," 'Cabeçalho das pretensões atécnicas
            If .chbSemReservatorioBomba.Value = True Then strTopicos = strTopicos & "21,,30,,100,,"
            If .chbSemReservacao.Value = True Then strTopicos = strTopicos & "22,,31,,100,,"
            If .chbSepararInstalacoesInternas.Value = True Then strTopicos = strTopicos & "23,,32,,100,,"
            If .chbDistanciaRede.Value = True Then strTopicos = strTopicos & "24,,33,,100,,"
            If .chbAltitudeInsuficiente.Value = True Then strTopicos = strTopicos & "34,,100,,"
            
        End If
        
        'Pedidos específicos de Realizar ligação
        If .chbDanMat.Value = True Then strTopicos = strTopicos & "120,,"
        If .chbDanoMoral.Value = True Then
            If .chbNaoHouveSolicitacao.Value = True Then
                strTopicos = strTopicos & "80,,"
            ElseIf .chbNaoHouveRecusa.Value = True Then
                strTopicos = strTopicos & "85,,"
            Else
                strTopicos = strTopicos & "70,,"
            End If
        End If
        
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosFinais(form)
        
    End With
    
    ColherTopicosContRealizarLigacao = strTopicos

End Function

Function ColherTopicosContEsgotoSemRede(form As Variant) As String
''
'' Recolhe os tópicos para a Contestação de Cobrança de esgoto sem ligação à rede. Aqui não importa a ordem (a ordem
''    dos tópicos será a que está na planilha de apoio).
''

    ' Agrupa os números dos tópicos numa string, cercados por vírgulas duplas.
    Dim strTopicos As String
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
        'Particularidades do caso
        If .chbOmiteDestinacao.Value = True Then strTopicos = strTopicos & "5,,"
        If .chbCacaEsgoto.Value = True Then strTopicos = strTopicos & "10,,"
        If .chbGeorref.Value = True Then strTopicos = strTopicos & "13,,"
        If .chbProvaFotografica.Value = True Then strTopicos = strTopicos & "16,,"
        If .chbInspecaoJudicial.Value = True Then strTopicos = strTopicos & "20,,25,,"
        
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosFinais(form)
        
    End With
    
    ColherTopicosContEsgotoSemRede = strTopicos

End Function

Function ColherTopicosContEsgotoAguaCortada(form As Variant) As String
''
'' Recolhe os tópicos para a Contestação de Cobrança de esgoto com água cortada. Aqui não importa a ordem (a ordem
''    dos tópicos será a que está na planilha de apoio).
''

    ' Agrupa os números dos tópicos numa string, cercados por vírgulas duplas.
    Dim strTopicos As String
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
        'Prova da habitação e uso do imóvel
        If .cmbProvaUsoImovel.Value = "Autor não alega imóvel desabitado; há inspeções" Then strTopicos = strTopicos & "10,,24,,35,,50,,"
        If .cmbProvaUsoImovel.Value = "Autor confessa que imóvel estava habitado" Then strTopicos = strTopicos & "10,,20,,35,,"
        If .cmbProvaUsoImovel.Value = "Inspeções realizadas pelos técnicos da Embasa" Then strTopicos = strTopicos & "10,,28,,35,,50,,"
        
        'Pedir depoimento pessoal
        If .cmbProvaUsoImovel.Value <> "Autor confessa que imóvel estava habitado" Then strTopicos = strTopicos & "95,,"
        
        'Requerimento administrativo
        If .chbSemRequerimentoAdm.Value = True Then strTopicos = strTopicos & "60,,"
        
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosFinais(form)
        
    End With
    
    ColherTopicosContEsgotoAguaCortada = strTopicos

End Function

Function ColherTopicosContClassificacaoTarifaria(form As Variant) As String
''
'' Recolhe os tópicos para a Contestação de Classificação tarifária. Aqui não importa a ordem (a ordem
''    dos tópicos será a que está na planilha de apoio).
''

    ' Agrupa os números dos tópicos numa string, cercados por vírgulas duplas.
    Dim strTopicos As String
    Dim intAno As Integer, btMes As Byte
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
        'Pretensão - alegação de multiplicação do mínimo por múltiplas economias?
        If .optPretMultEcon.Value = True Then
            strTopicos = strTopicos & "5,,"
        ElseIf .optPretUmaEcon.Value = True Then
            strTopicos = strTopicos & "10,,35,,"
        End If
        
        'Necessidade de esclarecer período?
        If .chbPeriodoMenor.Value = True Then
            If .chbTemAlteracao2.Value = True Then
                strTopicos = strTopicos & "17,,"
            Else
                strTopicos = strTopicos & "15,,"
            End If
            
        End If
        
        'Particularidades do caso
        If .chbInspecaoJud.Value = True Then strTopicos = strTopicos & "30,,120,,"
        If .chbDesvantajoso.Value = True Then strTopicos = strTopicos & "85,,125,,"
        If .chbRetroativoSemSolicitacao.Value = True Then strTopicos = strTopicos & "90,,"
        
        'Tarifa aplicável
        If .chbApresentarCalcExemplo.Value = True Then
            strTopicos = strTopicos & "40,,80,,"
            'Descobre qual a Resolução de tarifa
            intAno = CInt(Right(.txtMesRefExemplo.Text, 4))
            btMes = CByte(Left(.txtMesRefExemplo.Text, 2))
            Select Case btMes
            Case 1, 2, 3, 4, 5, 6
                intAno = intAno - 1
            End Select
            
            Select Case intAno
            Case 2013
                strTopicos = strTopicos & "45,,"
            Case 2014
                strTopicos = strTopicos & "50,,"
            Case 2015
                strTopicos = strTopicos & "55,,"
            Case 2016
                strTopicos = strTopicos & "60,,"
            Case 2017
                strTopicos = strTopicos & "65,,"
            Case 2018
                strTopicos = strTopicos & "70,,"
            Case 2019
                strTopicos = strTopicos & "75,,"
            End Select
        End If
        
        ' Cálculo em si
        If .chbApresentarCalcExemplo.Value = True Then
            If .chbTemDoisTipos.Value = True Then
                strTopicos = strTopicos & "78,,"
            Else
                strTopicos = strTopicos & "77,,"
            End If
        End If
        
        'Pedido contraposto
        If .chbReconvencao.Value = True Then strTopicos = strTopicos & "110,,135,,"
        
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosFinais(form)
        
    End With
    
    ColherTopicosContClassificacaoTarifaria = strTopicos

End Function

Function ColherTopicosContGato(form As Variant) As String
''
'' Recolhe os tópicos para a Contestação de Gato. Aqui não importa a ordem (a ordem
''    dos tópicos será a que está na planilha de apoio).
''

    ' Agrupa os números dos tópicos numa string, cercados por vírgulas duplas.
    Dim strTopicos As String
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
        'Tipo de gato
        If .cmbTipoGato.Value = "Desvio - bypass" Then strTopicos = strTopicos & "10,,"
        If .cmbTipoGato.Value = "Hidrômetro furado" Then strTopicos = strTopicos & "20,,"
        If .cmbTipoGato.Value = "Hidrômetro invertido" Then strTopicos = strTopicos & "22,,"
        
        'Processo administrativo
        If .chbProcessoAdm.Value = True Then strTopicos = strTopicos & "30,,"
        
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosFinais(form)
        
    End With
    
    ColherTopicosContGato = strTopicos

End Function

Function ColherTopicosContDebitoTerceiro(form As Variant) As String
''
'' Recolhe os tópicos para a Contestação de Débitos de terceiro. Aqui não importa a ordem (a ordem
''    dos tópicos será a que está na planilha de apoio).
''

    ' Agrupa os números dos tópicos numa string, cercados por vírgulas duplas.
    Dim strTopicos As String
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
        'Sobre as provas de não utilização do serviço
        If .chbSemProvaDeNaoUso.Value = True Then strTopicos = strTopicos & "10,,15,,"
        If .chbProvaAutorUsuario.Value = True Then strTopicos = strTopicos & "30,,"
        If .chbSemRequerimentoAdm.Value = True Then strTopicos = strTopicos & "10,,40,,20,,"
        
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosFinais(form)
        
    End With
    
    ColherTopicosContDebitoTerceiro = strTopicos

End Function

Function ColherTopicosContRespCivil(form As Variant, strCausaPedir As String) As String
''
'' Recolhe os tópicos para as Contestações de Responsabilidade civil. Aqui não importa a ordem (a ordem
''    dos tópicos será a que está na planilha de apoio).
''

    ' Agrupa os números dos tópicos numa string, cercados por vírgulas duplas.
    Dim strTopicos As String
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
        'Particularidades do caso
        Select Case strCausaPedir
        Case "Vaz. água ou extravas. esgoto com danos a patrimônio/morais"
            If .cmbOcorrencia.Value = "Fato pontual" Then
                strTopicos = strTopicos & "5,,"
            ElseIf .cmbOcorrencia.Value = "Reiterada em grande período de tempo" Then
                strTopicos = strTopicos & "10,,"
            End If
        
        Case "Acidente com pessoa/veículo em buraco"
            If .cmbOcorrencia.Value = "Veículo" Then
                strTopicos = strTopicos & "5,,"
            ElseIf .cmbOcorrencia.Value = "Pessoa" Then
                strTopicos = strTopicos & "10,,"
            End If
        
        Case "Acidente com veículo (colisão ou atropelamento)"
            If .cmbOcorrencia.Value = "Colisão" Then
                strTopicos = strTopicos & "5,,"
            ElseIf .cmbOcorrencia.Value = "Atropelamento" Then
                strTopicos = strTopicos & "10,,"
            End If
        
        End Select
        
        'Prescrição
        If .chbPrescricao.Value = True Then strTopicos = strTopicos & "15,,50,,"
        
        'Outros
        If .chbConsImpediuReparo.Value = True Then strTopicos = strTopicos & "20,,"
        If .chbOmissao.Value = True Then strTopicos = strTopicos & "25,,"
        
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosFinais(form)
        
    End With
    
    ColherTopicosContRespCivil = strTopicos

End Function

Function ColherTopicosContDesabastecimento(form As Variant) As String
''
'' Recolhe os tópicos para a Contestação de desabastecimento genérico. Aqui não importa a ordem (a ordem
''    dos tópicos será a que está na planilha de apoio).
''

    ' Agrupa os números dos tópicos numa string, cercados por vírgulas duplas.
    Dim strTopicos As String
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
        'Corresponsável
        If .cmbCorresponsavel.Value = "Não houve outro responsável" Or .cmbCorresponsavel.Value = "" Then
             strTopicos = strTopicos & "98,,"
        Else
             strTopicos = strTopicos & "96,,"
        End If
    
        'Prescrição trienal
        If .chbPrescricao.Value = True Then strTopicos = strTopicos & "15,,93,,"
        
        'Particularidades do caso
        If .chbMultiplosProcessos.Value = True Then strTopicos = strTopicos & "10,,"
        If .chbSemFatura.Value = True Then strTopicos = strTopicos & "60,,"
        If .chbBairroNaoAfetado.Value = True Then strTopicos = strTopicos & "20,,"
        If .chbInadimplenteEpoca.Value = True Then strTopicos = strTopicos & "63,,"
        If .chbReservatorio.Value = True Then strTopicos = strTopicos & "65,,"
        If .chbSemRequerimentoAdm.Value = True Then strTopicos = strTopicos & "50,,"
        
        'Alteração de consumo
        If .cmbAlteracaoConsumo.Value = "Não houve alteração relevante [gráfico]" Then
            strTopicos = strTopicos & "40,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Não houve alteração [print de contas juntadas pelo Autor]" Then
            strTopicos = strTopicos & "41,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Consumo aumentou no período do acidente [gráfico]" Then
            strTopicos = strTopicos & "42,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Abastecimento estava cortado por outro motivo" Then
            strTopicos = strTopicos & "45,,"
        End If
        
        'Tópico de danos morais adicional, específico de desabastecimento
        If form.chbDanoMoral.Value = True Then
            If .cmbCorresponsavel.Value = "Não houve outro responsável" Or .cmbCorresponsavel.Value = "" Then
                 strTopicos = strTopicos & "75,,"
            Else
                 strTopicos = strTopicos & "70,,"
            End If
        End If
        
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosFinais(form)
        
    End With
    
    ColherTopicosContDesabastecimento = strTopicos

End Function

Function ColherTopicosContFixoEsgoto(form As Variant, strCausaPedir As String) As String
''
'' Recolhe os tópicos para a Contestação das outras causas de pedir. Aqui não importa a ordem (a ordem
''    dos tópicos será a que está na planilha de apoio).
''

    ' Agrupa os números dos tópicos numa string, cercados por vírgulas duplas.
    Dim strTopicos As String
    Dim intAno As Integer
    Dim btMes As Byte
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
        'Postura do Autor
        If .cmbConfessaPoco.Value = "Confessa que existe poço" Or .cmbConfessaPoco.Value = "Confessa que foi instalado hidrômetro no poço" Then
            strTopicos = strTopicos & "5,,"
        Else
            strTopicos = strTopicos & "10,,"
        End If
        
        'Postura do Autor
        If .cmbConfessaPoco.Value = "Confessa que existe poço" Then
            strTopicos = strTopicos & "5,,"
        ElseIf .cmbConfessaPoco.Value = "Confessa que foi instalado hidrômetro no poço" Then
            strTopicos = strTopicos & "5,,15,,20,,"
        Else
            strTopicos = strTopicos & "10,,"
        End If
        
        'Necessidade de explicar volume paradigma
        If .cmbVolumeParadigma.Value = "Média anterior à instalação do poço" Then
            strTopicos = strTopicos & "35,,"
        ElseIf .cmbVolumeParadigma.Value = "Menor do que a média anterior ao poço" Then
            strTopicos = strTopicos & "40,,"
        End If
        
        'Particularidades do caso
        If .chbFotosHidrometroPoco.Value = True Then strTopicos = strTopicos & "15,,25,,"
        If .chbImpediuInstalacaoHidrometro.Value = True Then strTopicos = strTopicos & "30,,"
        If .chbNaoAplicaCDC.Value = True Then strTopicos = strTopicos & "55,,"
        
        'Descobre qual a Resolução de tarifa
        intAno = CInt(Right(.txtMesRefExemplo.Text, 4))
        btMes = CByte(Left(.txtMesRefExemplo.Text, 2))
        Select Case btMes
        Case 1, 2, 3, 4, 5, 6
            intAno = intAno - 1
        End Select
        
        Select Case intAno
        Case 2013
            strTopicos = strTopicos & "60,,"
        Case 2014
            strTopicos = strTopicos & "65,,"
        Case 2015
            strTopicos = strTopicos & "70,,"
        Case 2016
            strTopicos = strTopicos & "75,,"
        Case 2017
            strTopicos = strTopicos & "80,,"
        Case 2018
            strTopicos = strTopicos & "85,,"
        Case 2019
            strTopicos = strTopicos & "90,,"
        End Select
        
        ' Cálculo em si
        If .chbApresentarCalcExemplo.Value = True Then
            If .chbTemDoisTipos.Value = True Then
                strTopicos = strTopicos & "96,,"
            Else
                strTopicos = strTopicos & "95,,"
            End If
        End If
        
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosFinais(form)
                
    End With
    
    ColherTopicosContFixoEsgoto = strTopicos

End Function

Function ColherTopicosContVicio(form As Variant, strCausaPedir As String) As String
''
'' Recolhe os tópicos para a Contestação das outras causas de pedir. Aqui não importa a ordem (a ordem
''    dos tópicos será a que está na planilha de apoio).
''

    ' Agrupa os números dos tópicos numa string, cercados por vírgulas duplas.
    Dim strTopicos As String
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
        'Dos fatos
        Select Case strCausaPedir
        Case "Incidentes com cobrança de média"
            strTopicos = strTopicos & "5,,"
            
        Case "Desligar ligação de água"
            strTopicos = strTopicos & "10,,"
            
        Case "Hidrômetro do imóvel trocado"
            strTopicos = strTopicos & "15,,"
            
        Case Else
            strTopicos = strTopicos & "5,,"
            
        End Select
        
        'Mero vício?
        If .cbhHouveVicio.Value = True Then
            strTopicos = strTopicos & "40,,115,,"
        Else
            strTopicos = strTopicos & "45,,120,,"
        End If
        
        'Particularidades do caso
        If .chbSemRequerimentoAdm.Value = True Then strTopicos = strTopicos & "35,,50,,"
        If .chbCobrancasCanceladasAdm.Value = True Then strTopicos = strTopicos & "20,,"
        If .chbSemCorteImovelHabitado.Value = True Then strTopicos = strTopicos & "25,,"
        If .chbReclamacoesTratadasSeriedade.Value = True Then strTopicos = strTopicos & "30,,"
        If .chbLicitudeCobrancaMedia.Value = True Then strTopicos = strTopicos & "55,,60,,"
        If .chbLicitudeCobrancaMinimo.Value = True Then strTopicos = strTopicos & "60,,125,,"
        
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosFinais(form)
        
    End With
    
    ColherTopicosContVicio = strTopicos

End Function

Function ColherTopicosRIDesabastecimento(form As Variant) As String
''
'' Recolhe os tópicos para o Recurso Inominado de desabastecimento. Aqui não importa a ordem (a ordem
''    dos tópicos será a que está na planilha de apoio).
''

    ' Agrupa os números dos tópicos numa string, cercados por vírgulas duplas.
    Dim strTopicos As String
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
        'Corresponsável
        If .cmbCorresponsavel.Value = "Não houve outro responsável" Or .cmbCorresponsavel.Value = "" Then
             strTopicos = strTopicos & "98,,"
        Else
             strTopicos = strTopicos & "96,,"
        End If
    
        'Prescrição trienal
        If .chbPrescricao.Value = True Then strTopicos = strTopicos & "15,,93,,"
        
        'Particularidades do caso
        If .chbMultiplosProcessos.Value = True Then strTopicos = strTopicos & "10,,"
        If .chbSemFatura.Value = True Then strTopicos = strTopicos & "60,,"
        If .chbBairroNaoAfetado.Value = True Then strTopicos = strTopicos & "20,,"
        If .chbInadimplenteEpoca.Value = True Then strTopicos = strTopicos & "63,,"
        If .chbReservatorio.Value = True Then strTopicos = strTopicos & "65,,"
        If .chbSemRequerimentoAdm.Value = True Then strTopicos = strTopicos & "50,,"
        
        'Alteração de consumo
        If .cmbAlteracaoConsumo.Value = "Não houve alteração relevante [gráfico]" Then
            strTopicos = strTopicos & "40,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Não houve alteração [print de contas juntadas pelo Autor]" Then
            strTopicos = strTopicos & "41,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Consumo aumentou no período do acidente [gráfico]" Then
            strTopicos = strTopicos & "42,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Abastecimento estava cortado por outro motivo" Then
            strTopicos = strTopicos & "45,,"
        End If
        
        'Tópico de danos morais adicional, específico de desabastecimento
        If form.chbDanoMoral.Value = True Then
            If .cmbCorresponsavel.Value = "Não houve outro responsável" Or .cmbCorresponsavel.Value = "" Then
                 strTopicos = strTopicos & "75,,"
            Else
                 strTopicos = strTopicos & "70,,"
            End If
        End If
        
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosFinais(form)
        
        'Questões recursais
        strTopicos = Replace(strTopicos, ",,800,,", ",,") ' Remover dano material e devolução em dobro decorrentes
        strTopicos = Replace(strTopicos, ",,810,,", ",,") ' do pedido; o que contará é o da condenação, se tiver havido.

        If .chbJurosEventoDanoso.Value = True Then strTopicos = strTopicos & "330,,395,,"
        If .chbCondenouDevDobro.Value = True Then strTopicos = strTopicos & "800,,"
        If .chbCondenouDanosMateriais.Value = True Then strTopicos = strTopicos & "810,,"
        If .chbExcluiuCorresp.Value = True Then strTopicos = strTopicos & "310,,390,,"
        
        ' Testemunhas
        If .cmbTestemunhas.Value = "Autor não produziu prova testemunhal" Then
            strTopicos = strTopicos & "320,,"
        ElseIf .cmbTestemunhas.Value = "Testemunha imprecisa ou inverossímil" Then
            strTopicos = strTopicos & "323,,"
        ElseIf .cmbTestemunhas.Value = "Testemunha interessada, tem processo" Then
            strTopicos = strTopicos & "326,,"
        End If
    
    End With
    
    ColherTopicosRIDesabastecimento = strTopicos

End Function

Function ColherTopicosRICorte(form As Variant) As String
''
'' Recolhe os tópicos para o Recurso Inominado de Corte. Aqui não importa a ordem (a ordem
''    dos tópicos será a que está na planilha de apoio).
''

    ' Agrupa os números dos tópicos numa string, cercados por vírgulas duplas.
    Dim strTopicos As String
    strTopicos = ",,"
    
    
    With form
        'Ilegitimidade ativa
        If .chbIlegitimidade.Value = True Then strTopicos = strTopicos & "10,,65,,"
        
        
        'Particularidades do corte
        If .chbFaturasAberto.Value = True Then strTopicos = strTopicos & "20,,"
        If .chbInadimplente.Value = True Then strTopicos = strTopicos & "30,,"
        'If .chbContaErrada.Value = True Then strTopicos = strTopicos & "50,,"
        If .chbPagVespera.Value = True Then strTopicos = strTopicos & "40,,60,,"
        If .chbCorteBreve.Value = True Then strTopicos = strTopicos & "45,,"
            
        
        'Aviso de corte
        If .cmbAvisoCorte.Value = "Houve, em faturas anteriores / não houve" Then
            strTopicos = strTopicos & "57,,"
        ElseIf .cmbAvisoCorte.Value = "Houve, em correspondência específica" Then
            strTopicos = strTopicos & "53,,"
        End If
        
        ' Pedido (tópico especial para pagamento na véspera)
        If .chbPagVespera.Value = True Then strTopicos = strTopicos & "80,," Else strTopicos = strTopicos & "70,,"
        
        'Sentença
        If .chbJurosEventoDanoso.Value = True Then strTopicos = strTopicos & "63,,110,,"
        If .chbCondenouDanosMateriais.Value = True Then strTopicos = strTopicos & "59,,100,,"
        
    End With
    
    ColherTopicosRICorte = strTopicos

End Function

Function ColherTopicosRIDesabGenerico(form As Variant) As String
''
'' Recolhe os tópicos para o Recurso Inominado de CCR. Aqui não importa a ordem (a ordem
''    dos tópicos será a que está na planilha de apoio).
''

    ' Agrupa os números dos tópicos numa string, cercados por vírgulas duplas.
    Dim strTopicos As String
    strTopicos = ",,"
    
    
    With form
        'Corresponsável
        If .cmbCorresponsavel.Value = "Não houve outro responsável" Then
             strTopicos = strTopicos & "87,,98,,"
        Else
             strTopicos = strTopicos & "85,,96,,"
        End If
    
        'Ilegitimidade ativa
        If .chbIlegitimidade.Value = True Then strTopicos = strTopicos & "10,,90,,"
        
        'Prescrição trienal
        If .chbPrescricao.Value = True Then strTopicos = strTopicos & "17,,97,,"
        
        'Particularidades do caso
        If .chbInadimplenteEpoca.Value = True Then strTopicos = strTopicos & ",,"
        If .chbSemFatura.Value = True Then strTopicos = strTopicos & "60,,"
        If .chbBairroNaoAfetado.Value = True Then strTopicos = strTopicos & "20,,"
        If .chbReservatorio.Value = True Then strTopicos = strTopicos & "65,,"
        
        'Alteração de consumo
        If .cmbAlteracaoConsumo.Value = "Não houve alteração relevante [gráfico]" Then
            strTopicos = strTopicos & "40,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Não houve alteração [print de contas juntadas pelo Autor]" Then
            strTopicos = strTopicos & "41,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Consumo aumentou no período do acidente [gráfico]" Then
            strTopicos = strTopicos & "42,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Abastecimento estava cortado por outro motivo" Then
            strTopicos = strTopicos & "45,,"
        End If
        
        'Pedidos diferentes
        If .chbCondenouDanosMateriais.Value = True Then strTopicos = strTopicos & "70,,"
        If .chbCondenouDevDobro.Value = True Then strTopicos = strTopicos & "80,,"
        
        'Testemunha
        If .cmbTestemunhas.Value = "Autor não produziu prova testemunhal" Then
            strTopicos = strTopicos & "30,,"
        ElseIf .cmbTestemunhas.Value = "Testemunha imprecisa ou inverossímil" Then
            strTopicos = strTopicos & "33,,"
        ElseIf .cmbTestemunhas.Value = "Testemunha interessada, tem processo" Then
            strTopicos = strTopicos & "36,,"
        End If
        
        'Sentença
        If .chbExcluiuCorresp.Value = True Then strTopicos = strTopicos & "17,,97,,"
        If .chbJurosEventoDanoso.Value = True Then strTopicos = strTopicos & "63,,110,,"
        If .chbCondenouDanosMateriais.Value = True Then strTopicos = strTopicos & "70,,100,,"
        If .chbCondenouDevDobro.Value = True Then strTopicos = strTopicos & "80,,105,,"
        
    End With
    
    ColherTopicosRIDesabGenerico = strTopicos

End Function

Function ColherTopicosRICCR(form As Variant) As String
''
'' Recolhe os tópicos para o Recurso Inominado de CCR. Aqui não importa a ordem (a ordem
''    dos tópicos será a que está na planilha de apoio).
''

    ' Agrupa os números dos tópicos numa string, cercados por vírgulas duplas.
    Dim strTopicos As String
    strTopicos = ",,"
    
    
    With form
        'Ilegitimidade ativa
        If .chbIlegitimidade.Value = True Then strTopicos = strTopicos & "10,,90,,"
        
        'Prescrição trienal
        If .chbPrescricao.Value = True Then strTopicos = strTopicos & "17,,97,,"
        
        'Particularidades do caso
        If .chbInadimplenteEpoca.Value = True Then strTopicos = strTopicos & ",,"
        If .chbSemFatura.Value = True Then strTopicos = strTopicos & "60,,"
        If .chbBairroNaoAfetado.Value = True Then strTopicos = strTopicos & "20,,"
        If .chbReservatorio.Value = True Then strTopicos = strTopicos & "65,,"
        
        'Alteração de consumo
        If .cmbAlteracaoConsumo.Value = "Não houve alteração relevante [gráfico]" Then
            strTopicos = strTopicos & "40,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Não houve alteração [print de contas juntadas pelo Autor]" Then
            strTopicos = strTopicos & "41,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Consumo aumentou no período do acidente [gráfico]" Then
            strTopicos = strTopicos & "42,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Abastecimento estava cortado por outro motivo" Then
            strTopicos = strTopicos & "45,,"
        End If
        
        'Pedidos diferentes
        If .chbCondenouDanosMateriais.Value = True Then strTopicos = strTopicos & "70,,"
        If .chbCondenouDevDobro.Value = True Then strTopicos = strTopicos & "80,,"
        
        'Testemunha
        If .cmbTestemunhas.Value = "Autor não produziu prova testemunhal" Then
            strTopicos = strTopicos & "30,,"
        ElseIf .cmbTestemunhas.Value = "Testemunha imprecisa ou inverossímil" Then
            strTopicos = strTopicos & "33,,"
        ElseIf .cmbTestemunhas.Value = "Testemunha interessada, tem processo" Then
            strTopicos = strTopicos & "36,,"
        End If
        
        'Sentença
        If .chbExcluiuCorresp.Value = True Then strTopicos = strTopicos & "17,,97,,"
        If .chbJurosEventoDanoso.Value = True Then strTopicos = strTopicos & "63,,110,,"
        If .chbCondenouDanosMateriais.Value = True Then strTopicos = strTopicos & "70,,100,,"
        If .chbCondenouDevDobro.Value = True Then strTopicos = strTopicos & "80,,105,,"
        
    End With
    
    ColherTopicosRICCR = strTopicos

End Function

Function ColherTopicosRIDesabUruguai2016(form As Variant) As String
''
'' Recolhe os tópicos para o Recurso Inominado de Desabastecimento do Apagão Xingu. Aqui não importa a ordem (a ordem
''    dos tópicos será a que está na planilha de apoio).
''

    ' Agrupa os números dos tópicos numa string, cercados por vírgulas duplas.
    Dim strTopicos As String
    strTopicos = ",,"
    
    
    With form
        'Ilegitimidade ativa
        If .chbIlegitimidade.Value = True Then strTopicos = strTopicos & "10,,90,,"
        
        'Prescrição trienal
        If .chbPrescricao.Value = True Then strTopicos = strTopicos & "17,,97,,"
        
        'Particularidades do caso
        If .chbInadimplenteEpoca.Value = True Then strTopicos = strTopicos & ",,"
        If .chbSemFatura.Value = True Then strTopicos = strTopicos & "60,,"
        If .chbBairroNaoAfetado.Value = True Then strTopicos = strTopicos & "20,,"
        If .chbReservatorio.Value = True Then strTopicos = strTopicos & "65,,"
        
        'Alteração de consumo
        If .cmbAlteracaoConsumo.Value = "Não houve alteração relevante [gráfico]" Then
            strTopicos = strTopicos & "40,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Não houve alteração [print de contas juntadas pelo Autor]" Then
            strTopicos = strTopicos & "41,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Consumo aumentou no período do acidente [gráfico]" Then
            strTopicos = strTopicos & "42,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Abastecimento estava cortado por outro motivo" Then
            strTopicos = strTopicos & "45,,"
        End If
        
        'Pedidos diferentes
        If .chbCondenouDanosMateriais.Value = True Then strTopicos = strTopicos & "70,,"
        If .chbCondenouDevDobro.Value = True Then strTopicos = strTopicos & "80,,"
        
        'Testemunha
        If .cmbTestemunhas.Value = "Autor não produziu prova testemunhal" Then
            strTopicos = strTopicos & "30,,"
        ElseIf .cmbTestemunhas.Value = "Testemunha imprecisa ou inverossímil" Then
            strTopicos = strTopicos & "33,,"
        ElseIf .cmbTestemunhas.Value = "Testemunha interessada, tem processo" Then
            strTopicos = strTopicos & "36,,"
        End If
        
        'Sentença
        If .chbJurosEventoDanoso.Value = True Then strTopicos = strTopicos & "63,,110,,"
        If .chbCondenouDanosMateriais.Value = True Then strTopicos = strTopicos & "70,,100,,"
        If .chbCondenouDevDobro.Value = True Then strTopicos = strTopicos & "80,,105,,"
        
    End With
    
    ColherTopicosRIDesabUruguai2016 = strTopicos

End Function

Function ColherTopicosRIDesabLiberdade2017(form As Variant) As String
''
'' Recolhe os tópicos para o Recurso Inominado de Desabastecimento do Apagão Xingu. Aqui não importa a ordem (a ordem
''    dos tópicos será a que está na planilha de apoio).
''

    ' Agrupa os números dos tópicos numa string, cercados por vírgulas duplas.
    Dim strTopicos As String
    strTopicos = ",,"
    
    
    With form
        'Ilegitimidade ativa
        If .chbIlegitimidade.Value = True Then strTopicos = strTopicos & "10,,90,,"
        
        'Prescrição trienal
        If .chbPrescricao.Value = True Then strTopicos = strTopicos & "17,,97,,"
        
        'Particularidades do caso
        If .chbInadimplenteEpoca.Value = True Then strTopicos = strTopicos & ",,"
        If .chbSemFatura.Value = True Then strTopicos = strTopicos & "60,,"
        If .chbBairroNaoAfetado.Value = True Then strTopicos = strTopicos & "20,,"
        If .chbReservatorio.Value = True Then strTopicos = strTopicos & "65,,"
        
        'Alteração de consumo
        If .cmbAlteracaoConsumo.Value = "Não houve alteração relevante [gráfico]" Then
            strTopicos = strTopicos & "40,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Não houve alteração [print de contas juntadas pelo Autor]" Then
            strTopicos = strTopicos & "41,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Consumo aumentou no período do acidente [gráfico]" Then
            strTopicos = strTopicos & "42,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Abastecimento estava cortado por outro motivo" Then
            strTopicos = strTopicos & "45,,"
        End If
        
        'Pedidos diferentes
        If .chbCondenouDanosMateriais.Value = True Then strTopicos = strTopicos & "70,,"
        If .chbCondenouDevDobro.Value = True Then strTopicos = strTopicos & "80,,"
        
        'Testemunha
        If .cmbTestemunhas.Value = "Autor não produziu prova testemunhal" Then
            strTopicos = strTopicos & "30,,"
        ElseIf .cmbTestemunhas.Value = "Testemunha imprecisa ou inverossímil" Then
            strTopicos = strTopicos & "33,,"
        ElseIf .cmbTestemunhas.Value = "Testemunha interessada, tem processo" Then
            strTopicos = strTopicos & "36,,"
        End If
        
        'Sentença
        If .chbJurosEventoDanoso.Value = True Then strTopicos = strTopicos & "63,,110,,"
        If .chbCondenouDanosMateriais.Value = True Then strTopicos = strTopicos & "70,,100,,"
        If .chbCondenouDevDobro.Value = True Then strTopicos = strTopicos & "80,,105,,"
        
    End With
    
    ColherTopicosRIDesabLiberdade2017 = strTopicos

End Function

Function ColherTopicosRIDesabApagXingu2018(form As Variant) As String
''
'' Recolhe os tópicos para o Recurso Inominado de Desabastecimento do Apagão Xingu. Aqui não importa a ordem (a ordem
''    dos tópicos será a que está na planilha de apoio).
''

    ' Agrupa os números dos tópicos numa string, cercados por vírgulas duplas.
    Dim strTopicos As String
    strTopicos = ",,"
    
    
    With form
        'Ilegitimidade ativa
        If .chbIlegitimidade.Value = True Then strTopicos = strTopicos & "10,,90,,"
        
        'Prescrição trienal
        If .chbPrescricao.Value = True Then strTopicos = strTopicos & "17,,97,,"
        
        'Particularidades do caso
        If .chbInadimplenteEpoca.Value = True Then strTopicos = strTopicos & ",,"
        If .chbSemFatura.Value = True Then strTopicos = strTopicos & "60,,"
        If .chbBairroNaoAfetado.Value = True Then strTopicos = strTopicos & "20,,"
        If .chbReservatorio.Value = True Then strTopicos = strTopicos & "65,,"
        
        'Alteração de consumo
        If .cmbAlteracaoConsumo.Value = "Não houve alteração relevante [gráfico]" Then
            strTopicos = strTopicos & "40,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Não houve alteração [print de contas juntadas pelo Autor]" Then
            strTopicos = strTopicos & "41,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Consumo aumentou no período do acidente [gráfico]" Then
            strTopicos = strTopicos & "42,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Abastecimento estava cortado por outro motivo" Then
            strTopicos = strTopicos & "45,,"
        End If
        
        'Pedidos diferentes
        If .chbCondenouDanosMateriais.Value = True Then strTopicos = strTopicos & "70,,"
        If .chbCondenouDevDobro.Value = True Then strTopicos = strTopicos & "80,,"
        
        'Testemunha
        If .cmbTestemunhas.Value = "Autor não produziu prova testemunhal" Then
            strTopicos = strTopicos & "30,,"
        ElseIf .cmbTestemunhas.Value = "Testemunha imprecisa ou inverossímil" Then
            strTopicos = strTopicos & "33,,"
        ElseIf .cmbTestemunhas.Value = "Testemunha interessada, tem processo" Then
            strTopicos = strTopicos & "36,,"
        End If
        
        'Sentença
        If .chbExcluiuCorresp.Value = True Then strTopicos = strTopicos & "15,,95,,"
        If .chbJurosEventoDanoso.Value = True Then strTopicos = strTopicos & "63,,110,,"
        If .chbCondenouDanosMateriais.Value = True Then strTopicos = strTopicos & "70,,100,,"
        If .chbCondenouDevDobro.Value = True Then strTopicos = strTopicos & "80,,105,,"
        
    End With
    
    ColherTopicosRIDesabApagXingu2018 = strTopicos

End Function

Function ColherTopicosCRRIDesabGenerico(form As Variant) As String
''
'' Recolhe os tópicos para as Contrarrazões de Recurso Inominado de CCR. Aqui não importa a ordem (a ordem
''    dos tópicos será a que está na planilha de apoio).
''

    ' Agrupa os números dos tópicos numa string, cercados por vírgulas duplas.
    Dim strTopicos As String
    strTopicos = ",,"
    
    
    With form
        'Corresponsável
        If .cmbCorresponsavel.Value = "Não houve outro responsável" Then
             strTopicos = strTopicos & "87,,98,,"
        Else
             strTopicos = strTopicos & "85,,96,,"
        End If
    
        'Ilegitimidade ativa
        If .chbIlegitimidade.Value = True Then strTopicos = strTopicos & "10,,"
        
        'Prescrição trienal
        If .chbPrescricao.Value = True Then strTopicos = strTopicos & "15,,"
        
        'Particularidades do caso
        If .chbInadimplenteEpoca.Value = True Then strTopicos = strTopicos & ",,"
        If .chbSemFatura.Value = True Then strTopicos = strTopicos & "60,,"
        If .chbBairroNaoAfetado.Value = True Then strTopicos = strTopicos & "20,,"
        If .chbReservatorio.Value = True Then strTopicos = strTopicos & "65,,"
        
        'Alteração de consumo
        If .cmbAlteracaoConsumo.Value = "Não houve alteração relevante [gráfico]" Then
            strTopicos = strTopicos & "40,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Não houve alteração [print de contas juntadas pelo Autor]" Then
            strTopicos = strTopicos & "41,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Consumo aumentou no período do acidente [gráfico]" Then
            strTopicos = strTopicos & "42,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Abastecimento estava cortado por outro motivo" Then
            strTopicos = strTopicos & "45,,"
        End If
        
        'Sentença
        If .optImprocedente.Value = True Then strTopicos = strTopicos & "3,,90,,"
        If .optProcedenteEmParte.Value = True Then strTopicos = strTopicos & "6,,95,,"
        
        'Recurso Inominado do Adverso
        If .chbRecorreuJurosEventoDanoso.Value = True Then strTopicos = strTopicos & "63,,"
        If .chbRecorreuDanosMateriais.Value = True Then strTopicos = strTopicos & "70,,"
        If .chbRecorreuDevDobro.Value = True Then strTopicos = strTopicos & "80,,"

    End With
    
    ColherTopicosCRRIDesabGenerico = strTopicos

End Function

Function ColherTopicosCRRICCR(form As Variant) As String
''
'' Recolhe os tópicos para as Contrarrazões de Recurso Inominado de CCR. Aqui não importa a ordem (a ordem
''    dos tópicos será a que está na planilha de apoio).
''

    ' Agrupa os números dos tópicos numa string, cercados por vírgulas duplas.
    Dim strTopicos As String
    strTopicos = ",,"
    
    
    With form
        'Ilegitimidade ativa
        If .chbIlegitimidade.Value = True Then strTopicos = strTopicos & "10,,"
        
        'Prescrição trienal
        If .chbPrescricao.Value = True Then strTopicos = strTopicos & "15,,"
        
        'Particularidades do caso
        If .chbInadimplenteEpoca.Value = True Then strTopicos = strTopicos & ",,"
        If .chbSemFatura.Value = True Then strTopicos = strTopicos & "60,,"
        If .chbBairroNaoAfetado.Value = True Then strTopicos = strTopicos & "20,,"
        If .chbReservatorio.Value = True Then strTopicos = strTopicos & "65,,"
        
        'Alteração de consumo
        If .cmbAlteracaoConsumo.Value = "Não houve alteração relevante [gráfico]" Then
            strTopicos = strTopicos & "40,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Não houve alteração [print de contas juntadas pelo Autor]" Then
            strTopicos = strTopicos & "41,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Consumo aumentou no período do acidente [gráfico]" Then
            strTopicos = strTopicos & "42,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Abastecimento estava cortado por outro motivo" Then
            strTopicos = strTopicos & "45,,"
        End If
        
        'Sentença
        If .optImprocedente.Value = True Then strTopicos = strTopicos & "3,,90,,"
        If .optProcedenteEmParte.Value = True Then strTopicos = strTopicos & "6,,95,,"
        
        'Recurso Inominado do Adverso
        If .chbRecorreuJurosEventoDanoso.Value = True Then strTopicos = strTopicos & "63,,"
        If .chbRecorreuDanosMateriais.Value = True Then strTopicos = strTopicos & "70,,"
        If .chbRecorreuDevDobro.Value = True Then strTopicos = strTopicos & "80,,"

    End With
    
    ColherTopicosCRRICCR = strTopicos

End Function

Function ColherTopicosCRRIDesabApagXingu2018(form As Variant) As String
''
'' Recolhe os tópicos para as Contrarrazões de Recurso Inominado de CCR. Aqui não importa a ordem (a ordem
''    dos tópicos será a que está na planilha de apoio).
''

    ' Agrupa os números dos tópicos numa string, cercados por vírgulas duplas.
    Dim strTopicos As String
    strTopicos = ",,"
    
    
    With form
        'Ilegitimidade ativa
        If .chbIlegitimidade.Value = True Then strTopicos = strTopicos & "10,,"
        
        'Prescrição trienal
        If .chbPrescricao.Value = True Then strTopicos = strTopicos & "15,,"
        
        'Particularidades do caso
        If .chbInadimplenteEpoca.Value = True Then strTopicos = strTopicos & ",,"
        If .chbSemFatura.Value = True Then strTopicos = strTopicos & "60,,"
        If .chbBairroNaoAfetado.Value = True Then strTopicos = strTopicos & "20,,"
        If .chbReservatorio.Value = True Then strTopicos = strTopicos & "65,,"
        
        'Alteração de consumo
        If .cmbAlteracaoConsumo.Value = "Não houve alteração relevante [gráfico]" Then
            strTopicos = strTopicos & "40,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Não houve alteração [print de contas juntadas pelo Autor]" Then
            strTopicos = strTopicos & "41,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Consumo aumentou no período do acidente [gráfico]" Then
            strTopicos = strTopicos & "42,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Abastecimento estava cortado por outro motivo" Then
            strTopicos = strTopicos & "45,,"
        End If
        
        'Sentença
        If .optImprocedente.Value = True Then strTopicos = strTopicos & "3,,85,,"
        If .optProcedenteEmParte.Value = True Then strTopicos = strTopicos & "6,,87,,"
        
        'Recurso Inominado do Adverso
        If .chbRecorreuJurosEventoDanoso.Value = True Then strTopicos = strTopicos & "63,,"
        If .chbRecorreuDanosMateriais.Value = True Then strTopicos = strTopicos & "70,,"
        If .chbRecorreuDevDobro.Value = True Then strTopicos = strTopicos & "80,,"

    End With
    
    ColherTopicosCRRIDesabApagXingu2018 = strTopicos

End Function

Sub MontarContestacao(control As IRibbonControl)
''
'' Monta a Contestação.
''
    
    MontarPeticaoComplexa ActiveSheet.Cells(ActiveCell.Row, 5).Formula, "Contestação", ActiveSheet.Cells(ActiveCell.Row, 11).Formula, Format(ActiveSheet.Cells(ActiveCell.Row, 9).Formula, "dd/mm/yyyy")
    
End Sub

Sub MontarRIApelacao(control As IRibbonControl)
''
'' Monta o RI.
''
    
    MontarPeticaoComplexa ActiveSheet.Cells(ActiveCell.Row, 5).Formula, "RI/Apelação", ActiveSheet.Cells(ActiveCell.Row, 11).Formula, Format(ActiveSheet.Cells(ActiveCell.Row, 9).Formula, "dd/mm/yyyy")
    
End Sub

Sub MontarContrarrazoesRIApelacao(control As IRibbonControl)
''
'' Monta as Contrarrazões do RI.
''
    
    MontarPeticaoComplexa ActiveSheet.Cells(ActiveCell.Row, 5).Formula, "Contrarrazões de RI/Apelação", ActiveSheet.Cells(ActiveCell.Row, 11).Formula, Format(ActiveSheet.Cells(ActiveCell.Row, 9).Formula, "dd/mm/yyyy")
    
End Sub

Sub DetectarProvMontarPet(control As IRibbonControl)
''
'' Verifica a providência e chama a função correspondente.
''
    Dim strProvidencia As String
    Dim plan As Worksheet
    Dim contLinha As Long
    
    Set plan = ActiveSheet
    contLinha = ActiveCell.Row
    
    strProvidencia = plan.Cells(contLinha, 4).Text
    
    Select Case strProvidencia
    Case "Contestar", "Contestar - Remarcação de audiência"
        MontarContestacao control
        
    Case "Recorrer"
        MontarRIApelacao control
        
    Case "Contra-arrazoar recurso"
        MontarContrarrazoesRIApelacao control
    
    Case Else
        MsgBox "Sinto muito, " & DeterminarTratamento & "! O comando escolhido só serve para Contestar, Recorrer ou Contra-arrazoar. " & _
                "Será que um dos outros botões não vos satisfaria?", vbInformation + vbOKOnly, "Sísifo em treinamento"
    
    End Select

    
    
End Sub

