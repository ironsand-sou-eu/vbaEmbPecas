Attribute VB_Name = "modCriarPeticoesComplexas"
Option Explicit

Sub MontarPeticaoComplexa(strCausaPedir As String, strPeticao As String, strJuizoEspaider As String, strTermoInicialPrazo As String)
''
'' Monta uma peti��o complexa. strCausaPedir aceita qualquer valor das causas de pedir do Espaider.
''  strPeticao aceita os valores "Contesta��o", "Recurso Inominado" ou "Contrarraz�es de RI".
''  strJuizoEspaider aceita valores que contenham as express�es "Juizado" ou ???
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
    
    ' Estabelece qual o �rg�o em que tramita o processo. Se houver a express�o "Juizado", considera que � juizado. Se houver "procuradoria",
    '  "coordenadoria", "coordena��o", "Procon" ou "Codecon", considera que � Procon. Em qualquer outro caso, considera que � Vara c�vel.
    '   N�o diferencia mai�sculas de min�sculas.
    If InStr(1, LCase(strJuizoEspaider), "juizado") <> 0 Or InStr(1, LCase(strJuizoEspaider), "sje") <> 0 Then
        strOrgao = "JEC"
    
    ElseIf InStr(1, LCase(strJuizoEspaider), "procuradoria") <> 0 Or InStr(1, LCase(strJuizoEspaider), "coordenadoria") <> 0 Or _
            InStr(1, LCase(strJuizoEspaider), "coordena��o") <> 0 Or InStr(1, LCase(strJuizoEspaider), "Procon") <> 0 Or _
            InStr(1, LCase(strJuizoEspaider), "Codecon") <> 0 Then
        strOrgao = "Procon"
        
    Else
        strOrgao = "VC"
    
    End If
    
    
    '''''''''''''''''''''''''''''''''''''''''
    '' Define as causas de pedir-paradigma ''
    '''''''''''''''''''''''''''''''''''''''''
    
    Select Case strCausaPedir
    ' Primeiro, os casos em que n�o muda, porque h� descri��o espec�fica da pe�a
    Case "Revis�o de consumo elevado", "Corte no fornecimento", "Negativa��o no SPC", "Realizar liga��o de �gua", _
        "Cobran�a de esgoto em im�vel n�o ligado � rede", "Cobran�a de esgoto com �gua cortada", "Classifica��o tarifa ou qtd. de economias", _
        "Suspeita de by-pass", "D�bito de terceiro", "Desabastecimentos por per�odo e causa determinados", "Desabastecimento CCR 04/2015", _
        "Desabastecimento Uruguai 09/2016", "Desabastecimento Liberdade 10/2017", "Desabastecimento Apag�o Xingu 03/2018", _
        "Vaz. �gua ou extravas. esgoto com danos a patrim�nio/morais", "Obra da Embasa com danos a patrim�nio/morais", _
        "Acidente com pessoa/ve�culo em buraco", "Acidente com ve�culo (colis�o ou atropelamento)", "Fixo de esgoto"
        strCausaPedirParadigma = strCausaPedir
        
    Case "Consumo elevado com corte"
        strCausaPedirParadigma = "Revis�o de consumo elevado"
    
    Case "Desmembramento de liga��es"
        strCausaPedirParadigma = "Realizar liga��o de �gua"
    
    Case "Multa por infra��o"
        strCausaPedirParadigma = "Suspeita de by-pass"
    
    Case "Desabastecimento CCR SEGUNDO ACIDENTE 04/2016", "Desabastecimento Faz. Grande do Retiro Ver�o 2019", _
        "Desabastecimento Nova Bras�lia de Itapu� 10/2018", "Desabastecimento Novo Horizonte v�rios per�odos", "Desabastecimento Pernambu�s 11/2018", _
        "Desabastecimento Sub�rbio Ferrovi�rio 02/2017", "Desabastecimento Sussuarana 12/2018", "Desabastecimento Res. Bosque das Brom�lias 10/2019", "Irregularidade no abastecimento de �gua"
        strCausaPedirParadigma = "Desabastecimentos por per�odo e causa determinados"
    
    Case Else
        strCausaPedirParadigma = "Outros v�cios"
        
    End Select

    
    
    '''''''''''''''''''''''''
    '' Mostra o formul�rio ''
    '''''''''''''''''''''''''
    
    Set form = ConfigurarFormulario(strCausaPedirParadigma, strOrgao, strPeticao)
    If form Is Nothing Then Exit Sub
    
    If strCausaPedir = "Consumo elevado com corte" Then
        form.chbDanMorCorte.Value = True
    End If
    
    form.Show
    
    If form.chbDeveGerar.Value = False Then Exit Sub
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' Colhe os t�picos, na vari�vel strTopicos; tamb�m ajusta bolGrafico e btTabela, se houve gr�ficos ou tabelas ''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    strTopicos = ColherTopicosGeral(strCausaPedir, strCausaPedirParadigma, strPeticao, form, bolGrafico, btTabela)
    
    
    ''''''''''''''''''''''''''''''
    '' Colhe outras informa��es ''
    ''''''''''''''''''''''''''''''
    
    ' Pega o Ju�zo na reda��o do Espaider, depois pega o ju�zo na reda��o para ficar na planilha.
    strJuizo = BuscaJuizo(strJuizoEspaider) ' strJuizo assume a reda��o longa do ju�zo
    If strPeticao = "RI/Apela��o" Or strPeticao = "Contrarraz�es de RI/Apela��o" Then strJuizoResumido = BuscaJuizo(strJuizo) 'strJuizoResumido assume a reda��o curta do ju�zo
    
    ' Criar o documento a partir do modelo
    Set appword = New Word.Application
    appword.Visible = True
    Set wdDocPeticao = appword.Documents.Add(BuscarCaminhoPrograma & "modelos-automaticos\PPJCM Modelo.dotx")
    
    '''''''''''''''''''''''''''''
    '' Pega a planilha correta ''
    '''''''''''''''''''''''''''''
    
    Set plan = DescobrirPlanilhaDeEstrutura(strPeticao, strOrgao)
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' Copiar os t�picos selecionados para o documento-destino ''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ' Na planilha adequada, busca o cabe�alho de t�picos da peti��o e descobre qual a primeira linha da coluna n�m-t�pico
    
    Set rngCont = plan.Cells().Find(what:=strCausaPedirParadigma, lookat:=xlWhole, searchorder:=xlByRows, MatchCase:=False).Offset(1, 0).Offset(0, 3)
    
    ' Desce de c�lula em c�lula, copiando e colando cada vez que o valor da c�lula for zero ou estiver contido em strTopicos (cercado por duas v�rgulas)
    
    Do Until rngCont.Offset(1, 0).Text = ""
        Set rngCont = rngCont.Offset(1, 0)
        If rngCont.Formula = "0" Or InStr(1, strTopicos, ",," & rngCont.Formula & ",,") Then
            
            ' Pega o nome � direita, procura na pasta de t�picos personalizados e, n�o havendo, na Frankenstein normal.
            If Dir(BuscarCaminhoPrograma & "Frankenstein\T�picos personalizados\" & rngCont.Offset(0, 1).Formula) <> "" Then
                strCaminhoDocOrigem = BuscarCaminhoPrograma & "Frankenstein\T�picos personalizados\" & rngCont.Offset(0, 1).Formula
            Else
                strCaminhoDocOrigem = BuscarCaminhoPrograma & "Frankenstein\" & rngCont.Offset(0, 1).Formula
            End If
            
            ' Adiciona o documento respectivo ao arquivo Word principal (na ordem das linhas da planilha "plan").
            InserirArquivo strCaminhoDocOrigem, wdDocPeticao
            
            ' Se houver conte�do duas colunas � direita, armazenar numa array de vari�veis.
            If Trim(rngCont.Offset(0, 2).Formula) <> "" Then
                strVariaveis = strVariaveis & rngCont.Offset(0, 2).Formula & ","
            End If
        End If
    Loop
    
    wdDocPeticao.Activate
    wdDocPeticao.Paragraphs.Last.Range.Delete 'Apaga o par�grafo vazio que fica no final.
    
    ' Substituir as vari�veis Comarca, N�mero e Adverso.
    
    With appword.Selection.Find
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .Text = "<Ju�zo>"
        .Replacement.Text = strJuizo
        .Execute Replace:=wdReplaceAll
        'Se o ju�zo n�o foi encontrado no S�sifo, avisa.
        If strJuizo = "" Then MsgBox "Ju�zo n�o encontrado na base de dados no S�sifo. Lembre-se de acrescent�-lo manualmente no endere�amento da peti��o.", _
            vbCritical + vbOKOnly, "Alerta - ju�zo n�o encontrado"

        If strPeticao = "RI/Apela��o" Or strPeticao = "Contrarraz�es de RI/Apela��o" Then
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .Text = "<Ju�zo-Resumido>"
            .Replacement.Text = strJuizoResumido
            .Execute Replace:=wdReplaceAll
            'Se o ju�zo n�o foi encontrado no S�sifo, avisa.
            If strJuizoResumido = "" Then MsgBox "Nome resumido do ju�zo n�o encontrado na base de dados no S�sifo. Lembre-se de acrescent�-lo manualmente no endere�amento da peti��o.", _
                vbCritical + vbOKOnly, "Alerta - ju�zo resumido n�o encontrado"
        End If

        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .Text = "<N�mero>"
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
    '' Substituir as demais vari�veis.''
    ''''''''''''''''''''''''''''''''''''
    
    strVariaveis = RemoverDuplicadosArray(strVariaveis, ",")
    arrVariaveis = Split(strVariaveis, ",")
    
    For Each z In arrVariaveis
        If z = "data-audi�ncia-concilia��o" Then
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
    
    ' Substituir os gr�ficos, se houver
    If bolGrafico Then
        On Error Resume Next
        Set varCont = Application.InputBox(Prompt:=DeterminarTratamento & ", no Excel, clique em qualquer c�lula na planilha que cont�m o gr�fico de consumo, o qual ser� inserido na peti��o.", Title:="S�sifo - Selecione o gr�fico!", Type:=8)
        
        If Err.Number <> 0 Or varCont.Worksheet.ChartObjects.Count = 0 Then
            MsgBox DeterminarTratamento & ", n�o escolheste uma celula numa planilha com gr�fico. A peti��o ser� gerada sem o gr�fico; lembre-se de Adicionar " & _
                "um gr�fico ou tela.", vbCritical + vbOKOnly, "S�sifo - Peti��o gerada sem gr�fico"
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
                .Text = "<gr�fico-de-consumo>^p"
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
    
    ' Substituir as tabelas de m�dia, se houver
    If btTabela <> 0 Then
        Set rngCont = Application.InputBox(Prompt:=DeterminarTratamento & ", selecione as c�lulas para a tabela de m�dias (12 linhas e 2 colunas, SEM cabe�alho).", Title:="S�sifo - Selecione a tabela de m�dia!", Type:=8)
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
    
    
    ' Simplifica o nome para n�o ficar muito grande
    Select Case strCausaPedir
    Case "Desabastecimentos por per�odo e causa determinados"
        strCausaPedir = "Desabastecimentos"
    Case "Cobran�a de esgoto em im�vel n�o ligado � rede"
        strCausaPedir = "Esgoto sem rede"
    Case "Cobran�a de esgoto com �gua cortada"
        strCausaPedir = "Esgoto com �gua cortada"
    Case "Classifica��o tarifa ou qtd. de economias"
        strCausaPedir = "Economias"
    Case "Suspeita de by-pass", "Multa por infra��o"
        strCausaPedir = "Gato"
    Case "Vaz. �gua ou extravas. esgoto com danos a patrim�nio/morais"
        strCausaPedir = "Extravasamento"
    Case "Obra da Embasa com danos a patrim�nio/morais"
        strCausaPedir = "Obra"
    Case "Acidente com pessoa/ve�culo em buraco"
        strCausaPedir = "Acidente em buraco"
    Case "Acidente com ve�culo (colis�o ou atropelamento)"
        strCausaPedir = "Acidente com ve�culo"
    End Select
    
    ' Se tiver /, ajusta o nome (por exemplo, na CCR e outros desabastecimentos que t�m data, pois o windows n�o aceita nome de arquivo com "/")
    strCausaPedir = Replace(strCausaPedir, "/", ".")
    
    ' Ajusta o nome pelo tipo de peti��o e pelo �rg�o em que tramita
    Select Case strPeticao
    Case "Contesta��o"
        If strOrgao = "JEC" Or strOrgao = "VC" Then strPeticao = "Contestacao"
        If strOrgao = "Procon" Then strPeticao = "Impugnacao"
        
    Case "RI/Apela��o"
        Select Case strOrgao
        Case "JEC"
            strPeticao = "Recurso Inominado"
        
        Case "VC"
            strPeticao = "Apelacao"
        
        'Case "Procon"
        '    strPeticao = "Recurso Administrativo"
        
        End Select
    
    Case "Contrarraz�es de RI/Apela��o"
        Select Case strOrgao
        Case "JEC"
            strPeticao = "Contrarrazoes RI"
        
        Case "VC"
            strPeticao = "Contrarrazoes Apelacao"
        
        'Case "Procon"
        '    strPeticao = "Recurso Administrativo"
        
        End Select
        
    End Select
    
    ' Salvar o documento, ir para o in�cio e exibir
    If bolSsfPrazosBotaoPdfPressionado Then 'Gerar como PDF
        wdDocPeticao.ExportAsFixedFormat OutputFilename:=BuscarCaminhoPrograma & "01 " & strPeticao & " " & strCausaPedir & " - " & SeparaPrimeirosNomes(ActiveSheet.Cells(ActiveCell.Row, 2).Formula, 2) & ".pdf", _
            ExportFormat:=wdExportFormatPDF, OptimizeFor:=wdExportOptimizeForOnScreen, CreateBookmarks:=wdExportCreateHeadingBookmarks, BitmapMissingFonts:=False
        MsgBox DeterminarTratamento & ", o PDF foi gerado com sucesso.", vbInformation + vbOKOnly, "S�sifo - Documento gerado"
        wdDocPeticao.Close wdDoNotSaveChanges
        If appword.Documents.Count = 0 Then appword.Quit
        
    Else ' Gerar como Word
        wdDocPeticao.SaveAs BuscarCaminhoPrograma & Format(Date, "yyyy.mm.dd") & " - " & strPeticao & " " & strCausaPedir & " - " & SeparaPrimeirosNomes(ActiveSheet.Cells(ActiveCell.Row, 2).Formula, 2) & ".docx"
        wdDocPeticao.GoTo wdGoToPage, wdGoToFirst
        'wdDocPeticao.Activate
        'appword.Activate
        MsgBox DeterminarTratamento & ", o documento Word foi gerado com sucesso.", vbInformation + vbOKOnly, "S�sifo - Documento gerado"
    End If
    
End Sub

Function ConfigurarFormulario(strCausaPedir As String, strOrgao As String, strPeticao As String) As Variant
''
'' Configura os formul�rios a serem exibidos.
''

Dim form As Variant
    
Select Case strCausaPedir
Case "Revis�o de consumo elevado"
    Select Case strPeticao
    Case "Contesta��o"
        Set form = New frmContRevConsElevado
        With form
            .cmbAferHidrometro.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Aferi��odehidr�metro").Address
            .cmbPadraoConsumo.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Padr�odeconsumo").Address
            .cmbVazInterno.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Vazamentointerno").Address
            If strOrgao = "VC" Then .chbRequererPericia.Visible = True
            If strOrgao = "Procon" Then
                .chbDanMorCorte.Value = False
                .chbDanMorCorte.Visible = False
                .chbDanoMoral.Value = False
                .chbDanoMoral.Visible = False
            End If
        End With
    
    Case "RI/Apela��o"
        GoTo N�oFaz
        'Set form = New
        
    Case "Contrarraz�es de RI/Apela��o"
        GoTo N�oFaz
        'Set form = New
        
    End Select
    
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Devolu��oemdobro").Address
    
Case "Corte no fornecimento"
    Select Case strPeticao
    Case "Contesta��o"
        Set form = New frmContCorte
        
    Case "RI/Apela��o"
        GoTo N�oFaz
        'Set form = New frmRICorte
        
    Case "Contrarraz�es de RI/Apela��o"
        GoTo N�oFaz
        'Set form = New
        
    End Select
    
    form.cmbAvisoCorte.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("AvisoCorte").Address
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Devolu��oemdobro").Address
    
Case "Realizar liga��o de �gua"
    Select Case strPeticao
    Case "Contesta��o"
        Set form = New frmContFazerLigacao
    
    Case "RI/Apela��o"
        GoTo N�oFaz
        'Set form = New
        
    Case "Contrarraz�es de RI/Apela��o"
        GoTo N�oFaz
        'Set form = New
        
    End Select
    
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Devolu��oemdobro").Address
    
Case "Negativa��o no SPC"
    Select Case strPeticao
    Case "Contesta��o"
        Set form = New frmContNegativacao
        form.cmbAtitudeAutor.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("AtitudeAutorNegativa��o").Address
        form.cmbPerfilContrato.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("PerfilContrato").Address
    
    Case "RI/Apela��o"
        GoTo N�oFaz
        'Set form = New
        
    Case "Contrarraz�es de RI/Apela��o"
        GoTo N�oFaz
        'Set form = New
        
    End Select
    
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Devolu��oemdobro").Address
    
Case "Classifica��o tarifa ou qtd. de economias"
    Select Case strPeticao
    Case "Contesta��o"
        Set form = New frmContClasTarif
        form.cmbPorcentEsgoto.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("TarifaEsgotoPercent").Address
    
    Case "RI/Apela��o"
        GoTo N�oFaz
        'Set form = New frmRIDesabGenerico
        
    Case "Contrarraz�es de RI/Apela��o"
        GoTo N�oFaz
        'Set form = New frmCRRIDesabGenerico
        
    End Select
    
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Devolu��oemdobro").Address
    
Case "Suspeita de by-pass", "Multa por infra��o"
    Select Case strPeticao
    Case "Contesta��o"
        Set form = New frmContGato
    
    Case "RI/Apela��o"
        GoTo N�oFaz
        'Set form = New frmRIDesabGenerico
        
    Case "Contrarraz�es de RI/Apela��o"
        GoTo N�oFaz
        'Set form = New frmCRRIDesabGenerico
        
    End Select
    
    form.cmbTipoGato.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("TipoGato").Address
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Devolu��oemdobro").Address
    
Case "D�bito de terceiro"
    Select Case strPeticao
    Case "Contesta��o"
        Set form = New frmContDebitoTerceiro
    
    Case "RI/Apela��o"
        GoTo N�oFaz
        'Set form = New frmRIDesabGenerico
        
    Case "Contrarraz�es de RI/Apela��o"
        GoTo N�oFaz
        'Set form = New frmCRRIDesabGenerico
        
    End Select
    
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Devolu��oemdobro").Address
    
Case "Cobran�a de esgoto em im�vel n�o ligado � rede"
    Select Case strPeticao
    Case "Contesta��o"
        Set form = New frmContEsgotoSemRede
        
    Case "RI/Apela��o"
        GoTo N�oFaz
        'Set form = New frmRIDesabGenerico
        
    Case "Contrarraz�es de RI/Apela��o"
        GoTo N�oFaz
        'Set form = New frmCRRIDesabGenerico
        
    End Select
        
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Devolu��oemdobro").Address
    
Case "Cobran�a de esgoto com �gua cortada"
    Select Case strPeticao
    Case "Contesta��o"
        Set form = New frmContEsgotoAguaCortada
        
    Case "RI/Apela��o"
        GoTo N�oFaz
        'Set form = New frmRIDesabGenerico
        
    Case "Contrarraz�es de RI/Apela��o"
        GoTo N�oFaz
        'Set form = New frmCRRIDesabGenerico
        
    End Select
    
    form.cmbAtitudeAutor.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("AtitudeAutorEsgoto�guaCortada").Address
    form.cmbProvaUsoImovel.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("ProvaUsoIm�vel").Address
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Devolu��oemdobro").Address
    
Case "Vaz. �gua ou extravas. esgoto com danos a patrim�nio/morais"
    Select Case strPeticao
    Case "Contesta��o"
        Set form = New frmContRespCivil
        
    Case "RI/Apela��o"
        GoTo N�oFaz
        'Set form = New frmRICorte
        
    Case "Contrarraz�es de RI/Apela��o"
        GoTo N�oFaz
        'Set form = New frmCRRIDesabGenerico
        
    End Select
    
    form.cmbOcorrencia.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Recorr�ncia").Address
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Devolu��oemdobro").Address
    
Case "Obra da Embasa com danos a patrim�nio/morais"
    Select Case strPeticao
    Case "Contesta��o"
        Set form = New frmContRespCivil
        
    Case "RI/Apela��o"
        GoTo N�oFaz
        'Set form = New frmRICorte
        
    Case "Contrarraz�es de RI/Apela��o"
        GoTo N�oFaz
        'Set form = New frmCRRIDesabGenerico
        
    End Select
    
    form.cmbOcorrencia.Enabled = False
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Devolu��oemdobro").Address
    
Case "Acidente com pessoa/ve�culo em buraco"
    Select Case strPeticao
    Case "Contesta��o"
        Set form = New frmContRespCivil
        
    Case "RI/Apela��o"
        GoTo N�oFaz
        'Set form = New frmRICorte
        
    Case "Contrarraz�es de RI/Apela��o"
        GoTo N�oFaz
        'Set form = New frmCRRIDesabGenerico
        
    End Select
    
    form.cmbOcorrencia.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("AcidenteBuraco").Address
    form.cmbOcorrencia.Text = "Ve�culo"
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Devolu��oemdobro").Address
    
Case "Acidente com ve�culo (colis�o ou atropelamento)"
    Select Case strPeticao
    Case "Contesta��o"
        Set form = New frmContRespCivil
        
    Case "RI/Apela��o"
        GoTo N�oFaz
        'Set form = New frmRICorte
        
    Case "Contrarraz�es de RI/Apela��o"
        GoTo N�oFaz
        'Set form = New frmCRRIDesabGenerico
        
    End Select
    
    form.cmbOcorrencia.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("ColisaoAtropelamento").Address
    form.cmbOcorrencia.Text = "Ve�culo"
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Devolu��oemdobro").Address
    
Case "Desabastecimentos por per�odo e causa determinados"
    Select Case strPeticao
    Case "Contesta��o"
        Set form = New frmContDesabastecimento
        
    Case "RI/Apela��o"
        Set form = New frmRIDesabastecimento
        form.cmbTestemunhas.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Testemunhas").Address
        
    Case "Contrarraz�es de RI/Apela��o"
        GoTo N�oFaz
        'Set form = New frmCRRIDesabGenerico
        'form.cmbTestemunhas.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Testemunhas").Address
        
    End Select
    
    With form
        .frmDesabastecimento.Caption = "Desabastecimentos por per�odo e causa determinados - Gen�rico"
        .chbPrescricao.Caption = "Ajuizamento posterior a (prescri��o trienal)"
        .chbSemFatura.Caption = "Autor n�o juntou conta do m�s em quest�o"
        .cmbAlteracaoConsumo.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Altera��oConsumoCCR").Address
        .cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Devolu��oemdobro").Address
        .cmbCorresponsavel.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Correspons�veis").Address
        .chbDanMorCorte.Value = False
        .chbDanMorCorte.Enabled = False
        .chbDanMorMeraCobranca.Value = False
        .chbDanMorMeraCobranca.Enabled = False
    End With
    
Case "Desabastecimento CCR 04/2015"
    Select Case strPeticao
    Case "Contesta��o"
        Set form = New frmContDesabastecimento
        
    Case "RI/Apela��o"
        Set form = New frmRIDesabastecimento
        form.cmbTestemunhas.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Testemunhas").Address
        
    Case "Contrarraz�es de RI/Apela��o"
        GoTo N�oFaz
        'Set form = New frmCRRIDesabGenerico
        'form.cmbTestemunhas.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Testemunhas").Address
        
    End Select
    
    With form
        .frmDesabastecimento.Caption = "Desabastecimento - CCR - Abril/2015"
        .chbPrescricao.Caption = "Ajuizamento posterior a 02/04/2018"
        .chbSemFatura.Caption = "Autor n�o juntou conta do m�s em quest�o (normalmente, Maio/2015)"
        .cmbAlteracaoConsumo.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Altera��oConsumoCCR").Address
        .cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Devolu��oemdobro").Address
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
    Case "Contesta��o"
        Set form = New frmContDesabastecimento
        
    Case "RI/Apela��o"
        Set form = New frmRIDesabastecimento
        form.cmbTestemunhas.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Testemunhas").Address
        
    Case "Contrarraz�es de RI/Apela��o"
        GoTo N�oFaz
        'Set form = New frmCRRIDesabGenerico
        'form.cmbTestemunhas.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Testemunhas").Address
        
    End Select
    
    With form
        .frmDesabastecimento.Caption = "Desabastecimento - Uruguai, Mares - Set/2016"
        .chbPrescricao.Caption = "Ajuizamento posterior a 14/09/2019"
        .chbSemFatura.Caption = "Autor n�o juntou conta do m�s em quest�o (normalmente, Out e Nov/2016)"
        .cmbAlteracaoConsumo.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Altera��oConsumoCCR").Address
        .cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Devolu��oemdobro").Address
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
    Case "Contesta��o"
        Set form = New frmContDesabastecimento
        
    Case "RI/Apela��o"
        Set form = New frmRIDesabastecimento
        form.cmbTestemunhas.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Testemunhas").Address
        
    Case "Contrarraz�es de RI/Apela��o"
        GoTo N�oFaz
        'Set form = New frmCRRIDesabGenerico
        'form.cmbTestemunhas.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Testemunhas").Address
        
    End Select
    
    With form
        .frmDesabastecimento.Caption = "Desabastecimento - Liberdade, IAPI, Pero Vaz, Curuzu, Santa M�nica - Out/2017"
        .chbPrescricao.Caption = "Ajuizamento posterior a 02/11/2020"
        .chbSemFatura.Caption = "Autor n�o juntou conta do m�s em quest�o (normalmente, Dezembro/2017)"
        .cmbAlteracaoConsumo.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Altera��oConsumoCCR").Address
        .cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Devolu��oemdobro").Address
        .cmbCorresponsavel.Text = ""
        .cmbCorresponsavel.Enabled = False
        .txtDataInicio.Text = "30/10/2017"
        .txtDuracao.Text = "quinze dias"
        .chbDanMorCorte.Value = False
        .chbDanMorCorte.Enabled = False
        .chbDanMorMeraCobranca.Value = False
        .chbDanMorMeraCobranca.Enabled = False
    End With
    
    
Case "Desabastecimento Apag�o Xingu 03/2018"
    Select Case strPeticao
    Case "Contesta��o"
        Set form = New frmContDesabastecimento
        
    Case "RI/Apela��o"
        Set form = New frmRIDesabastecimento
        form.cmbTestemunhas.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Testemunhas").Address
        
    Case "Contrarraz�es de RI/Apela��o"
        GoTo N�oFaz
        'Set form = New frmCRRIDesabGenerico
        'form.cmbTestemunhas.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Testemunhas").Address
        
    End Select
    
    With form
        .frmDesabastecimento.Caption = "Desabastecimento - Apag�o Xingu 03/2018"
        .chbPrescricao.Caption = "Ajuizamento posterior a 21/03/2021"
        .chbSemFatura.Caption = "Autor n�o juntou conta do m�s em quest�o (normalmente, Abr ou Mai/2018)"
        .cmbAlteracaoConsumo.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Altera��oConsumoCCR").Address
        .cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Devolu��oemdobro").Address
        .cmbCorresponsavel.Text = "Operador Nacional do Sistema El�trico - ONS"
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
    Case "Contesta��o"
        Set form = New frmContFixoEsgoto
        
    Case "RI/Apela��o"
        GoTo N�oFaz
        
    Case "Contrarraz�es de RI/Apela��o"
        GoTo N�oFaz
        
    End Select
    
    With form
        .cmbConfessaPoco.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("PosturaAutorPo�o").Address
        .cmbVolumeParadigma.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("VolumeParadigmaEsgoto").Address
        .cmbPorcentEsgoto.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("TarifaEsgotoPercent").Address
        .cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Devolu��oemdobro").Address
        .cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Devolu��oemdobro").Address
    End With
    
Case "Outros v�cios"
    Select Case strPeticao
    Case "Contesta��o"
        Set form = New frmContVicio
        
    Case "RI/Apela��o"
        GoTo N�oFaz
        
    Case "Contrarraz�es de RI/Apela��o"
        GoTo N�oFaz
        
    End Select
    
    form.cmbPagamento.RowSource = "'[" & ThisWorkbook.Name & "]cfConfigura��es'!" & ThisWorkbook.Sheets("cfConfigura��es").Range("Devolu��oemdobro").Address
    
Case Else
N�oFaz:
    MsgBox "Sinto muito, " & DeterminarTratamento & "! Eu ainda n�o sei fazer " & strPeticao & " de " & strCausaPedir & ". Aguarde uma nova vers�o -- eu juro que vou tentar aprender!", vbInformation + vbOKOnly, "S�sifo em treinamento"
    Set ConfigurarFormulario = Nothing
    Exit Function
    
End Select

Set ConfigurarFormulario = form

End Function

Function ObterVariaveis(strVariavel As String, strCausaPedir As String, form As Variant, Optional varCont As Variant)
''
'' Retorna o valor de uma vari�vel, conforme a strCausaPedir.
''
Dim X As Variant
Dim btCont As Byte
Dim strCont As String
Dim fatImpugnada As Fatura, fatPretendida As Fatura
        

' Vari�veis de mais de uma causa de pedir
Select Case strVariavel
    Case "data-audi�ncia-concilia��o"
        X = Trim(Application.InputBox(DeterminarTratamento & ", informe a data da audi�ncia de concilia��o", "S�sifo - Dados de tempestividade", varCont, Type:=2))
    Case "termo-decad�ncia"
        X = CDate(InputBox(DeterminarTratamento & ", qual foi a data da reclama��o administrativa ou ajuizamento da a��o?", "Informa��es para decad�ncia", Format(Date - 30, "dd/mm/yyyy"))) - 30
        GoTo AtribuiVariavel
    Case "pediu-devolu��o"
        X = IIf(form.chbDevolDobro.Value = True, ", com devolu��o em dobro dos valores pagos de forma alegadamente indevida no per�odo impugnado", "")
        GoTo AtribuiVariavel
    Case "pediu-danos-materiais"
        X = IIf(form.chbDanMat.Value = True, " e indeniza��o por danos materiais", "")
        GoTo AtribuiVariavel
    Case "pediu-danos-morais"
        X = IIf(form.chbDanoMoral.Value = True, ", al�m de indeniza��o por danos morais", "")
        GoTo AtribuiVariavel
    Case "comarca-competente"
        X = Trim(form.txtComarcaCompetente.Value)
        GoTo AtribuiVariavel
    Case "m�s-inicial"
        X = InputBox(DeterminarTratamento & ", informe o m�s de refer�ncia da fatura inicial do per�odo impugnado pela parte Adversa", "Informa��es sobre pedido", "abril/2019")
        GoTo AtribuiVariavel
    Case "resumo-da-m�-f�"
        X = InputBox(DeterminarTratamento & ", fa�a um resumo da conduta da parte Adversa que constitui m�-f�, completando a seguinte frase:" & vbCrLf & "A parte Autora ...", "Informa��es sobre m�-f�", "sonegou a informa��o de que esta empresa verificara o vazamento e a cientificara previamente")
        GoTo AtribuiVariavel
End Select

Select Case strCausaPedir
Case "Revis�o de consumo elevado"
    ' Vari�veis de consumo elevado
    Select Case strVariavel
    Case "per�odo"
        X = InputBox("Informe o per�odo impugnado:", "Informe o per�odo", "de Setembro a Novembro/2019")
    Case "m�s-in�cio-medi��es"
        X = InputBox("Informe o m�s de refer�ncia da fatura em que se iniciaram as medi��es de consumo por hidr�metro:", "Informa��es sobre consumo", "08/2019")
    Case "m�dia-real"
        X = Trim(Application.InputBox(DeterminarTratamento & ", informe a m�dia de consumo real dos 12 meses anteriores ao per�odo impugnado, em m3:", "Informe a m�dia real", Type:=2))
    Case "m�dia-afirmada"
        X = InputBox(DeterminarTratamento & ", informe o valor que a parte alega ser sua m�dia de consumo, em m3:", "Informe a m�dia alegada", "10")
    Case "tempo-sem-medi��o"
        X = InputBox(DeterminarTratamento & ", informe o tempo que o im�vel passou sem medi��es reais de consumo antes de ser instalado hidr�metro, completando a frase abaixo:" & vbCrLf & """H� ..., a parte Demandante n�o tem consumo mensal no valor que afirma ser sua m�dia.""", "Informa��es sobre consumo", "anos")
    Case "consumo-impugnado"
        X = Trim(Application.InputBox(DeterminarTratamento & ", informe o valor do(s) consumo(s) impugnado(s), em m3:", "Informe o consumo impugnado", "13,  15", Type:=2))
    Case "consumo-fict�cio-utilizado"
        X = Trim(Application.InputBox(DeterminarTratamento & ", informe o valor do(s) consumo(s) fict�cio(s) utilizados antes da instala��o do hidr�metro, em m�:", "Informa��es sobre consumo", "06", Type:=2))
    Case "n�mero-hidr�metro"
        X = Trim(Application.InputBox(DeterminarTratamento & ", informe o n�mero do hidr�metro", "Informe o n�mero do hidr�metro", Type:=2))
    Case "substitui��o-hidr�metro"
        X = Trim(Application.InputBox(DeterminarTratamento & ", informe a data em que o hidr�metro foi instalado ou substitu�do", "Informe data da substitui��o", Type:=2))
    Case "n�meros-processos"
        X = InputBox(DeterminarTratamento & ", informe os n�meros dos processos anteriores que resultaram na m�dia viciada", "Informe n�meros de processos anteriores", "XX, YY e ZZ")
    Case "tempo-medi��es-fict�cias"
        X = InputBox(DeterminarTratamento & ", informe por quanto tempo, aproximadamente, o Autor pagou por m�nimo ou m�dia de consumo", "Informe tempo pagando m�dia", "cerca de um ano")
    Case "m�dias-estabelecidas-judicialmente"
        X = InputBox(DeterminarTratamento & ", informe o valor das m�dias estabelecidas judicialmente nos processos anteriores", "Informe valor das m�dias judiciais", "10 ou 13")
    Case "medi��o-maior"
        X = InputBox(DeterminarTratamento & ", informe o valor da maior medi��o nos �ltimos tempos", "Informe valor mais alto")
    Case "quantidade-habitantes"
        X = InputBox(DeterminarTratamento & ", informe a quantidade de habitantes do im�vel, segundo informa��es dos autos", "Quantidade de habitantes", "03")
        btCont = CByte(X)
    Case "proje��o-m�dia"
        X = btCont * 5.4
    Case "exemplos-consumo"
        X = Trim(Application.InputBox(DeterminarTratamento & ", qual � a faixa de consumo do per�odo impugnado, para compara��o com a quantidade de habitantes?", "Informe a faixa de consumo impugnada em m3", "10 a 13", Type:=2))
    Case "devolu��o-em-dobro"
        X = IIf(form.chbDevolDobro.Value = True, " Pede a devolu��o em dobro dos valores pagos em excesso.", "")
    Case "dano-moral"
        X = IIf(form.chbDanoMoral.Value = True, " Requer, ainda, indeniza��o por danos morais.", "")
    End Select

Case "Corte no fornecimento"
    ' Vari�veis de corte
    Select Case strVariavel
    Case "data-corte"
        X = InputBox(DeterminarTratamento & ", informe a data do corte:", "Informa��es da s�ntese dos fatos", Format(Date - 60, "dd/mm/yyyy"))
    Case "conta-raz�o-corte"
        X = InputBox(DeterminarTratamento & ", informe o per�odo que motivou o corte:", "Informa��es da s�ntese dos fatos", "de Setembro e Outubro/2017")
    Case "data-pagamento-alegado"
        X = InputBox(DeterminarTratamento & ", informe a data em que o Autor alega haver pago:", "Informa��es da s�ntese dos fatos", Format(Date - 15, "dd/mm/yyyy"))
    Case "vencimentos-d�bitos"
        X = InputBox(DeterminarTratamento & ", digite as datas de vencimentos das faturas que estavam em aberto", "Informa��es sobre exist�ncia de d�bito", "11/05/2017, 11/06/2017")
    Case "maior-per�odo-atraso"
        X = InputBox(DeterminarTratamento & ", qual foi o maior per�odo que a parte Autora atrasou os pagamentos na �poca dos fatos?", "Informa��es sobre exist�ncia de d�bito", "quase tr�s meses")
    Case "quantas-equipes-corte"
        X = InputBox(DeterminarTratamento & ", quantas SSs de corte foram iniciadas no hist�rico da parte Autora?", "Informa��es sobre inadimpl�ncia contumaz", "cinco")
    Case "atraso-pagamento-v�spera"
        X = InputBox(DeterminarTratamento & ", o pagamento feito na v�spera do corte ocorreu com quanto tempo de atraso?", "Informa��es sobre pagamento na v�spera", "um m�s e meio")
    Case "dura��o-do-corte"
        X = InputBox(DeterminarTratamento & ", qual foi a dura��o do corte?", "Informa��es sobre o corte de curta dura��o", "algumas poucas horas")
    End Select

Case "Negativa��o no SPC"
    ' Vari�veis de Negativa��o
    Select Case strVariavel
    Case "matr�cula"
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
    'Case "data-negativa��o"
    '    X = Trim(Application.InputBox(DeterminarTratamento & ", qual � a data da negativa��o realizada pela Embasa?", "Informe data da negativa��o", Type:=2))
    Case "m�s-final-do-uso-regular"
        X = Trim(form.txtMesFinalUsoRegular.Text)
    Case "empresas-negativadoras-anteriores"
        X = Trim(InputBox(DeterminarTratamento & ", quais foram as empresas que negativaram a parte Autora ANTES da negativa��o realizada pela Embasa?", "Informe as empresas negativadoras anteriores"))
    Case "data-parcelamento"
        X = CDate(Trim(Application.InputBox(DeterminarTratamento & ", qual � a data da realiza��o do parcelamento?", "Informe data do parcelamento", Type:=2)))
    Case "per�odo-parcelado"
        X = Trim(InputBox(DeterminarTratamento & ", quais foram os meses das contas parceladas?", "Informe o per�odo abrangido pelo parcelamento", "01/2016 at� 06/2017"))
    Case "n�meros-processos-negativa��o"
        X = Trim(InputBox(DeterminarTratamento & ", quais foram os n�meros dos outros processos sobre as outras linhas da mesma negativa��o?", "Informe os outros processos"))
    Case "endere�o-matr�cula"
        X = Trim(InputBox(DeterminarTratamento & ", qual � o endere�o da liga��o registrado no SCI?", "Informe o endere�o da matr�cula"))
    End Select

Case "Realizar liga��o de �gua", "Desmembramento de liga��es"
    ' Vari�veis de Realizar liga��o
    Select Case strVariavel
    Case "data-solicita��o"
        X = CDate(Trim(Application.InputBox(DeterminarTratamento & ", informe a data da solicita��o de liga��o:", "Informe data da solicita��o", Type:=2)))
    Case "n�mero-SS"
        X = Trim(Application.InputBox(DeterminarTratamento & ", informe o n�mero da Solicita��o de Servi�o de liga��o:", "Informe SS", Type:=2))
    Case "data-liga��o"
        X = CDate(Trim(Application.InputBox(DeterminarTratamento & ", informe a data da liga��o:", "Informe data da liga��o", Type:=2)))
    Case "exigencia-alegada"
        If form.chbSemReservatorioBomba.Value = True Then
            X = " (instala��o de reservat�rio ao n�vel do solo com bomba)"
        ElseIf form.chbSepararInstalacoesInternas.Value = True Then
            X = " (separa��o das instala��es hidr�ulicas)"
        ElseIf form.chbAltitudeInsuficiente.Value = True Then
            X = " (dist�ncia superior � regulamentar da extremidade da rede e altitude demasiada do seu terreno)"
        End If
    Case "alega-compara��o-com-vizinhos"
        X = IIf(form.chbComparacaoVizinhos.Value = True, " Afirma equivocadamente que tal tratamento seria uma arbitrariedade contra sua pessoa, e que n�o foi exigido de nenhum de seus vizinhos.", "")
    Case "defesa-compara��o-com-vizinhos"
        X = IIf(form.chbComparacaoVizinhos.Value = True, "3) N�o houve tratamento diferenciado entre Demandante e vizinhos. Se houve alguma exce��o em im�vel vizinho - no que n�o acreditamos, e n�o h� provas nos autos -, certamente o foi por ordem judicial ou anterior � vig�ncia do regulamento.", "")
    Case "defesa-compara��o-com-vizinhos2"
        X = IIf(form.chbComparacaoVizinhos.Value = True, " Ademais, eventual erro ocorrido com vizinho n�o gera para a parte Autora direito subjetivo a um erro similar! Esse pensamento � absurdo: um erro � um erro; ele deve ser corrigido, e n�o gera direito a novos erros!", "")
    Case "pedidos-compara��o-com-vizinhos"
        X = IIf(form.chbComparacaoVizinhos.Value = True, "; bem como por n�o ter havido discrimina��o para com vizinhos - e mesmo que, por lapso, um vizinho tenha conseguido liga��o de �gua infringindo as normas t�cnicas, um erro anterior deve ser consertado, e n�o gera direito a novos erros", "")
    End Select

Case "Cobran�a de esgoto em im�vel n�o ligado � rede"
    ' Vari�veis de Esgoto sem liga��o � rede
    Select Case strVariavel
    Case "motivo-dano-moral"
        X = "cobran�a de tarifa de esgoto em per�odo no qual alega n�o ter utilizado o servi�o"
    Case "di�metro-rede"
        X = InputBox(DeterminarTratamento & ", informe o di�metro da rede que serve o im�vel da parte Autora", "Informa��es sobre cobran�a de esgoto sem liga��o � rede", "150 mm")
    Case "data-implanta��o-esgoto"
        X = Trim(form.txtDataImplantacao.Text)
    Case "m�s-implanta��o-esgoto"
        X = Trim(Format(form.txtDataImplantacao, "mmmm/yyyy"))
    Case "quantidade-liga��es-no-m�s"
        X = InputBox(DeterminarTratamento & ", informe quantas liga��es foram realizadas pela unidade no mesmo m�s que a do im�vel do Autor", "Informa��es sobre cobran�a de esgoto sem liga��o � rede", "30")
    Case "escrit�rio-local"
        X = InputBox(DeterminarTratamento & ", informe qual o escrit�rio local do im�vel da parte Autora, conforme agrupamento no gr�fico de liga��es de esgoto", "Informa��es sobre cobran�a de esgoto sem liga��o � rede", "Escrit�rio de Servi�os de Itapu�")
    End Select

Case "Cobran�a de esgoto com �gua cortada"
    ' Vari�veis de Esgoto sem liga��o � rede
    Select Case strVariavel
    Case "�gua-cortada"
        X = IIf(form.cmbAtitudeAutor.Value = "�gua estava cortada, n�o pode ser cobrado esgoto", "seu im�vel est� com o abastecimento de �gua cortado, mas recebe cobran�as do servi�o de coleta de esgoto", "")
    Case "im�vel-desabitado"
        X = IIf(form.cmbAtitudeAutor.Value = "Im�vel estava desabitado no per�odo impugnado", "seu im�vel est� desabitado, sem uso do servi�o de abastecimento, mas recebe cobran�as do servi�o de coleta de esgoto", "")
    Case "solicitou-suspens�o"
        X = IIf(form.cmbAtitudeAutor.Value = "Afirma apenas que solicitou cancelamento", "solicitou suspens�o do abastecimento de �gua de seu im�vel, mas recebe cobran�as do servi�o de coleta de esgoto", "")
    Case "motivo-dano-moral"
        X = "cobran�a de tarifa de esgoto em per�odo no qual n�o houve abastecimento de �gua"
    Case "pede-devolu��o-em-dobro"
        X = IIf(form.chbDevolDobro.Value = True, ", a devolu��o em dobro dos valores pagos", "")
    End Select

Case "Classifica��o tarifa ou qtd. de economias"
    ' Vari�veis de Classifica��o tarif�ria
    Select Case strVariavel
    Case "motivo-dano-moral"
        X = "cobran�a de tarifa em desconformidade com a classifica��o do im�vel"
    Case "m�s-final"
        X = Trim(InputBox(DeterminarTratamento & ", informe o m�s de refer�ncia da fatura final do per�odo pleiteado pela parte Adversa", "Informa��es sobre pedido", Trim(form.txtRefAlteracao2.Text)))
    Case "classifica��o-original"
        X = Trim(form.txtClassifOriginal.Text)
    Case "classifica��o-original-sem-ponto"
        X = CodClassificacaoFatura(Trim(form.txtClassifOriginal.Text))
    Case "categoria-original"
        X = CategoriaExtenso(Trim(form.txtClassifOriginal.Text))
    Case "qtd-economias-original"
        X = EconomiasExtenso(Trim(form.txtClassifOriginal.Text))
    Case "data-segunda-classifica��o"
        X = Trim(form.txtDataAlteracao1.Text)
    Case "classifica��o-segunda"
        X = Trim(form.txtClassifAlteracao1.Text)
    Case "classifica��o-segunda-sem-ponto"
        X = CodClassificacaoFatura(Trim(form.txtClassifAlteracao1.Text))
    Case "categoria-segunda"
        X = CategoriaExtenso(Trim(form.txtClassifAlteracao1.Text))
    Case "qtd-economias-segunda"
        X = EconomiasExtenso(Trim(form.txtClassifAlteracao1.Text))
    Case "refer�ncia-segunda-classifica��o"
        X = Trim(form.txtRefAlteracao1.Text)
    Case "motivo-terceira-classifica��o"
        X = InputBox(DeterminarTratamento & ", informe o motivo da segunda reclassifica��o, completando a frase abaixo" & vbCrLf & vbCrLf & "fiscaliza��o na qual se percebeu que...", "Informa��es sobre classifica��o tarif�ria", "o im�vel efetivamente tinha composi��o diferente")
    Case "data-terceira-classifica��o"
        X = Trim(form.txtDataAlteracao2.Text)
    Case "classifica��o-terceira"
        X = Trim(form.txtClassifAlteracao2.Text)
    Case "classifica��o-terceira-sem-ponto"
        X = CodClassificacaoFatura(Trim(form.txtClassifAlteracao2.Text))
    Case "categoria-terceira"
        X = CategoriaExtenso(Trim(form.txtClassifAlteracao2.Text))
    Case "qtd-economias-terceira"
        X = EconomiasExtenso(Trim(form.txtClassifAlteracao2.Text))
    Case "refer�ncia-terceira-classifica��o"
        X = Trim(form.txtRefAlteracao2.Text)
    Case "composi��o-alegada-embasa-extenso"
        X = InputBox(DeterminarTratamento & ", informe a classifica��o tarif�ria defendida pela Embasa, por extenso -- lembrando que o Manual Comercial da Embasa permite que o consumidor opte por ser faturado como uma �nica economia", "Informa��es sobre classifica��o pretendida", ClassificacaoExtenso(Trim(form.txtClassifAlteracao1.Text)))
    Case "classifica��o-pretendida-extenso"
        X = InputBox(DeterminarTratamento & ", informe a classifica��o tarif�ria pretendida pela parte Autora, por extenso", "Informa��es sobre classifica��o pretendida", ClassificacaoExtenso(Trim(form.txtExPretCat.Text) & "." & Trim(form.txtExPretEconomias.Text)))
    Case "m�s-inicial-contraposto"
        X = InputBox(DeterminarTratamento & ", informe o m�s de refer�ncia da fatura inicial do per�odo calculado a cobrar da parte Adversa", "Informa��es sobre reconven��o ou pedido contraposto", Trim(form.txtRefAlteracao1.Text))
    Case "m�s-final-contraposto"
        X = InputBox(DeterminarTratamento & ", informe o m�s de refer�ncia da fatura final do per�odo calculado a cobrar da parte Adversa", "Informa��es sobre reconven��o ou pedido contraposto", Trim(form.txtRefAlteracao2.Text))
    Case "valor-diferen�a"
        X = InputBox(DeterminarTratamento & ", informe o valor da diferen�a total a ser demandado da parte Autora", "Informa��es sobre reconven��o ou pedido contraposto")
        X = Format(X, "#,##0.00")
    Case "inexist�ncia-de-solicita��o"
        X = IIf(form.chbRetroativoSemSolicitacao.Value = True, ", e nunca houve solicita��o do consumidor para a mudan�a", "")
    Case "pede-reclassifica��o-retroativa"
        X = IIf(form.chbRetroativoSemSolicitacao.Value = True, "especialmente o de reclassifica��o retroativa, ", "")
    End Select
    
    If form.chbApresentarCalcExemplo.Value = True Then
        ''Configura fatura e faz os c�lculos
        Set fatImpugnada = New Fatura
        Set fatPretendida = New Fatura
        
        If form.chbTemDoisTipos.Value = False Then '1 categoria s�
            fatImpugnada.BuscarEstruturaTarifaria form.txtMesRefExemplo.Text, form.txtConsExemplo.Text, form.cmbPorcentEsgoto.Text, form.txtExFaturado1Cat.Text, form.txtExFaturado1Economias.Text
        Else '2 categorias
            fatImpugnada.BuscarEstruturaTarifaria form.txtMesRefExemplo.Text, form.txtConsExemplo.Text, form.cmbPorcentEsgoto.Text, form.txtExFaturado1Cat.Text, form.txtExFaturado1Economias.Text, form.txtExFaturado2Cat.Text, form.txtExFaturado2Economias.Text
        End If
        
        fatPretendida.BuscarEstruturaTarifaria form.txtMesRefExemplo.Text, form.txtConsExemplo.Text, form.cmbPorcentEsgoto.Text, form.txtExPretCat.Text, form.txtExPretEconomias.Text
        
        fatImpugnada.CalcularTotal
        fatPretendida.CalcularTotal
        
        ''Atribui as vari�veis (valores de faixas primeiro, porque s�o eliminadas mais f�cil, enquanto o Select testa uma por uma).
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
        
        ' Demais vari�veis que n�o s�o valores de faixas
        Select Case strVariavel
        Case "refer�ncia-conta-exemplo"
            X = fatImpugnada.MesReferencia
        Case "consumo-conta-exemplo"
            X = fatImpugnada.ConsumoTotal
        Case "consumo-por-economia-imp"
            X = fatImpugnada.ConsumoPorEconomia
        Case "total-�gua-imp"
            X = fatImpugnada.TotalAgua
            X = Format(X, "#,##0.00")
        Case "porcentagem-esgoto"
            X = fatImpugnada.PorcentEsgoto
        Case "total-�gua-esgoto-imp"
            X = fatImpugnada.TotalAguaEsgoto
            X = Format(X, "#,##0.00")
        Case "qt-economias-imp"
            X = fatImpugnada.QtdTotalEconomias
        Case "total-esgoto-imp"
            X = fatImpugnada.TotalEsgoto
            X = Format(X, "#,##0.00")
        Case "qt-economias-imp1"
            X = fatImpugnada.Categorias(1).QtdEconomias
        Case "classifica��o-imp1"
            X = CategoriaExtenso(fatImpugnada.Categorias(1).Categoria)
        Case "�gua-por-econ-imp1"
            X = fatImpugnada.Categorias(1).AguaPorEconomia
            X = Format(X, "#,##0.00")
        Case "subt-�gua-imp1"
            X = fatImpugnada.Categorias(1).SubtotalAgua
            X = Format(X, "#,##0.00")
        Case "qt-economias-imp2"
            X = fatImpugnada.Categorias(2).QtdEconomias
        Case "classifica��o-imp2"
            X = CategoriaExtenso(fatImpugnada.Categorias(2).Categoria)
        Case "�gua-por-econ-imp2"
            X = fatImpugnada.Categorias(2).AguaPorEconomia
            X = Format(X, "#,##0.00")
        Case "subt-�gua-imp2"
            X = fatImpugnada.Categorias(2).SubtotalAgua
            X = Format(X, "#,##0.00")
        Case "classifica��o-pretendida"
            X = ClassificacaoExtenso(form.txtExPretCat.Text & "." & form.txtExPretEconomias.Text)
        Case "consumo-por-economia-pret"
            X = fatPretendida.ConsumoPorEconomia
        Case "total-�gua-pret"
            X = fatPretendida.TotalAgua
            X = Format(X, "#,##0.00")
        Case "porcentagem-esgoto"
            X = fatPretendida.PorcentEsgoto
        Case "total-�gua-esgoto-pret"
            X = fatPretendida.TotalAguaEsgoto
            X = Format(X, "#,##0.00")
        Case "qt-economias-pret"
            X = fatPretendida.QtdTotalEconomias
        Case "total-esgoto-pret"
            X = fatPretendida.TotalEsgoto
            X = Format(X, "#,##0.00")
        Case "qt-economias-pret1"
            X = fatPretendida.Categorias(1).QtdEconomias
        Case "classifica��o-pret1"
            X = CategoriaExtenso(fatPretendida.Categorias(1).Categoria)
        Case "�gua-por-econ-pret1"
            X = fatPretendida.Categorias(1).AguaPorEconomia
            X = Format(X, "#,##0.00")
        Case "subt-�gua-pret1"
            X = fatPretendida.Categorias(1).SubtotalAgua
            X = Format(X, "#,##0.00")
        End Select
            
    End If

Case "Suspeita de by-pass"
    'Vari�veis de gato
    Select Case strVariavel
    Case "nega-autoria"
        X = IIf(form.optNegaAutoria.Value = True, "; alega que n�o foi o respons�vel pelo il�cito", "")
    Case "nega-gato"
        X = IIf(form.optNegaExistencia.Value = True, "; alega que n�o houve gato", "")
    Case "fez-reclama��es-n�o-atendidas"
        X = IIf(form.optReclamacoesNaoAtendidas.Value = True, "; alega que realizou reclama��es, as quais n�o foram atendidas", "")
    Case "pede-devolu��o-em-dobro"
        X = IIf(form.chbDevolDobro.Value = True, " e que os valores pagos sejam devolvidos em dobro", "")
    Case "data-fiscaliza��o"
        X = Trim(form.txtDataRetiradaGato.Text)
    Case "fatura-regulariza��o-consumo"
        X = Trim(form.txtMesRefRegulaConsumo.Text)
    Case "fatura-multa"
        X = Trim(form.txtMesRefMulta.Text)
    Case "total-san��es"
        X = Trim(form.txtTotalSancoes.Text)
    Case "valor-multa"
        X = IIf(Trim(form.txtValorMulta.Text) <> "", "; multa sancionat�ria pela pr�tica do il�cito, no importe de R$ " & form.txtValorMulta.Text, "")
    Case "valor-recupera��o-consumo"
        X = IIf(Trim(form.txtValorRecCons.Text) <> "", "; recupera��o de consumo pela frui��o do il�cito, calculada pela m�dia de consumo anterior ao per�odo do il�cito, no importe de R$ " & form.txtValorRecCons.Text, "")
    Case "valor-custos-reparo"
        X = IIf(Trim(form.txtValorReparo.Text) <> "", "; ressarcimento dos custos pela repara��o do il�cito, no importe de R$ " & form.txtValorReparo.Text, "")
    
    End Select

Case "D�bito de terceiro"
    'Vari�veis de gato
    Select Case strVariavel
    Case "motivo-dano-moral"
        X = "cobran�a contra si de d�bito de per�odo em que alega que o servi�o era utilizado por outro usu�rio"
    
    End Select

Case "Fixo de esgoto"
    'Vari�veis de gato
    Select Case strVariavel
    Case "motivo-dano-moral"
        X = "cobran�a de tarifa de esgoto em per�odo no qual alega n�o ter utilizado o servi�o"
    Case "confessa-hidr�metro"
        X = IIf(form.cmbConfessaPoco.Value = "Confessa que foi instalado hidr�metro no po�o", " Afirma que a Embasa instalou hidr�metro no seu po�o artesiano para medi��o do consumo deste, a fim de mensurar o valor da cobran�a de esgoto.", "")
    Case "tipo-estabelecimento"
        X = Trim(InputBox(DeterminarTratamento & ", informe o tipo de estabelecimento da parte Adversa", "Informa��es sobre pedido", "um sal�o de beleza"))
        
    End Select
    
    ''Dimensiona objetos, configura fatura e faz os c�lculos
    Set fatImpugnada = New Fatura
    
    If form.chbTemDoisTipos.Value = False Then '1 categoria s�
        fatImpugnada.BuscarEstruturaTarifaria form.txtMesRefExemplo.Text, form.txtConsExemplo.Text, form.cmbPorcentEsgoto.Text, form.txtExCat1.Text, form.txtExEconomias1.Text
    Else '2 categorias
        fatImpugnada.BuscarEstruturaTarifaria form.txtMesRefExemplo.Text, form.txtConsExemplo.Text, form.cmbPorcentEsgoto.Text, form.txtExCat1.Text, form.txtExEconomias1.Text, form.txtExCat2.Text, form.txtExEconomias2.Text
    End If
    
    fatImpugnada.CalcularTotal
    
    ''Atribui as vari�veis (valores de faixas primeiro, porque s�o eliminadas mais f�cil, enquanto o Select testa uma por uma).
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
        
    ' Demais vari�veis que n�o s�o valores de faixas
    Select Case strVariavel
    Case "refer�ncia-conta-exemplo"
        X = fatImpugnada.MesReferencia
    Case "consumo-conta-exemplo"
        X = fatImpugnada.ConsumoTotal
    Case "volume-paradigma"
        X = fatImpugnada.ConsumoTotal
    Case "consumo-por-economia"
        X = fatImpugnada.ConsumoPorEconomia
    Case "total-�gua"
        X = fatImpugnada.TotalAgua
        X = Format(X, "#,##0.00")
    Case "porcentagem-esgoto"
        X = fatImpugnada.PorcentEsgoto
    Case "total-�gua-esgoto"
        X = fatImpugnada.TotalAguaEsgoto
        X = Format(X, "#,##0.00")
    Case "qt-economias"
        X = fatImpugnada.QtdTotalEconomias
    Case "total-esgoto"
        X = fatImpugnada.TotalEsgoto
        X = Format(X, "#,##0.00")
    Case "qt-economias-1"
        X = fatImpugnada.Categorias(1).QtdEconomias
    Case "classifica��o-1"
        X = CategoriaExtenso(fatImpugnada.Categorias(1).Categoria)
    Case "�gua-por-econ-1"
        X = fatImpugnada.Categorias(1).AguaPorEconomia
        X = Format(X, "#,##0.00")
    Case "subt-�gua-1"
        X = fatImpugnada.Categorias(1).SubtotalAgua
        X = Format(X, "#,##0.00")
    Case "qt-economias-2"
        X = fatImpugnada.Categorias(2).QtdEconomias
    Case "classifica��o-2"
        X = CategoriaExtenso(fatImpugnada.Categorias(2).Categoria)
    Case "�gua-por-econ-2"
        X = fatImpugnada.Categorias(2).AguaPorEconomia
        X = Format(X, "#,##0.00")
    Case "subt-�gua-2"
        X = fatImpugnada.Categorias(2).SubtotalAgua
        X = Format(X, "#,##0.00")
    End Select
    
Case "Vaz. �gua ou extravas. esgoto com danos a patrim�nio/morais", "Obra da Embasa com danos a patrim�nio/morais", _
    "Acidente com pessoa/ve�culo em buraco", "Acidente com ve�culo (colis�o ou  atropelamento)"
    ' Vari�veis de Responsabilidade civil
    Select Case strVariavel
    Case "data-incidente"
        X = Trim(form.txtDataFato.Text)
    Case "local-incidente"
        X = InputBox(DeterminarTratamento & ", onde ocorreu o incidente narrado pelo Adverso?", "Informa��es sobre responsabilidade civil", "em Plataforma, bairro onde reside")
    Case "pediu-dano-material-RC"
        X = IIf(form.chbDanMat.Value = True, "indeniza��o por danos materiais no valor pleiteado na Inicial, al�m de ", "")
    Case "descri��o-dano-material-RC"
        X = IIf(form.chbDanMat.Value = True, " Alega ter sofrido preju�zos referentes " & InputBox(DeterminarTratamento & _
        ", descreva o dano material sofrido pela parte Autora: ""Alega ter sofrido preju�zos materiais referentes...""", _
        "Informa��es sobre responsabilidade civil", IIf(form.cmbOcorrencia.Text = "Pessoa", "a um aparelho de telefone celular quebrado", "aos reparos que se fizeram necess�rios")) & ".", "")
    Case "termo-ad-quem-prescri��o"
        X = DateAdd("yyyy", 3, CDate(form.txtDataFato))
    Case "ato-apontado-il�cito"
        If strCausaPedir = "Vaz. �gua ou extravas. esgoto com danos a patrim�nio/morais" Then
            X = InputBox(DeterminarTratamento & ", qual � a conduta que a parte Adversa aponta como il�cito?", "Informa��es sobre responsabilidade civil", "ocorreu extravasamento da rede p�blica para dentro de seu im�vel")
        
        ElseIf strCausaPedir = "Obra da Embasa com danos a patrim�nio/morais" Then
            X = InputBox(DeterminarTratamento & ", qual � a conduta que a parte Adversa aponta como il�cito?", "Informa��es sobre responsabilidade civil", "sofreu preju�zos em decorr�ncia de obra realizada pela Embasa")
        
        ElseIf strCausaPedir = "Acidente com pessoa/ve�culo em buraco" And form.cmbOcorrencia.Value = "Ve�culo" Then
            X = InputBox(DeterminarTratamento & ", qual � a conduta que a parte Adversa aponta como il�cito?", "Informa��es sobre responsabilidade civil", "um buraco realizado pela Embasa teria causado preju�zos ao seu ve�culo")
        
        ElseIf strCausaPedir = "Acidente com pessoa/ve�culo em buraco" And form.cmbOcorrencia.Value = "Pessoa" Then
            X = InputBox(DeterminarTratamento & ", qual � a conduta que a parte Adversa aponta como il�cito?", "Informa��es sobre responsabilidade civil", "sofreu preju�zos por ter ca�do em buraco realizado pela Embasa")
        
        ElseIf strCausaPedir = "Acidente com ve�culo (colis�o ou atropelamento)" And form.cmbOcorrencia.Value = "Ve�culo" Then
            X = InputBox(DeterminarTratamento & ", qual � a conduta que a parte Adversa aponta como il�cito?", "Informa��es sobre responsabilidade civil", "um ve�culo a servi�o da Embasa teria sido culpado por colis�o com o seu ve�culo")
        
        ElseIf strCausaPedir = "Acidente com ve�culo (colis�o ou atropelamento)" And form.cmbOcorrencia.Value = "Pessoa" Then
            X = InputBox(DeterminarTratamento & ", qual � a conduta que a parte Adversa aponta como il�cito?", "Informa��es sobre responsabilidade civil", "um ve�culo a servi�o da Embasa teria sido culpado por atropelamento")
        End If
        
    Case "motivo-culpa-exclusiva"
        X = InputBox(DeterminarTratamento & ", qual � o motivo da culpa exclusiva do consumidor?", "Informa��es sobre responsabilidade civil", "n�o havia ningu�m para franquear acesso ao im�vel")
    Case "descri��o-lucros-cessantes"
        X = InputBox(DeterminarTratamento & ", favor descrever qual � o motivo aceit�vel dos lucros cessantes, se houver", "Informa��es sobre responsabilidade civil", "devem restringir-se a ")
    Case "valor-lucros-cessantes"
        X = InputBox(DeterminarTratamento & ", informe o valor aceit�vel dos lucros cessantes, se houver", "Informa��es sobre responsabilidade civil", "1.000,00")
    
    End Select

Case "Desabastecimentos por per�odo e causa determinados"
    ' Vari�veis de Desabastecimento
    Select Case strVariavel
    Case "bairro"
        X = InputBox(DeterminarTratamento & ", informe o bairro do im�vel da parte Autora", "Informa��es sobre desabastecimento")
    Case "correspons�vel"
        X = form.cmbCorresponsavel.Value
    Case "data-desabastecimento"
        X = Trim(form.txtDataInicio.Text)
    Case "dura��o-alegada-desabastecimento"
        X = Trim(form.txtDuracao.Text)
    Case "m�s-refer�ncia-fatura"
        X = InputBox(DeterminarTratamento & ", informe a fatura em que est� o per�odo impugnado", "Informa��es espec�ficas de caso de Desabastecimento", "fatura de refer�ncia de Abril/2018")
    Case "prazo-final-prescri��o"
        X = Day(CDate(form.txtDataFim.Text)) & "/" & Month(CDate(form.txtDataFim.Text)) & "/" & (Year(CDate(form.txtDataFim.Text)) + 3)
    Case "lista-de-processos-matr�cula"
        X = Trim(form.txtProcessosSimilares.Text)
    Case "valor-condena��o"
        X = form.txtValorCondenacao
        X = Format(X, "#,##0.00")
    Case "pediu-dano-material"
        X = IIf(form.chbDanMat.Value = True, "materiais e ", "")
    Case "pediu-devolu��o-em-dobro"
        X = IIf(form.chbDevolDobro.Value = True, ", bem como devolu��o em dobro das faturas do per�odo", "")
    Case "exclus�o-correspons�vel"
        X = IIf(form.chbExcluiuCorresp.Value = True, " Entendeu que a " & form.cmbCorresponsavel.Value & " n�o � parte leg�tima para figurar no polo passivo da presente " & _
            "demanda.", "")
    Case "exclus�o-correspons�vel"
        If form.chbExcluiuCorresp.Visible = True Then X = IIf(form.chbExcluiuCorresp.Value = True, " Entendeu que a corr� n�o � parte leg�tima para figurar no polo passivo da presente demanda", "")
    Case "condenou-dano-material"
        X = IIf(form.chbCondenouDanosMateriais.Value = True, "indeniza��o por danos materiais, ", "")
    Case "condenou-devolu��o-em-dobro"
        X = IIf(form.chbCondenouDevDobro.Value = True, ", bem como devolu��o em dobro das faturas do per�odo", "")
    Case "bairro-n�o-atingido"
        X = IIf(form.chbBairroNaoAfetado.Value = True, ", e o bairro da parte Autora n�o foi atingido", "")
    Case "sem-altera��o-consumo"
        X = IIf(form.cmbAlteracaoConsumo.Value = "N�o houve altera��o relevante [gr�fico]" Or _
            form.cmbAlteracaoConsumo.Value = "N�o houve altera��o [print de contas juntadas pelo Autor]", _
            ", al�m de que n�o houve altera��o no consumo do im�vel", "")
    Case "motivo-incidente"
        X = InputBox(DeterminarTratamento & ", informe o motivo do desabastecimento alegado pela parte Autora, de forma resumida, para ser usado na introdu��o de alguns t�picos", "Informa��es espec�ficas de caso de Desabastecimento", "em decorr�ncia de conserto realizado na rede p�blica de abastecimento na sua regi�o")
    Case "motivo-culpa-exclusiva-terceiro"
        X = InputBox(DeterminarTratamento & ", informe o motivo da culpa exclusiva de terceiro, de forma resumida, para constar na Contesta��o", "Informa��es espec�ficas de caso de Desabastecimento", _
            "ignorou deliberadamente as m�ltiplas orienta��es recebidas desta empresa e danificou a adutora, com a inten��o ego�sta de cumprir seus prazos, " & _
            "mesmo que fosse �s custas de destruir o abastecimento de �gua da coletividade.")
    End Select

Case "Desabastecimento Apag�o Xingu 03/2018"
    ' Vari�veis de Desabastecimento
    Select Case strVariavel
    Case "bairro"
        X = InputBox(DeterminarTratamento & ", informe o bairro do im�vel da parte Autora", "Informa��es espec�ficas de caso CCR")
    Case "correspons�vel"
        X = "Operadora Nacional do Sistema El�trico - ONS, associa��o civil de CNPJ 02.831.210/0002-38"
    Case "motivo-culpa-exclusiva-terceiro"
        X = "suspendeu, por for�a maior, o fornecimento de energia el�trica necess�rio � opera��o das bombas hidr�ulicas, essenciais para a distribui��o de �gua"
    Case "data-desabastecimento"
        X = Trim(form.txtDataInicio.Text)
    Case "dura��o-alegada-desabastecimento"
        X = Trim(form.txtDuracao.Text)
    Case "m�s-refer�ncia-fatura"
        X = "faturas de refer�ncia de Abril/2018"
    Case "prazo-final-prescri��o"
        X = "21/03/2021"
    Case "lista-de-processos-matr�cula"
        X = Trim(form.txtProcessosSimilares.Text)
    Case "valor-condena��o"
        X = form.txtValorCondenacao
        X = Format(X, "#,##0.00")
    Case "pediu-dano-material"
        X = IIf(form.chbDanMat.Value = True, "materiais e ", "")
    Case "pediu-devolu��o-em-dobro"
        X = IIf(form.chbDevolDobro.Value = True, ", bem como devolu��o em dobro das faturas do per�odo", "")
    Case "exclus�o-correspons�vel"
        X = IIf(form.chbExcluiuCorresp.Value = True, " Entendeu que a " & form.cmbCorresponsavel.Value & " n�o � parte leg�tima para figurar no polo passivo da presente " & _
            "demanda.", "")
    Case "exclus�o-correspons�vel"
        If form.chbExcluiuCorresp.Visible = True Then X = IIf(form.chbExcluiuCorresp.Value = True, " Entendeu que a corr� n�o � parte leg�tima para figurar no polo passivo da presente demanda", "")
    Case "condenou-dano-material"
        X = IIf(form.chbCondenouDanosMateriais.Value = True, "indeniza��o por danos materiais, ", "")
    Case "condenou-devolu��o-em-dobro"
        X = IIf(form.chbCondenouDevDobro.Value = True, ", bem como devolu��o em dobro das faturas do per�odo", "")
    Case "bairro-n�o-atingido"
        X = IIf(form.chbBairroNaoAfetado.Value = True, ", e o bairro da parte Autora n�o foi atingido", "")
    Case "sem-altera��o-consumo"
        X = IIf(form.cmbAlteracaoConsumo.Value = "N�o houve altera��o relevante [gr�fico]" Or _
            form.cmbAlteracaoConsumo.Value = "N�o houve altera��o [print de contas juntadas pelo Autor]", _
            ", al�m de que n�o houve altera��o no consumo do im�vel", "")
    Case "motivo-incidente"
        X = ", em raz�o da falta de fornecimento de energia el�trica, decorrente de um apag�o de energia el�trica que atingiu as regi�es Norte, Nordeste e (parcialmente) Sudeste"
    
    End Select

Case "Desabastecimento Uruguai 09/2016"
    ' Vari�veis de Desabastecimento
    Select Case strVariavel
    Case "bairro"
        X = InputBox(DeterminarTratamento & ", informe o bairro do im�vel da parte Autora", "Informa��es espec�ficas de caso CCR")
    Case "data-desabastecimento"
        X = Trim(form.txtDataInicio.Text)
    Case "dura��o-alegada-desabastecimento"
        X = Trim(form.txtDuracao.Text)
    Case "m�s-refer�ncia-fatura"
        X = "faturas de refer�ncia de Outubro e Novembro/2016"
    Case "prazo-final-prescri��o"
        X = "14/09/2019"
    Case "lista-de-processos-matr�cula"
        X = Trim(form.txtProcessosSimilares.Text)
    Case "valor-condena��o"
        X = form.txtValorCondenacao
        X = Format(X, "#,##0.00")
    Case "pediu-dano-material"
        X = IIf(form.chbDanMat.Value = True, "materiais e ", "")
    Case "pediu-devolu��o-em-dobro"
        X = IIf(form.chbDevolDobro.Value = True, ", bem como devolu��o em dobro das faturas do per�odo", "")
    Case "exclus�o-correspons�vel"
        X = IIf(form.chbExcluiuCorresp.Value = True, " Entendeu que a " & form.cmbCorresponsavel.Value & " n�o � parte leg�tima para figurar no polo passivo da presente " & _
            "demanda.", "")
    Case "exclus�o-correspons�vel"
        If form.chbExcluiuCorresp.Visible = True Then X = IIf(form.chbExcluiuCorresp.Value = True, " Entendeu que a corr� n�o � parte leg�tima para figurar no polo passivo da presente demanda", "")
    Case "condenou-dano-material"
        X = IIf(form.chbCondenouDanosMateriais.Value = True, "indeniza��o por danos materiais, ", "")
    Case "condenou-devolu��o-em-dobro"
        X = IIf(form.chbCondenouDevDobro.Value = True, ", bem como devolu��o em dobro das faturas do per�odo", "")
    Case "bairro-n�o-atingido"
        X = IIf(form.chbBairroNaoAfetado.Value = True, ", e o bairro da parte Autora n�o foi atingido", "")
    Case "sem-altera��o-consumo"
        X = IIf(form.cmbAlteracaoConsumo.Value = "N�o houve altera��o relevante [gr�fico]" Or _
            form.cmbAlteracaoConsumo.Value = "N�o houve altera��o [print de contas juntadas pelo Autor]", _
            ", al�m de que n�o houve altera��o no consumo do im�vel", "")
    Case "motivo-incidente"
        X = ", em decorr�ncia de conserto realizado na rede p�blica de abastecimento na sua regi�o"
    
    End Select

Case "Desabastecimento Liberdade 10/2017"
    ' Vari�veis de Desabastecimento
    Select Case strVariavel
    Case "bairro"
        X = InputBox(DeterminarTratamento & ", informe o bairro do im�vel da parte Autora", "Informa��es espec�ficas de caso CCR")
    Case "data-desabastecimento"
        X = Trim(form.txtDataInicio.Text)
    Case "dura��o-alegada-desabastecimento"
        X = Trim(form.txtDuracao.Text)
    Case "m�s-refer�ncia-fatura"
        X = "fatura de refer�ncia de Dezembro/2017"
    Case "prazo-final-prescri��o"
        X = "02/11/2020"
    Case "lista-de-processos-matr�cula"
        X = Trim(form.txtProcessosSimilares.Text)
    Case "valor-condena��o"
        X = form.txtValorCondenacao
        X = Format(X, "#,##0.00")
    Case "pediu-dano-material"
        X = IIf(form.chbDanMat.Value = True, "materiais e ", "")
    Case "pediu-devolu��o-em-dobro"
        X = IIf(form.chbDevolDobro.Value = True, ", bem como devolu��o em dobro das faturas do per�odo", "")
    Case "exclus�o-correspons�vel"
        X = IIf(form.chbExcluiuCorresp.Value = True, " Entendeu que a " & form.cmbCorresponsavel.Value & " n�o � parte leg�tima para figurar no polo passivo da presente " & _
            "demanda.", "")
    Case "exclus�o-correspons�vel"
        If form.chbExcluiuCorresp.Visible = True Then X = IIf(form.chbExcluiuCorresp.Value = True, " Entendeu que a corr� n�o � parte leg�tima para figurar no polo passivo da presente demanda", "")
    Case "condenou-dano-material"
        X = IIf(form.chbCondenouDanosMateriais.Value = True, "indeniza��o por danos materiais, ", "")
    Case "condenou-devolu��o-em-dobro"
        X = IIf(form.chbCondenouDevDobro.Value = True, ", bem como devolu��o em dobro das faturas do per�odo", "")
    Case "bairro-n�o-atingido"
        X = IIf(form.chbBairroNaoAfetado.Value = True, ", e o bairro da parte Autora n�o foi atingido", "")
    Case "sem-altera��o-consumo"
        X = IIf(form.cmbAlteracaoConsumo.Value = "N�o houve altera��o relevante [gr�fico]" Or _
            form.cmbAlteracaoConsumo.Value = "N�o houve altera��o [print de contas juntadas pelo Autor]", _
            ", al�m de que n�o houve altera��o no consumo do im�vel", "")
    Case "motivo-incidente"
        X = ", em decorr�ncia de conserto realizado na rede p�blica de abastecimento na sua regi�o"
    
    End Select

Case "Desabastecimento CCR 04/2015"
    ' Vari�veis de Desabastecimento
    Select Case strVariavel
    Case "bairro"
        X = InputBox(DeterminarTratamento & ", informe o bairro do im�vel da parte Autora", "Informa��es espec�ficas de caso CCR")
    Case "correspons�vel"
        X = "CCR Metr�"
    Case "data-desabastecimento"
        X = Trim(form.txtDataInicio.Text)
    Case "dura��o-alegada-desabastecimento"
        X = Trim(form.txtDuracao.Text)
    Case "m�s-refer�ncia-fatura"
        X = "fatura de refer�ncia de Maio/2015"
    Case "prazo-final-prescri��o"
        X = "08/04/2018"
    Case "lista-de-processos-matr�cula"
        X = Trim(form.txtProcessosSimilares.Text)
    Case "valor-condena��o"
        X = form.txtValorCondenacao
        X = Format(X, "#,##0.00")
    Case "pediu-dano-material"
        X = IIf(form.chbDanMat.Value = True, "materiais e ", "")
    Case "pediu-devolu��o-em-dobro"
        X = IIf(form.chbDevolDobro.Value = True, ", bem como devolu��o em dobro das faturas do per�odo", "")
    Case "exclus�o-correspons�vel"
        X = IIf(form.chbExcluiuCorresp.Value = True, " Entendeu que a corr� n�o � parte leg�tima para figurar no polo passivo da presente demanda", "")
    Case "condenou-dano-material"
        X = IIf(form.chbCondenouDanosMateriais.Value = True, "indeniza��o por danos materiais, ", "")
    Case "condenou-devolu��o-em-dobro"
        X = IIf(form.chbCondenouDevDobro.Value = True, ", bem como devolu��o em dobro das faturas do per�odo", "")
    Case "bairro-n�o-atingido"
        X = IIf(form.chbBairroNaoAfetado.Value = True, ", e o bairro da parte Autora n�o foi atingido", "")
    Case "sem-altera��o-consumo"
        X = IIf(form.cmbAlteracaoConsumo.Value = "N�o houve altera��o relevante [gr�fico]" Or _
            form.cmbAlteracaoConsumo.Value = "N�o houve altera��o [print de contas juntadas pelo Autor]", _
            ", al�m de que n�o houve altera��o no consumo do im�vel", "")
    Case "motivo-incidente"
        X = ", em decorr�ncia do not�rio rompimento de uma adutora, causado pela Companhia do Metr� da Bahia em 01/04/2015"
    Case "motivo-culpa-exclusiva-terceiro"
        X = "pois a CCR ignorou as orienta��es da Embasa e destruiu a adutora que abastece parte relevante da cidade com dolo eventual, pelo motivo esp�rio de n�o pagar multa administrativa pelo atraso na obra"
    
    End Select

Case "Outros v�cios"
    ' Vari�veis de outros v�cios
    Select Case strVariavel
    Case "sanado-administrativamente"
        X = IIf(form.chbVicioConsertado.Value = True, ", o qual foi sanado administrativamente", "")
    Case "m�s-final"
        X = InputBox(DeterminarTratamento & ", informe o m�s de refer�ncia da fatura final do per�odo impugnado pela parte Adversa", "Informa��es sobre pedido", "agosto/2019")
    Case "consumo-afirmado-autor"
        X = Trim(InputBox(DeterminarTratamento & ", informe o consumo alegado pelo Autor, completando a frase abaixo:" & vbCrLf & """Alega que seu consumo efetivo � ... m3""", "Informa��es sobre pedido", "06"))
    Case "motivo-desligamento"
        strCont = Trim(InputBox(DeterminarTratamento & ", informe o motivo do pedido de suspens�o do abastecimento, completando a frase abaixo:" & vbCrLf & """(...) requereu o desligamento do abastecimento de �gua, ...""", "Informa��es sobre pedido", "haja vista haver deixado de residir no im�vel"))
        X = IIf(strCont <> "", ", " & strCont, "")
    Case "motivo-inexist�ncia-v�cio"
        strCont = Trim(InputBox(DeterminarTratamento & ", informe o motivo do pedido de inexist�ncia de v�cio, completando a frase abaixo:" & vbCrLf & """(...) n�o houve nem defeito nem v�cio do servi�o. ...""", "Informa��es sobre pedido", "As cobran�as realizadas pela Embasa foram feitas nos estritos limites dos valores contratados"))
        X = IIf(strCont <> "", ". " & strCont, "")
    Case "pretens�o-autoral"
        X = Trim(InputBox(DeterminarTratamento & ", informe a pretens�o da parte Autora, completando a frase abaixo:" & vbCrLf & """(...) Caso a parte Autora, por quest�es particulares, pretenda ...""", "Informa��es sobre pedido", "suspender as cobran�as de tarifa de esgoto"))
    Case "requisitos-pretens�o"
        strCont = Trim(InputBox(DeterminarTratamento & ", informe os requisitos da pretens�o da parte Autora, completando a frase abaixo:" & vbCrLf & """(...) deveria requerer � Embasa e demonstrar o cumprimento dos requisitos (...)""", "Informa��es sobre pedido", "desabita��o do im�vel"))
        X = IIf(strCont <> "", " (" & strCont & ")", "")
    Case "data-requerimento-corte"
        X = Trim(InputBox(DeterminarTratamento & ", informe a data em que a parte Autora requereu o corte", "Informa��es sobre pedido"))
    
    End Select

End Select

AtribuiVariavel:

ObterVariaveis = X

End Function

Function DescobrirPlanilhaDeEstrutura(strPeticao As String, strOrgao As String) As Excel.Worksheet
''
'' Descobre qual planilha tem a lista correta de t�picos, conforme o tipo de peti��o e o �rg�o.
''
    Dim plan As Excel.Worksheet
    
    Select Case strPeticao
    Case "Contesta��o"
        Select Case strOrgao
        Case "JEC"
            Set plan = ThisWorkbook.Sheets("cfContesta��esJEC")
        Case "VC"
            Set plan = ThisWorkbook.Sheets("cfContesta��esVC")
        Case "Procon"
            Set plan = ThisWorkbook.Sheets("cfContesta��esProcon")
        End Select
        
    Case "RI/Apela��o"
        Select Case strOrgao
        Case "JEC"
            Set plan = ThisWorkbook.Sheets("cfRIsJEC")
        Case "VC"
            Set plan = ThisWorkbook.Sheets("cfApela��esVC")
        'Case "Procon"
        '    Set Plan = ThisWorkbook.Sheets("cfRIsProcon")
        End Select
                
    Case "Contrarraz�es de RI/Apela��o"
        Select Case strOrgao
        Case "JEC"
            Set plan = ThisWorkbook.Sheets("cfCRRIsJEC")
        Case "VC"
            Set plan = ThisWorkbook.Sheets("cfCRApela��esVC")
        'Case "Procon"
        '    Set Plan = ThisWorkbook.Sheets("cfContesta��esProcon")
        End Select
        
    End Select
    
    Set DescobrirPlanilhaDeEstrutura = plan

End Function


Function ColherTopicosGeral(strCausaPedir As String, strCausaPedirParadigma As String, strTipoPeticao As String, form As Variant, ByRef bolGrafico As Boolean, ByRef btTabela As Byte) As String
''
'' Chama as fun��es de colher os t�picos conforme a causa de pedir. Tamb�m altera a vari�vel "
''
    
    Dim strTopicos As String
    
    Select Case strCausaPedirParadigma
    Case "Revis�o de consumo elevado"
        Select Case strTipoPeticao
        Case "Contesta��o"
            strTopicos = ColherTopicosContRevisaoConsumo(form)
            'Verifica se � um dos t�picos que cont�m gr�ficos
            If InStr(1, strTopicos, ",,50,,") <> 0 Or InStr(1, strTopicos, ",,55,,") <> 0 Then bolGrafico = True
            
            'Verifica se � um dos t�picos que cont�m tabelas de m�dia, anotando quantas tabelas s�o.
            btTabela = 0
            If InStr(1, strTopicos, ",,95,,") <> 0 Then btTabela = btTabela + 1
            If InStr(1, strTopicos, ",,115,,") <> 0 Then btTabela = btTabela + 1
            If InStr(1, strTopicos, ",,120,,") <> 0 Then btTabela = btTabela + 1
        
        'Case "RI/Apela��o"
        '    strTopicos = ColherTopicosRICCR(form)
        '    If InStr(1, strTopicos, ",,40,,") <> 0 Or InStr(1, strTopicos, ",,42,,") <> 0 Then bolGrafico = True 'Verifica se � um dos t�picos que cont�m gr�ficos
            
        'Case "Contrarraz�es de RI/Apela��o"
        '    strTopicos = ColherTopicosCRRICCR(form)
        '    If InStr(1, strTopicos, ",,40,,") <> 0 Or InStr(1, strTopicos, ",,42,,") <> 0 Then bolGrafico = True 'Verifica se � um dos t�picos que cont�m gr�ficos
            
        End Select
        
    Case "Corte no fornecimento"
        Select Case strTipoPeticao
        Case "Contesta��o"
            strTopicos = ColherTopicosContCorte(form)
            'Verifica se � um dos t�picos que cont�m gr�ficos
            If InStr(1, strTopicos, ",,50,,") <> 0 Or InStr(1, strTopicos, ",,55,,") <> 0 Then bolGrafico = True
                        
        Case "RI/Apela��o"
            strTopicos = ColherTopicosRICorte(form)
            
        Case "Contrarraz�es de RI/Apela��o"
            'strTopicos = ColherTopicosCRRIDesabGenerico(form)
            
        End Select
    
    Case "Realizar liga��o de �gua"
        strTopicos = ColherTopicosContRealizarLigacao(form)
        
    Case "Negativa��o no SPC"
        strTopicos = ColherTopicosContNegativacao(form)
    
    Case "Cobran�a de esgoto em im�vel n�o ligado � rede"
        Select Case strTipoPeticao
        Case "Contesta��o"
            strTopicos = ColherTopicosContEsgotoSemRede(form)
            
        Case "RI/Apela��o"
            'strTopicos = ColherTopicosRIDesabGenerico(form)
            
        Case "Contrarraz�es de RI/Apela��o"
            'strTopicos = ColherTopicosCRRIDesabGenerico(form)
            
        End Select
        
    Case "Cobran�a de esgoto com �gua cortada"
        Select Case strTipoPeticao
        Case "Contesta��o"
            strTopicos = ColherTopicosContEsgotoAguaCortada(form)
            
        Case "RI/Apela��o"
            'strTopicos = ColherTopicosRIDesabGenerico(form)
            
        Case "Contrarraz�es de RI/Apela��o"
            'strTopicos = ColherTopicosCRRIDesabGenerico(form)
            
        End Select
        
    Case "Classifica��o tarifa ou qtd. de economias"
        Select Case strTipoPeticao
        Case "Contesta��o"
            strTopicos = ColherTopicosContClassificacaoTarifaria(form)
            
        Case "RI/Apela��o"
            'strTopicos = ColherTopicosRIDesabGenerico(form)
            
        Case "Contrarraz�es de RI/Apela��o"
            'strTopicos = ColherTopicosCRRIDesabGenerico(form)
            
        End Select
        
    Case "Suspeita de by-pass"
        Select Case strTipoPeticao
        Case "Contesta��o"
            strTopicos = ColherTopicosContGato(form)
            
            'O t�pico de gr�fico est� sempre
            bolGrafico = True
            
        Case "RI/Apela��o"
            'strTopicos = ColherTopicosRIDesabGenerico(form)
            
        Case "Contrarraz�es de RI/Apela��o"
            'strTopicos = ColherTopicosCRRIDesabGenerico(form)
            
        End Select
        
    Case "D�bito de terceiro"
        Select Case strTipoPeticao
        Case "Contesta��o"
            strTopicos = ColherTopicosContDebitoTerceiro(form)
            
        Case "RI/Apela��o"
            'strTopicos = ColherTopicosRIDesabGenerico(form)
            
        Case "Contrarraz�es de RI/Apela��o"
            'strTopicos = ColherTopicosCRRIDesabGenerico(form)
            
        End Select
        
    Case "Vaz. �gua ou extravas. esgoto com danos a patrim�nio/morais", "Obra da Embasa com danos a patrim�nio/morais", _
        "Acidente com pessoa/ve�culo em buraco", "Acidente com ve�culo (colis�o ou atropelamento)"
        Select Case strTipoPeticao
        Case "Contesta��o"
            strTopicos = ColherTopicosContRespCivil(form, strCausaPedir)
            
        Case "RI/Apela��o"
            'strTopicos = ColherTopicosRIRespCivil(form)
            
        Case "Contrarraz�es de RI/Apela��o"
            'strTopicos = ColherTopicosCRRIRespCivil(form)
            
        End Select
        
    Case "Desabastecimentos por per�odo e causa determinados", "Desabastecimento CCR 04/2015", "Desabastecimento Uruguai 09/2016", _
        "Desabastecimento Liberdade 10/2017", "Desabastecimento Apag�o Xingu 03/2018"
        Select Case strTipoPeticao
        Case "Contesta��o"
            strTopicos = ColherTopicosContDesabastecimento(form)
            
        Case "RI/Apela��o"
            strTopicos = ColherTopicosRIDesabastecimento(form)
            
        Case "Contrarraz�es de RI/Apela��o"
            'GoTo N�oFaz
            'strTopicos = ColherTopicosCRRIDesabGenerico(form)
            
        End Select
        
        If InStr(1, strTopicos, ",,40,,") <> 0 Or InStr(1, strTopicos, ",,42,,") <> 0 Then bolGrafico = True 'Verifica se � um dos t�picos que cont�m gr�ficos
            
    Case "Fixo de esgoto"
        Select Case strTipoPeticao
        Case "Contesta��o"
            strTopicos = ColherTopicosContFixoEsgoto(form, strCausaPedir)
            
        Case "RI/Apela��o"
            'strTopicos = ColherTopicosRIDesabApagXingu2018(form)
            
        Case "Contrarraz�es de RI/Apela��o"
            'strTopicos = ColherTopicosCRRIDesabApagXingu2018(form)
            
        End Select
        
    Case "Outros v�cios"
        Select Case strTipoPeticao
        Case "Contesta��o"
            strTopicos = ColherTopicosContVicio(form, strCausaPedir)
            
        Case "RI/Apela��o"
            'strTopicos = ColherTopicosRIDesabApagXingu2018(form)
            
        Case "Contrarraz�es de RI/Apela��o"
            'strTopicos = ColherTopicosCRRIDesabApagXingu2018(form)
            
        End Select
        
    End Select
    
    ColherTopicosGeral = strTopicos

End Function

Function ColherTopicosContPedidosIniciais(form As Variant) As String
''
'' Metafun��o para ser usado em outras fun��es de colher t�picos. Colhe os t�picos exclusivamente da frame de pedidos iniciais.
''
    ' Agrupa os n�meros dos t�picos numa string, cercados por v�rgulas duplas.
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
'' Metafun��o para ser usado em outras fun��es de colher t�picos. Colhe os t�picos exclusivamente da frame de pedidos gerais.
''
    ' Agrupa os n�meros dos t�picos numa string, cercados por v�rgulas duplas.
    Dim strTopicos As String
    
    With form
        ''Devolu��o em dobro
        If .chbDevolDobro.Value = True Then
            If .cmbPagamento.Value = "Houve pagamento" Then
                strTopicos = strTopicos & "800,,897,,"
            ElseIf .cmbPagamento.Value = "N�o houve pagamento" Then
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
        
        ''Litig�ncia de m�-f�
        If .chbLitigMaFe.Value = True Then strTopicos = strTopicos & "880,,899,,"
        
    End With
    
    ColherTopicosContPedidosFinais = strTopicos

End Function

Function ColherTopicosContRevisaoConsumo(form As Variant) As String
''
'' Recolhe os t�picos para a Contesta��o de Revis�o de Consumo. Aqui n�o importa a ordem (a ordem
''    dos t�picos ser� a que est� na planilha de apoio).
''

    ' Agrupa os n�meros dos t�picos numa string, cercados por v�rgulas duplas.
    Dim strTopicos As String
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
       'Perfil de consumo
        If .cmbPadraoConsumo.Value = "H� padr�o, mas o consumo impugnado � id�ntico ou menor que a m�dia de consumo" Then
            strTopicos = strTopicos & "25,,95,,155,,"
        ElseIf .cmbPadraoConsumo.Value = "H� padr�o, mas o consumo impugnado � razoavelmente compat�vel com a m�dia" Then
            strTopicos = strTopicos & "55,,"
        ElseIf .cmbPadraoConsumo.Value = "N�o h� padr�o definido, consumo cheio de altos e baixos" Then
            strTopicos = strTopicos & "50,,"
        ElseIf .cmbPadraoConsumo.Value = "Sem padr�o anterior, impugna consumos medidos ap�s in�cio do contrato pelo m�nimo" Then
            strTopicos = strTopicos & "60,,"
        ElseIf .cmbPadraoConsumo.Value = "Sem padr�o anterior, impugna consumos medidos ap�s longo tempo sem hidr�metro" Then
            strTopicos = strTopicos & "61,,"
        End If
        
        'H� Medi��o individualizada, mas consumo rateado n�o foi relevante no aumento
        If .chbMIIrrelevante.Value = True Then strTopicos = strTopicos & "5,,"
        
        'Defender parcelamento
        If .chbParcelamento.Value = True Then strTopicos = strTopicos & "125,,"
        
        ' Defender cobran�a de esgoto durante corte
        If .chbEsgotoAposCorte.Value = True Then strTopicos = strTopicos & "65,,"
        
        'Consumo de acordo com a M�dia do Procon/SP
        If .chbMediaProcon.Value = True Then strTopicos = strTopicos & "110,,"
        
        'Exist�ncia de aferi��o
        If .cmbAferHidrometro.Value = "H�, hidr�metro regular" Then
            strTopicos = strTopicos & "75,,180,,"
        ElseIf .cmbAferHidrometro.Value = "H�, irregular contra a fornecedora (medindo a menor)" Then
            strTopicos = strTopicos & "80,,185,,"
        End If
        
        'Requerer aferi��o
        If .chbRequerAfericao.Value = True Then
            If .chbMIIrrelevante.Value = True Then
                strTopicos = strTopicos & "70,,170,,"
            Else
                strTopicos = strTopicos & "65,,165,,"
            End If
        End If
        
        'Vazamento interno
        If .cmbVazInterno.Value = "H� SS no SCI (anexar SS!)" Then
            strTopicos = strTopicos & "15,,30,,90,,150,,"
        ElseIf .cmbVazInterno.Value = "H� confiss�o" Then
            strTopicos = strTopicos & "10,,23,,85,,150,,"
        End If
        
        'Hidr�metro j� foi substitu�do, hidr�metro novo corrobora medi��es
        If .chbHidrTrocado.Value = True Then
            strTopicos = strTopicos & "100,,"
        End If
        
        'M�dia de consumo viciada por processos anteriores
        If .chbMediaViciada.Value = True Then strTopicos = strTopicos & "105,,"
        
        'M�dia de consumo correta
        If .chbMediaCorreta.Value = True Then
            If .chbMediaConsRetificado.Value = True Then
                strTopicos = strTopicos & "115,,190,,"
            Else
                strTopicos = strTopicos & "120,,190,,"
            End If
        End If
        
        'Se j� n�o tem t�pico de m�rito com hidr�metro regular ou a menor, coloca o gen�rico
        If InStr(1, strTopicos, ",,180,,") = 0 And InStr(1, strTopicos, ",,185,,") = 0 Then strTopicos = strTopicos & "175,,"
        
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosFinais(form)
        
    End With
    
    ColherTopicosContRevisaoConsumo = strTopicos

End Function

Function ColherTopicosContCorte(form As Variant) As String
''
'' Recolhe os t�picos para a Contesta��o de Corte. Aqui n�o importa a ordem (a ordem
''    dos t�picos ser� a que est� na planilha de apoio).
''

    ' Agrupa os n�meros dos t�picos numa string, cercados por v�rgulas duplas.
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
        If .cmbAvisoCorte.Value = "Houve, em faturas anteriores / n�o houve" Then
            strTopicos = strTopicos & "50,,"
        ElseIf .cmbAvisoCorte.Value = "Houve, em correspond�ncia espec�fica" Then
            strTopicos = strTopicos & "45,,"
        End If
        
        ' Pedido (t�pico especial para pagamento na v�spera)
        If .chbPagVespera.Value = True Then strTopicos = strTopicos & "80,," Else strTopicos = strTopicos & "70,,"
        
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosFinais(form)
        
    End With
    
    ColherTopicosContCorte = strTopicos

End Function

Function ColherTopicosContNegativacao(form As Variant) As String
''
'' Recolhe os t�picos para a Contesta��o de Realizar liga��o. Aqui n�o importa a ordem (a ordem
''    dos t�picos ser� a que est� na planilha de apoio).
''

    ' Agrupa os n�meros dos t�picos numa string, cercados por v�rgulas duplas.
    Dim strTopicos As String
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
        'Atitude do Autor
        If .cmbAtitudeAutor.Value = "Autor afirma, de forma gen�rica, ""desconhecer"" d�vida" Then
            strTopicos = strTopicos & "20,,80,,72,,90,,"
        ElseIf .cmbAtitudeAutor.Value = "Autor afirma claramente que n�o firmou contrato" Then
            strTopicos = strTopicos & "10,,72,,90,,"
        ElseIf .cmbAtitudeAutor.Value = "Autor reconhece contrato mas nega d�bitos" Then
            strTopicos = strTopicos & "15,,"
        ElseIf .cmbAtitudeAutor.Value = "Autor diz que se mudou e pediu suspens�o do servi�o" Then
            strTopicos = strTopicos & "17,,68,,72,,90,,"
        End If
        
        'Perfil do contrato
        If .cmbPerfilContrato.Value = "Sem uso nem pagamento, apar�ncia de fraude (n�o comentar)" Then
            strTopicos = strTopicos & "35,,"
        ElseIf .cmbPerfilContrato.Value = "H� uso e pagamentos mais ou menos regulares" Then
            strTopicos = strTopicos & "40,,80,,"
        ElseIf .cmbPerfilContrato.Value = "Houve uso e pagamentos regulares at� certa data" Then
            strTopicos = strTopicos & "42,,80,,"
        ElseIf .cmbPerfilContrato.Value = "A negativa��o foi de parcelamento firmado pelo Autor" Then
            strTopicos = strTopicos & "45,,"
        End If
        
        ' Prova da negativa��o
        If .chbProvaNegativacao.Value = True Then
            If .cmbAtitudeAutor.Value <> "Autor diz que se mudou e pediu suspens�o do servi�o" Then strTopicos = strTopicos & "27,,"
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
        
        'Endere�o da qualifica��o e contrato com a Coelba
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
'' Recolhe os t�picos para a Contesta��o de Realizar liga��o. Aqui n�o importa a ordem (a ordem
''    dos t�picos ser� a que est� na planilha de apoio).
''

    ' Agrupa os n�meros dos t�picos numa string, cercados por v�rgulas duplas.
    Dim strTopicos As String
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
        'Particularidades do caso
        If (.chbNaoHouveSolicitacao.Value = True Or .chbNaoHouveRecusa.Value = True) Then
            strTopicos = strTopicos & "10,," 'Cabe�alho dos esclarecimentos simples
            If .chbNaoHouveSolicitacao.Value = True Then strTopicos = strTopicos & "13,,28,,110,,"
            If .chbNaoHouveRecusa.Value = True Then strTopicos = strTopicos & "15,,115,,"
            
        Else
            strTopicos = strTopicos & "20,," 'Cabe�alho das pretens�es at�cnicas
            If .chbSemReservatorioBomba.Value = True Then strTopicos = strTopicos & "21,,30,,100,,"
            If .chbSemReservacao.Value = True Then strTopicos = strTopicos & "22,,31,,100,,"
            If .chbSepararInstalacoesInternas.Value = True Then strTopicos = strTopicos & "23,,32,,100,,"
            If .chbDistanciaRede.Value = True Then strTopicos = strTopicos & "24,,33,,100,,"
            If .chbAltitudeInsuficiente.Value = True Then strTopicos = strTopicos & "34,,100,,"
            
        End If
        
        'Pedidos espec�ficos de Realizar liga��o
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
'' Recolhe os t�picos para a Contesta��o de Cobran�a de esgoto sem liga��o � rede. Aqui n�o importa a ordem (a ordem
''    dos t�picos ser� a que est� na planilha de apoio).
''

    ' Agrupa os n�meros dos t�picos numa string, cercados por v�rgulas duplas.
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
'' Recolhe os t�picos para a Contesta��o de Cobran�a de esgoto com �gua cortada. Aqui n�o importa a ordem (a ordem
''    dos t�picos ser� a que est� na planilha de apoio).
''

    ' Agrupa os n�meros dos t�picos numa string, cercados por v�rgulas duplas.
    Dim strTopicos As String
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
        'Prova da habita��o e uso do im�vel
        If .cmbProvaUsoImovel.Value = "Autor n�o alega im�vel desabitado; h� inspe��es" Then strTopicos = strTopicos & "10,,24,,35,,50,,"
        If .cmbProvaUsoImovel.Value = "Autor confessa que im�vel estava habitado" Then strTopicos = strTopicos & "10,,20,,35,,"
        If .cmbProvaUsoImovel.Value = "Inspe��es realizadas pelos t�cnicos da Embasa" Then strTopicos = strTopicos & "10,,28,,35,,50,,"
        
        'Pedir depoimento pessoal
        If .cmbProvaUsoImovel.Value <> "Autor confessa que im�vel estava habitado" Then strTopicos = strTopicos & "95,,"
        
        'Requerimento administrativo
        If .chbSemRequerimentoAdm.Value = True Then strTopicos = strTopicos & "60,,"
        
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosFinais(form)
        
    End With
    
    ColherTopicosContEsgotoAguaCortada = strTopicos

End Function

Function ColherTopicosContClassificacaoTarifaria(form As Variant) As String
''
'' Recolhe os t�picos para a Contesta��o de Classifica��o tarif�ria. Aqui n�o importa a ordem (a ordem
''    dos t�picos ser� a que est� na planilha de apoio).
''

    ' Agrupa os n�meros dos t�picos numa string, cercados por v�rgulas duplas.
    Dim strTopicos As String
    Dim intAno As Integer, btMes As Byte
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
        'Pretens�o - alega��o de multiplica��o do m�nimo por m�ltiplas economias?
        If .optPretMultEcon.Value = True Then
            strTopicos = strTopicos & "5,,"
        ElseIf .optPretUmaEcon.Value = True Then
            strTopicos = strTopicos & "10,,35,,"
        End If
        
        'Necessidade de esclarecer per�odo?
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
        
        'Tarifa aplic�vel
        If .chbApresentarCalcExemplo.Value = True Then
            strTopicos = strTopicos & "40,,80,,"
            'Descobre qual a Resolu��o de tarifa
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
        
        ' C�lculo em si
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
'' Recolhe os t�picos para a Contesta��o de Gato. Aqui n�o importa a ordem (a ordem
''    dos t�picos ser� a que est� na planilha de apoio).
''

    ' Agrupa os n�meros dos t�picos numa string, cercados por v�rgulas duplas.
    Dim strTopicos As String
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
        'Tipo de gato
        If .cmbTipoGato.Value = "Desvio - bypass" Then strTopicos = strTopicos & "10,,"
        If .cmbTipoGato.Value = "Hidr�metro furado" Then strTopicos = strTopicos & "20,,"
        If .cmbTipoGato.Value = "Hidr�metro invertido" Then strTopicos = strTopicos & "22,,"
        
        'Processo administrativo
        If .chbProcessoAdm.Value = True Then strTopicos = strTopicos & "30,,"
        
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosFinais(form)
        
    End With
    
    ColherTopicosContGato = strTopicos

End Function

Function ColherTopicosContDebitoTerceiro(form As Variant) As String
''
'' Recolhe os t�picos para a Contesta��o de D�bitos de terceiro. Aqui n�o importa a ordem (a ordem
''    dos t�picos ser� a que est� na planilha de apoio).
''

    ' Agrupa os n�meros dos t�picos numa string, cercados por v�rgulas duplas.
    Dim strTopicos As String
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
        'Sobre as provas de n�o utiliza��o do servi�o
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
'' Recolhe os t�picos para as Contesta��es de Responsabilidade civil. Aqui n�o importa a ordem (a ordem
''    dos t�picos ser� a que est� na planilha de apoio).
''

    ' Agrupa os n�meros dos t�picos numa string, cercados por v�rgulas duplas.
    Dim strTopicos As String
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
        'Particularidades do caso
        Select Case strCausaPedir
        Case "Vaz. �gua ou extravas. esgoto com danos a patrim�nio/morais"
            If .cmbOcorrencia.Value = "Fato pontual" Then
                strTopicos = strTopicos & "5,,"
            ElseIf .cmbOcorrencia.Value = "Reiterada em grande per�odo de tempo" Then
                strTopicos = strTopicos & "10,,"
            End If
        
        Case "Acidente com pessoa/ve�culo em buraco"
            If .cmbOcorrencia.Value = "Ve�culo" Then
                strTopicos = strTopicos & "5,,"
            ElseIf .cmbOcorrencia.Value = "Pessoa" Then
                strTopicos = strTopicos & "10,,"
            End If
        
        Case "Acidente com ve�culo (colis�o ou atropelamento)"
            If .cmbOcorrencia.Value = "Colis�o" Then
                strTopicos = strTopicos & "5,,"
            ElseIf .cmbOcorrencia.Value = "Atropelamento" Then
                strTopicos = strTopicos & "10,,"
            End If
        
        End Select
        
        'Prescri��o
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
'' Recolhe os t�picos para a Contesta��o de desabastecimento gen�rico. Aqui n�o importa a ordem (a ordem
''    dos t�picos ser� a que est� na planilha de apoio).
''

    ' Agrupa os n�meros dos t�picos numa string, cercados por v�rgulas duplas.
    Dim strTopicos As String
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
        'Correspons�vel
        If .cmbCorresponsavel.Value = "N�o houve outro respons�vel" Or .cmbCorresponsavel.Value = "" Then
             strTopicos = strTopicos & "98,,"
        Else
             strTopicos = strTopicos & "96,,"
        End If
    
        'Prescri��o trienal
        If .chbPrescricao.Value = True Then strTopicos = strTopicos & "15,,93,,"
        
        'Particularidades do caso
        If .chbMultiplosProcessos.Value = True Then strTopicos = strTopicos & "10,,"
        If .chbSemFatura.Value = True Then strTopicos = strTopicos & "60,,"
        If .chbBairroNaoAfetado.Value = True Then strTopicos = strTopicos & "20,,"
        If .chbInadimplenteEpoca.Value = True Then strTopicos = strTopicos & "63,,"
        If .chbReservatorio.Value = True Then strTopicos = strTopicos & "65,,"
        If .chbSemRequerimentoAdm.Value = True Then strTopicos = strTopicos & "50,,"
        
        'Altera��o de consumo
        If .cmbAlteracaoConsumo.Value = "N�o houve altera��o relevante [gr�fico]" Then
            strTopicos = strTopicos & "40,,"
        ElseIf .cmbAlteracaoConsumo.Value = "N�o houve altera��o [print de contas juntadas pelo Autor]" Then
            strTopicos = strTopicos & "41,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Consumo aumentou no per�odo do acidente [gr�fico]" Then
            strTopicos = strTopicos & "42,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Abastecimento estava cortado por outro motivo" Then
            strTopicos = strTopicos & "45,,"
        End If
        
        'T�pico de danos morais adicional, espec�fico de desabastecimento
        If form.chbDanoMoral.Value = True Then
            If .cmbCorresponsavel.Value = "N�o houve outro respons�vel" Or .cmbCorresponsavel.Value = "" Then
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
'' Recolhe os t�picos para a Contesta��o das outras causas de pedir. Aqui n�o importa a ordem (a ordem
''    dos t�picos ser� a que est� na planilha de apoio).
''

    ' Agrupa os n�meros dos t�picos numa string, cercados por v�rgulas duplas.
    Dim strTopicos As String
    Dim intAno As Integer
    Dim btMes As Byte
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
        'Postura do Autor
        If .cmbConfessaPoco.Value = "Confessa que existe po�o" Or .cmbConfessaPoco.Value = "Confessa que foi instalado hidr�metro no po�o" Then
            strTopicos = strTopicos & "5,,"
        Else
            strTopicos = strTopicos & "10,,"
        End If
        
        'Postura do Autor
        If .cmbConfessaPoco.Value = "Confessa que existe po�o" Then
            strTopicos = strTopicos & "5,,"
        ElseIf .cmbConfessaPoco.Value = "Confessa que foi instalado hidr�metro no po�o" Then
            strTopicos = strTopicos & "5,,15,,20,,"
        Else
            strTopicos = strTopicos & "10,,"
        End If
        
        'Necessidade de explicar volume paradigma
        If .cmbVolumeParadigma.Value = "M�dia anterior � instala��o do po�o" Then
            strTopicos = strTopicos & "35,,"
        ElseIf .cmbVolumeParadigma.Value = "Menor do que a m�dia anterior ao po�o" Then
            strTopicos = strTopicos & "40,,"
        End If
        
        'Particularidades do caso
        If .chbFotosHidrometroPoco.Value = True Then strTopicos = strTopicos & "15,,25,,"
        If .chbImpediuInstalacaoHidrometro.Value = True Then strTopicos = strTopicos & "30,,"
        If .chbNaoAplicaCDC.Value = True Then strTopicos = strTopicos & "55,,"
        
        'Descobre qual a Resolu��o de tarifa
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
        
        ' C�lculo em si
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
'' Recolhe os t�picos para a Contesta��o das outras causas de pedir. Aqui n�o importa a ordem (a ordem
''    dos t�picos ser� a que est� na planilha de apoio).
''

    ' Agrupa os n�meros dos t�picos numa string, cercados por v�rgulas duplas.
    Dim strTopicos As String
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
        'Dos fatos
        Select Case strCausaPedir
        Case "Incidentes com cobran�a de m�dia"
            strTopicos = strTopicos & "5,,"
            
        Case "Desligar liga��o de �gua"
            strTopicos = strTopicos & "10,,"
            
        Case "Hidr�metro do im�vel trocado"
            strTopicos = strTopicos & "15,,"
            
        Case Else
            strTopicos = strTopicos & "5,,"
            
        End Select
        
        'Mero v�cio?
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
'' Recolhe os t�picos para o Recurso Inominado de desabastecimento. Aqui n�o importa a ordem (a ordem
''    dos t�picos ser� a que est� na planilha de apoio).
''

    ' Agrupa os n�meros dos t�picos numa string, cercados por v�rgulas duplas.
    Dim strTopicos As String
    
    With form
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosIniciais(form)
        
        'Correspons�vel
        If .cmbCorresponsavel.Value = "N�o houve outro respons�vel" Or .cmbCorresponsavel.Value = "" Then
             strTopicos = strTopicos & "98,,"
        Else
             strTopicos = strTopicos & "96,,"
        End If
    
        'Prescri��o trienal
        If .chbPrescricao.Value = True Then strTopicos = strTopicos & "15,,93,,"
        
        'Particularidades do caso
        If .chbMultiplosProcessos.Value = True Then strTopicos = strTopicos & "10,,"
        If .chbSemFatura.Value = True Then strTopicos = strTopicos & "60,,"
        If .chbBairroNaoAfetado.Value = True Then strTopicos = strTopicos & "20,,"
        If .chbInadimplenteEpoca.Value = True Then strTopicos = strTopicos & "63,,"
        If .chbReservatorio.Value = True Then strTopicos = strTopicos & "65,,"
        If .chbSemRequerimentoAdm.Value = True Then strTopicos = strTopicos & "50,,"
        
        'Altera��o de consumo
        If .cmbAlteracaoConsumo.Value = "N�o houve altera��o relevante [gr�fico]" Then
            strTopicos = strTopicos & "40,,"
        ElseIf .cmbAlteracaoConsumo.Value = "N�o houve altera��o [print de contas juntadas pelo Autor]" Then
            strTopicos = strTopicos & "41,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Consumo aumentou no per�odo do acidente [gr�fico]" Then
            strTopicos = strTopicos & "42,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Abastecimento estava cortado por outro motivo" Then
            strTopicos = strTopicos & "45,,"
        End If
        
        'T�pico de danos morais adicional, espec�fico de desabastecimento
        If form.chbDanoMoral.Value = True Then
            If .cmbCorresponsavel.Value = "N�o houve outro respons�vel" Or .cmbCorresponsavel.Value = "" Then
                 strTopicos = strTopicos & "75,,"
            Else
                 strTopicos = strTopicos & "70,,"
            End If
        End If
        
        'Pedidos gerais
        strTopicos = strTopicos & ColherTopicosContPedidosFinais(form)
        
        'Quest�es recursais
        strTopicos = Replace(strTopicos, ",,800,,", ",,") ' Remover dano material e devolu��o em dobro decorrentes
        strTopicos = Replace(strTopicos, ",,810,,", ",,") ' do pedido; o que contar� � o da condena��o, se tiver havido.

        If .chbJurosEventoDanoso.Value = True Then strTopicos = strTopicos & "330,,395,,"
        If .chbCondenouDevDobro.Value = True Then strTopicos = strTopicos & "800,,"
        If .chbCondenouDanosMateriais.Value = True Then strTopicos = strTopicos & "810,,"
        If .chbExcluiuCorresp.Value = True Then strTopicos = strTopicos & "310,,390,,"
        
        ' Testemunhas
        If .cmbTestemunhas.Value = "Autor n�o produziu prova testemunhal" Then
            strTopicos = strTopicos & "320,,"
        ElseIf .cmbTestemunhas.Value = "Testemunha imprecisa ou inveross�mil" Then
            strTopicos = strTopicos & "323,,"
        ElseIf .cmbTestemunhas.Value = "Testemunha interessada, tem processo" Then
            strTopicos = strTopicos & "326,,"
        End If
    
    End With
    
    ColherTopicosRIDesabastecimento = strTopicos

End Function

Function ColherTopicosRICorte(form As Variant) As String
''
'' Recolhe os t�picos para o Recurso Inominado de Corte. Aqui n�o importa a ordem (a ordem
''    dos t�picos ser� a que est� na planilha de apoio).
''

    ' Agrupa os n�meros dos t�picos numa string, cercados por v�rgulas duplas.
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
        If .cmbAvisoCorte.Value = "Houve, em faturas anteriores / n�o houve" Then
            strTopicos = strTopicos & "57,,"
        ElseIf .cmbAvisoCorte.Value = "Houve, em correspond�ncia espec�fica" Then
            strTopicos = strTopicos & "53,,"
        End If
        
        ' Pedido (t�pico especial para pagamento na v�spera)
        If .chbPagVespera.Value = True Then strTopicos = strTopicos & "80,," Else strTopicos = strTopicos & "70,,"
        
        'Senten�a
        If .chbJurosEventoDanoso.Value = True Then strTopicos = strTopicos & "63,,110,,"
        If .chbCondenouDanosMateriais.Value = True Then strTopicos = strTopicos & "59,,100,,"
        
    End With
    
    ColherTopicosRICorte = strTopicos

End Function

Function ColherTopicosRIDesabGenerico(form As Variant) As String
''
'' Recolhe os t�picos para o Recurso Inominado de CCR. Aqui n�o importa a ordem (a ordem
''    dos t�picos ser� a que est� na planilha de apoio).
''

    ' Agrupa os n�meros dos t�picos numa string, cercados por v�rgulas duplas.
    Dim strTopicos As String
    strTopicos = ",,"
    
    
    With form
        'Correspons�vel
        If .cmbCorresponsavel.Value = "N�o houve outro respons�vel" Then
             strTopicos = strTopicos & "87,,98,,"
        Else
             strTopicos = strTopicos & "85,,96,,"
        End If
    
        'Ilegitimidade ativa
        If .chbIlegitimidade.Value = True Then strTopicos = strTopicos & "10,,90,,"
        
        'Prescri��o trienal
        If .chbPrescricao.Value = True Then strTopicos = strTopicos & "17,,97,,"
        
        'Particularidades do caso
        If .chbInadimplenteEpoca.Value = True Then strTopicos = strTopicos & ",,"
        If .chbSemFatura.Value = True Then strTopicos = strTopicos & "60,,"
        If .chbBairroNaoAfetado.Value = True Then strTopicos = strTopicos & "20,,"
        If .chbReservatorio.Value = True Then strTopicos = strTopicos & "65,,"
        
        'Altera��o de consumo
        If .cmbAlteracaoConsumo.Value = "N�o houve altera��o relevante [gr�fico]" Then
            strTopicos = strTopicos & "40,,"
        ElseIf .cmbAlteracaoConsumo.Value = "N�o houve altera��o [print de contas juntadas pelo Autor]" Then
            strTopicos = strTopicos & "41,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Consumo aumentou no per�odo do acidente [gr�fico]" Then
            strTopicos = strTopicos & "42,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Abastecimento estava cortado por outro motivo" Then
            strTopicos = strTopicos & "45,,"
        End If
        
        'Pedidos diferentes
        If .chbCondenouDanosMateriais.Value = True Then strTopicos = strTopicos & "70,,"
        If .chbCondenouDevDobro.Value = True Then strTopicos = strTopicos & "80,,"
        
        'Testemunha
        If .cmbTestemunhas.Value = "Autor n�o produziu prova testemunhal" Then
            strTopicos = strTopicos & "30,,"
        ElseIf .cmbTestemunhas.Value = "Testemunha imprecisa ou inveross�mil" Then
            strTopicos = strTopicos & "33,,"
        ElseIf .cmbTestemunhas.Value = "Testemunha interessada, tem processo" Then
            strTopicos = strTopicos & "36,,"
        End If
        
        'Senten�a
        If .chbExcluiuCorresp.Value = True Then strTopicos = strTopicos & "17,,97,,"
        If .chbJurosEventoDanoso.Value = True Then strTopicos = strTopicos & "63,,110,,"
        If .chbCondenouDanosMateriais.Value = True Then strTopicos = strTopicos & "70,,100,,"
        If .chbCondenouDevDobro.Value = True Then strTopicos = strTopicos & "80,,105,,"
        
    End With
    
    ColherTopicosRIDesabGenerico = strTopicos

End Function

Function ColherTopicosRICCR(form As Variant) As String
''
'' Recolhe os t�picos para o Recurso Inominado de CCR. Aqui n�o importa a ordem (a ordem
''    dos t�picos ser� a que est� na planilha de apoio).
''

    ' Agrupa os n�meros dos t�picos numa string, cercados por v�rgulas duplas.
    Dim strTopicos As String
    strTopicos = ",,"
    
    
    With form
        'Ilegitimidade ativa
        If .chbIlegitimidade.Value = True Then strTopicos = strTopicos & "10,,90,,"
        
        'Prescri��o trienal
        If .chbPrescricao.Value = True Then strTopicos = strTopicos & "17,,97,,"
        
        'Particularidades do caso
        If .chbInadimplenteEpoca.Value = True Then strTopicos = strTopicos & ",,"
        If .chbSemFatura.Value = True Then strTopicos = strTopicos & "60,,"
        If .chbBairroNaoAfetado.Value = True Then strTopicos = strTopicos & "20,,"
        If .chbReservatorio.Value = True Then strTopicos = strTopicos & "65,,"
        
        'Altera��o de consumo
        If .cmbAlteracaoConsumo.Value = "N�o houve altera��o relevante [gr�fico]" Then
            strTopicos = strTopicos & "40,,"
        ElseIf .cmbAlteracaoConsumo.Value = "N�o houve altera��o [print de contas juntadas pelo Autor]" Then
            strTopicos = strTopicos & "41,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Consumo aumentou no per�odo do acidente [gr�fico]" Then
            strTopicos = strTopicos & "42,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Abastecimento estava cortado por outro motivo" Then
            strTopicos = strTopicos & "45,,"
        End If
        
        'Pedidos diferentes
        If .chbCondenouDanosMateriais.Value = True Then strTopicos = strTopicos & "70,,"
        If .chbCondenouDevDobro.Value = True Then strTopicos = strTopicos & "80,,"
        
        'Testemunha
        If .cmbTestemunhas.Value = "Autor n�o produziu prova testemunhal" Then
            strTopicos = strTopicos & "30,,"
        ElseIf .cmbTestemunhas.Value = "Testemunha imprecisa ou inveross�mil" Then
            strTopicos = strTopicos & "33,,"
        ElseIf .cmbTestemunhas.Value = "Testemunha interessada, tem processo" Then
            strTopicos = strTopicos & "36,,"
        End If
        
        'Senten�a
        If .chbExcluiuCorresp.Value = True Then strTopicos = strTopicos & "17,,97,,"
        If .chbJurosEventoDanoso.Value = True Then strTopicos = strTopicos & "63,,110,,"
        If .chbCondenouDanosMateriais.Value = True Then strTopicos = strTopicos & "70,,100,,"
        If .chbCondenouDevDobro.Value = True Then strTopicos = strTopicos & "80,,105,,"
        
    End With
    
    ColherTopicosRICCR = strTopicos

End Function

Function ColherTopicosRIDesabUruguai2016(form As Variant) As String
''
'' Recolhe os t�picos para o Recurso Inominado de Desabastecimento do Apag�o Xingu. Aqui n�o importa a ordem (a ordem
''    dos t�picos ser� a que est� na planilha de apoio).
''

    ' Agrupa os n�meros dos t�picos numa string, cercados por v�rgulas duplas.
    Dim strTopicos As String
    strTopicos = ",,"
    
    
    With form
        'Ilegitimidade ativa
        If .chbIlegitimidade.Value = True Then strTopicos = strTopicos & "10,,90,,"
        
        'Prescri��o trienal
        If .chbPrescricao.Value = True Then strTopicos = strTopicos & "17,,97,,"
        
        'Particularidades do caso
        If .chbInadimplenteEpoca.Value = True Then strTopicos = strTopicos & ",,"
        If .chbSemFatura.Value = True Then strTopicos = strTopicos & "60,,"
        If .chbBairroNaoAfetado.Value = True Then strTopicos = strTopicos & "20,,"
        If .chbReservatorio.Value = True Then strTopicos = strTopicos & "65,,"
        
        'Altera��o de consumo
        If .cmbAlteracaoConsumo.Value = "N�o houve altera��o relevante [gr�fico]" Then
            strTopicos = strTopicos & "40,,"
        ElseIf .cmbAlteracaoConsumo.Value = "N�o houve altera��o [print de contas juntadas pelo Autor]" Then
            strTopicos = strTopicos & "41,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Consumo aumentou no per�odo do acidente [gr�fico]" Then
            strTopicos = strTopicos & "42,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Abastecimento estava cortado por outro motivo" Then
            strTopicos = strTopicos & "45,,"
        End If
        
        'Pedidos diferentes
        If .chbCondenouDanosMateriais.Value = True Then strTopicos = strTopicos & "70,,"
        If .chbCondenouDevDobro.Value = True Then strTopicos = strTopicos & "80,,"
        
        'Testemunha
        If .cmbTestemunhas.Value = "Autor n�o produziu prova testemunhal" Then
            strTopicos = strTopicos & "30,,"
        ElseIf .cmbTestemunhas.Value = "Testemunha imprecisa ou inveross�mil" Then
            strTopicos = strTopicos & "33,,"
        ElseIf .cmbTestemunhas.Value = "Testemunha interessada, tem processo" Then
            strTopicos = strTopicos & "36,,"
        End If
        
        'Senten�a
        If .chbJurosEventoDanoso.Value = True Then strTopicos = strTopicos & "63,,110,,"
        If .chbCondenouDanosMateriais.Value = True Then strTopicos = strTopicos & "70,,100,,"
        If .chbCondenouDevDobro.Value = True Then strTopicos = strTopicos & "80,,105,,"
        
    End With
    
    ColherTopicosRIDesabUruguai2016 = strTopicos

End Function

Function ColherTopicosRIDesabLiberdade2017(form As Variant) As String
''
'' Recolhe os t�picos para o Recurso Inominado de Desabastecimento do Apag�o Xingu. Aqui n�o importa a ordem (a ordem
''    dos t�picos ser� a que est� na planilha de apoio).
''

    ' Agrupa os n�meros dos t�picos numa string, cercados por v�rgulas duplas.
    Dim strTopicos As String
    strTopicos = ",,"
    
    
    With form
        'Ilegitimidade ativa
        If .chbIlegitimidade.Value = True Then strTopicos = strTopicos & "10,,90,,"
        
        'Prescri��o trienal
        If .chbPrescricao.Value = True Then strTopicos = strTopicos & "17,,97,,"
        
        'Particularidades do caso
        If .chbInadimplenteEpoca.Value = True Then strTopicos = strTopicos & ",,"
        If .chbSemFatura.Value = True Then strTopicos = strTopicos & "60,,"
        If .chbBairroNaoAfetado.Value = True Then strTopicos = strTopicos & "20,,"
        If .chbReservatorio.Value = True Then strTopicos = strTopicos & "65,,"
        
        'Altera��o de consumo
        If .cmbAlteracaoConsumo.Value = "N�o houve altera��o relevante [gr�fico]" Then
            strTopicos = strTopicos & "40,,"
        ElseIf .cmbAlteracaoConsumo.Value = "N�o houve altera��o [print de contas juntadas pelo Autor]" Then
            strTopicos = strTopicos & "41,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Consumo aumentou no per�odo do acidente [gr�fico]" Then
            strTopicos = strTopicos & "42,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Abastecimento estava cortado por outro motivo" Then
            strTopicos = strTopicos & "45,,"
        End If
        
        'Pedidos diferentes
        If .chbCondenouDanosMateriais.Value = True Then strTopicos = strTopicos & "70,,"
        If .chbCondenouDevDobro.Value = True Then strTopicos = strTopicos & "80,,"
        
        'Testemunha
        If .cmbTestemunhas.Value = "Autor n�o produziu prova testemunhal" Then
            strTopicos = strTopicos & "30,,"
        ElseIf .cmbTestemunhas.Value = "Testemunha imprecisa ou inveross�mil" Then
            strTopicos = strTopicos & "33,,"
        ElseIf .cmbTestemunhas.Value = "Testemunha interessada, tem processo" Then
            strTopicos = strTopicos & "36,,"
        End If
        
        'Senten�a
        If .chbJurosEventoDanoso.Value = True Then strTopicos = strTopicos & "63,,110,,"
        If .chbCondenouDanosMateriais.Value = True Then strTopicos = strTopicos & "70,,100,,"
        If .chbCondenouDevDobro.Value = True Then strTopicos = strTopicos & "80,,105,,"
        
    End With
    
    ColherTopicosRIDesabLiberdade2017 = strTopicos

End Function

Function ColherTopicosRIDesabApagXingu2018(form As Variant) As String
''
'' Recolhe os t�picos para o Recurso Inominado de Desabastecimento do Apag�o Xingu. Aqui n�o importa a ordem (a ordem
''    dos t�picos ser� a que est� na planilha de apoio).
''

    ' Agrupa os n�meros dos t�picos numa string, cercados por v�rgulas duplas.
    Dim strTopicos As String
    strTopicos = ",,"
    
    
    With form
        'Ilegitimidade ativa
        If .chbIlegitimidade.Value = True Then strTopicos = strTopicos & "10,,90,,"
        
        'Prescri��o trienal
        If .chbPrescricao.Value = True Then strTopicos = strTopicos & "17,,97,,"
        
        'Particularidades do caso
        If .chbInadimplenteEpoca.Value = True Then strTopicos = strTopicos & ",,"
        If .chbSemFatura.Value = True Then strTopicos = strTopicos & "60,,"
        If .chbBairroNaoAfetado.Value = True Then strTopicos = strTopicos & "20,,"
        If .chbReservatorio.Value = True Then strTopicos = strTopicos & "65,,"
        
        'Altera��o de consumo
        If .cmbAlteracaoConsumo.Value = "N�o houve altera��o relevante [gr�fico]" Then
            strTopicos = strTopicos & "40,,"
        ElseIf .cmbAlteracaoConsumo.Value = "N�o houve altera��o [print de contas juntadas pelo Autor]" Then
            strTopicos = strTopicos & "41,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Consumo aumentou no per�odo do acidente [gr�fico]" Then
            strTopicos = strTopicos & "42,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Abastecimento estava cortado por outro motivo" Then
            strTopicos = strTopicos & "45,,"
        End If
        
        'Pedidos diferentes
        If .chbCondenouDanosMateriais.Value = True Then strTopicos = strTopicos & "70,,"
        If .chbCondenouDevDobro.Value = True Then strTopicos = strTopicos & "80,,"
        
        'Testemunha
        If .cmbTestemunhas.Value = "Autor n�o produziu prova testemunhal" Then
            strTopicos = strTopicos & "30,,"
        ElseIf .cmbTestemunhas.Value = "Testemunha imprecisa ou inveross�mil" Then
            strTopicos = strTopicos & "33,,"
        ElseIf .cmbTestemunhas.Value = "Testemunha interessada, tem processo" Then
            strTopicos = strTopicos & "36,,"
        End If
        
        'Senten�a
        If .chbExcluiuCorresp.Value = True Then strTopicos = strTopicos & "15,,95,,"
        If .chbJurosEventoDanoso.Value = True Then strTopicos = strTopicos & "63,,110,,"
        If .chbCondenouDanosMateriais.Value = True Then strTopicos = strTopicos & "70,,100,,"
        If .chbCondenouDevDobro.Value = True Then strTopicos = strTopicos & "80,,105,,"
        
    End With
    
    ColherTopicosRIDesabApagXingu2018 = strTopicos

End Function

Function ColherTopicosCRRIDesabGenerico(form As Variant) As String
''
'' Recolhe os t�picos para as Contrarraz�es de Recurso Inominado de CCR. Aqui n�o importa a ordem (a ordem
''    dos t�picos ser� a que est� na planilha de apoio).
''

    ' Agrupa os n�meros dos t�picos numa string, cercados por v�rgulas duplas.
    Dim strTopicos As String
    strTopicos = ",,"
    
    
    With form
        'Correspons�vel
        If .cmbCorresponsavel.Value = "N�o houve outro respons�vel" Then
             strTopicos = strTopicos & "87,,98,,"
        Else
             strTopicos = strTopicos & "85,,96,,"
        End If
    
        'Ilegitimidade ativa
        If .chbIlegitimidade.Value = True Then strTopicos = strTopicos & "10,,"
        
        'Prescri��o trienal
        If .chbPrescricao.Value = True Then strTopicos = strTopicos & "15,,"
        
        'Particularidades do caso
        If .chbInadimplenteEpoca.Value = True Then strTopicos = strTopicos & ",,"
        If .chbSemFatura.Value = True Then strTopicos = strTopicos & "60,,"
        If .chbBairroNaoAfetado.Value = True Then strTopicos = strTopicos & "20,,"
        If .chbReservatorio.Value = True Then strTopicos = strTopicos & "65,,"
        
        'Altera��o de consumo
        If .cmbAlteracaoConsumo.Value = "N�o houve altera��o relevante [gr�fico]" Then
            strTopicos = strTopicos & "40,,"
        ElseIf .cmbAlteracaoConsumo.Value = "N�o houve altera��o [print de contas juntadas pelo Autor]" Then
            strTopicos = strTopicos & "41,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Consumo aumentou no per�odo do acidente [gr�fico]" Then
            strTopicos = strTopicos & "42,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Abastecimento estava cortado por outro motivo" Then
            strTopicos = strTopicos & "45,,"
        End If
        
        'Senten�a
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
'' Recolhe os t�picos para as Contrarraz�es de Recurso Inominado de CCR. Aqui n�o importa a ordem (a ordem
''    dos t�picos ser� a que est� na planilha de apoio).
''

    ' Agrupa os n�meros dos t�picos numa string, cercados por v�rgulas duplas.
    Dim strTopicos As String
    strTopicos = ",,"
    
    
    With form
        'Ilegitimidade ativa
        If .chbIlegitimidade.Value = True Then strTopicos = strTopicos & "10,,"
        
        'Prescri��o trienal
        If .chbPrescricao.Value = True Then strTopicos = strTopicos & "15,,"
        
        'Particularidades do caso
        If .chbInadimplenteEpoca.Value = True Then strTopicos = strTopicos & ",,"
        If .chbSemFatura.Value = True Then strTopicos = strTopicos & "60,,"
        If .chbBairroNaoAfetado.Value = True Then strTopicos = strTopicos & "20,,"
        If .chbReservatorio.Value = True Then strTopicos = strTopicos & "65,,"
        
        'Altera��o de consumo
        If .cmbAlteracaoConsumo.Value = "N�o houve altera��o relevante [gr�fico]" Then
            strTopicos = strTopicos & "40,,"
        ElseIf .cmbAlteracaoConsumo.Value = "N�o houve altera��o [print de contas juntadas pelo Autor]" Then
            strTopicos = strTopicos & "41,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Consumo aumentou no per�odo do acidente [gr�fico]" Then
            strTopicos = strTopicos & "42,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Abastecimento estava cortado por outro motivo" Then
            strTopicos = strTopicos & "45,,"
        End If
        
        'Senten�a
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
'' Recolhe os t�picos para as Contrarraz�es de Recurso Inominado de CCR. Aqui n�o importa a ordem (a ordem
''    dos t�picos ser� a que est� na planilha de apoio).
''

    ' Agrupa os n�meros dos t�picos numa string, cercados por v�rgulas duplas.
    Dim strTopicos As String
    strTopicos = ",,"
    
    
    With form
        'Ilegitimidade ativa
        If .chbIlegitimidade.Value = True Then strTopicos = strTopicos & "10,,"
        
        'Prescri��o trienal
        If .chbPrescricao.Value = True Then strTopicos = strTopicos & "15,,"
        
        'Particularidades do caso
        If .chbInadimplenteEpoca.Value = True Then strTopicos = strTopicos & ",,"
        If .chbSemFatura.Value = True Then strTopicos = strTopicos & "60,,"
        If .chbBairroNaoAfetado.Value = True Then strTopicos = strTopicos & "20,,"
        If .chbReservatorio.Value = True Then strTopicos = strTopicos & "65,,"
        
        'Altera��o de consumo
        If .cmbAlteracaoConsumo.Value = "N�o houve altera��o relevante [gr�fico]" Then
            strTopicos = strTopicos & "40,,"
        ElseIf .cmbAlteracaoConsumo.Value = "N�o houve altera��o [print de contas juntadas pelo Autor]" Then
            strTopicos = strTopicos & "41,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Consumo aumentou no per�odo do acidente [gr�fico]" Then
            strTopicos = strTopicos & "42,,"
        ElseIf .cmbAlteracaoConsumo.Value = "Abastecimento estava cortado por outro motivo" Then
            strTopicos = strTopicos & "45,,"
        End If
        
        'Senten�a
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
'' Monta a Contesta��o.
''
    
    MontarPeticaoComplexa ActiveSheet.Cells(ActiveCell.Row, 5).Formula, "Contesta��o", ActiveSheet.Cells(ActiveCell.Row, 11).Formula, Format(ActiveSheet.Cells(ActiveCell.Row, 9).Formula, "dd/mm/yyyy")
    
End Sub

Sub MontarRIApelacao(control As IRibbonControl)
''
'' Monta o RI.
''
    
    MontarPeticaoComplexa ActiveSheet.Cells(ActiveCell.Row, 5).Formula, "RI/Apela��o", ActiveSheet.Cells(ActiveCell.Row, 11).Formula, Format(ActiveSheet.Cells(ActiveCell.Row, 9).Formula, "dd/mm/yyyy")
    
End Sub

Sub MontarContrarrazoesRIApelacao(control As IRibbonControl)
''
'' Monta as Contrarraz�es do RI.
''
    
    MontarPeticaoComplexa ActiveSheet.Cells(ActiveCell.Row, 5).Formula, "Contrarraz�es de RI/Apela��o", ActiveSheet.Cells(ActiveCell.Row, 11).Formula, Format(ActiveSheet.Cells(ActiveCell.Row, 9).Formula, "dd/mm/yyyy")
    
End Sub

Sub DetectarProvMontarPet(control As IRibbonControl)
''
'' Verifica a provid�ncia e chama a fun��o correspondente.
''
    Dim strProvidencia As String
    Dim plan As Worksheet
    Dim contLinha As Long
    
    Set plan = ActiveSheet
    contLinha = ActiveCell.Row
    
    strProvidencia = plan.Cells(contLinha, 4).Text
    
    Select Case strProvidencia
    Case "Contestar", "Contestar - Remarca��o de audi�ncia"
        MontarContestacao control
        
    Case "Recorrer"
        MontarRIApelacao control
        
    Case "Contra-arrazoar recurso"
        MontarContrarrazoesRIApelacao control
    
    Case Else
        MsgBox "Sinto muito, " & DeterminarTratamento & "! O comando escolhido s� serve para Contestar, Recorrer ou Contra-arrazoar. " & _
                "Ser� que um dos outros bot�es n�o vos satisfaria?", vbInformation + vbOKOnly, "S�sifo em treinamento"
    
    End Select

    
    
End Sub

