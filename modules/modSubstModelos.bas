Attribute VB_Name = "modSubstModelos"
Sub EnviarPropostaAcordo(control As IRibbonControl)
    Dim strProcesso As String, strAdverso As String, strMatricula As String, strComarca As String, strAudiencia As String
    Dim rsAlcada As Currency
    Dim plan As Worksheet
    Dim contLinha As Long
    
    Set plan = ActiveSheet
    contLinha = ActiveCell.Row
    
    strProcesso = plan.Cells(contLinha, 1).Text
    strAdverso = plan.Cells(contLinha, 2).Text
    strMatricula = plan.Cells(contLinha, 3).Text
    strComarca = plan.Cells(contLinha, 12).Text
    strAudiencia = InputBox(DeterminarTratamento & ", qual a data e hora da audi�ncia?", "Informa��es sobre acordo", plan.Cells(contLinha, 7).Text)
    rsAlcada = CCur(InputBox(DeterminarTratamento & ", qual a al�ada para o acordo?", "Informa��es sobre acordo", "3.000,00"))
    
    GerarEmailSolicitacao "Acordo", strProcesso, strAdverso, strMatricula, strComarca, strAudiencia:=strAudiencia, rsAlcada:=rsAlcada
    
    plan.Cells(contLinha, 3).Interior.ColorIndex = 44
    
End Sub

Sub PedirLaudoIbametro(control As IRibbonControl)
    Dim strProcesso As String, strAdverso As String, strMatricula As String, strComarca As String
    Dim plan As Worksheet
    Dim dtPrazo As Date
    Dim contLinha As Long
    
    Set plan = ActiveSheet
    contLinha = ActiveCell.Row
    
    strProcesso = plan.Cells(contLinha, 1).Text
    strAdverso = plan.Cells(contLinha, 2).Text
    strMatricula = plan.Cells(contLinha, 3).Text
    strComarca = plan.Cells(contLinha, 12).Text
    dtPrazo = WorksheetFunction.WorkDay(CDate(plan.Cells(contLinha, 6).Text), -1)
    
    GerarEmailSolicitacao "Ibametro", strProcesso, strAdverso, strMatricula, strComarca, dtPrazo
    
    plan.Cells(contLinha, 3).Interior.ColorIndex = 44
    
End Sub

Sub PedirPagamentoCustas(control As IRibbonControl)
    Dim strProcesso As String, strAdverso As String, strMatricula As String, strComarca As String
    Dim plan As Worksheet
    Dim dtPrazo As Date
    Dim contLinha As Long
    
    Set plan = ActiveSheet
    contLinha = ActiveCell.Row
    
    strProcesso = plan.Cells(contLinha, 1).Text
    strAdverso = plan.Cells(contLinha, 2).Text
    strMatricula = plan.Cells(contLinha, 3).Text
    strComarca = plan.Cells(contLinha, 12).Text
    dtPrazo = WorksheetFunction.WorkDay(CDate(plan.Cells(contLinha, 6).Text), -1)
    
    GerarEmailSolicitacao "Preparo", strProcesso, strAdverso, strMatricula, strComarca, dtPrazo
End Sub

Sub PedirCumprimentoSentenca(control As IRibbonControl)
    Dim strProcesso As String, strAdverso As String, strMatricula As String
    Dim plan As Worksheet
    Dim contLinha As Long
    
    Set plan = ActiveSheet
    contLinha = ActiveCell.Row
    
    strProcesso = plan.Cells(contLinha, 1).Text
    strAdverso = plan.Cells(contLinha, 2).Text
    strMatricula = plan.Cells(contLinha, 3).Text
    
    GerarEmailSolicitacao "Cumprimento", strProcesso, strAdverso, strMatricula
    
End Sub

Sub PedirSubsidios(control As IRibbonControl)
    Dim strProcesso As String, strAdverso As String, strMatricula As String, strCausaPedir As String
    Dim dtPrazo As Date
    Dim plan As Worksheet
    Dim contLinha As Long
    
    Set plan = ActiveSheet
    contLinha = ActiveCell.Row
    
    strProcesso = plan.Cells(contLinha, 1).Text
    strAdverso = plan.Cells(contLinha, 2).Text
    strMatricula = plan.Cells(contLinha, 3).Text
    strCausaPedir = plan.Cells(contLinha, 5).Text
    dtPrazo = WorksheetFunction.WorkDay(CDate(plan.Cells(contLinha, 6).Text), -1)
    
    GerarEmailSolicitacao "Subsidios", strProcesso, strAdverso, strMatricula, , dtPrazo, strCausaPedir
    
    plan.Cells(contLinha, 3).Interior.ColorIndex = 44
    
End Sub

Sub JuntadaPagamento(control As IRibbonControl)
    Dim strJuizo As String, strProcesso As String, strAdverso As String ', strValor As String
    Dim bolCompensacao As Boolean
    Dim form As Variant
    Dim plan As Worksheet
    Dim contLinha As Long
    
    Set plan = ActiveSheet
    contLinha = ActiveCell.Row
        
    strProcesso = plan.Cells(contLinha, 1).Text
    strAdverso = plan.Cells(contLinha, 2).Text
    ' Pega o Ju�zo na reda��o do Espaider, depois pega o ju�zo na reda��o para ficar na planilha.
    strJuizo = ActiveSheet.Cells(ActiveCell.Row, 11).Formula ' Ju�zo na reda��o Espaider
    strJuizo = BuscaJuizo(strJuizo)

    'Ajusta e exibe o formul�rio
    Set form = New frmJuntPagamento
    form.Show
    If form.chbDeveGerar.Value = False Then Exit Sub
    
    If form.chbExisteDebito.Value = False Then
        'Juntada simples
        GerarPeticaoSimples strJuizo, strProcesso, strAdverso, "pagamento"
    Else
        'Juntada com compensa��o
        GerarPeticaoSimples strJuizo, strProcesso, strAdverso, "compensa��o", CCur(form.txtValCondenacao.Text), CCur(form.txtDebMatricula.Text)
    End If
    
End Sub

Sub JuntadaObrigFazer(control As IRibbonControl)
    Dim strJuizo As String, strProcesso As String, strAdverso As String ', strValor As String
    Dim plan As Worksheet
    Dim contLinha As Long
    
    Set plan = ActiveSheet
    contLinha = ActiveCell.Row
    
    strProcesso = plan.Cells(contLinha, 1).Text
    strAdverso = plan.Cells(contLinha, 2).Text
    ' Pega o Ju�zo na reda��o do Espaider, depois pega o ju�zo na reda��o para ficar na planilha.
    strJuizo = ActiveSheet.Cells(ActiveCell.Row, 11).Formula ' Ju�zo na reda��o Espaider
    strJuizo = BuscaJuizo(strJuizo)

    GerarPeticaoSimples strJuizo, strProcesso, strAdverso, "fazer"
End Sub

Sub JuntadaLiberacaoPenhora(Optional control As IRibbonControl)
    Dim strJuizo As String, strProcesso As String, strAdverso As String, strTextoObs As String, strNumeroConta As String
    Dim strExpressaoInicio As String, strExpressaoFinal As String
    Dim intInicio As Integer, intFim As Integer, intExtensaoExprInicio As Integer, intExtensaoExprFim As Integer
    Dim cValorPenhora As Currency
    Dim plan As Worksheet
    Dim contLinha As Long
    
    Set plan = ActiveSheet
    contLinha = ActiveCell.Row
    
    strProcesso = plan.Cells(contLinha, 1).Text
    strAdverso = plan.Cells(contLinha, 2).Text
    ' Pega o Ju�zo na reda��o do Espaider, depois pega o ju�zo na reda��o para ficar na planilha.
    strJuizo = ActiveSheet.Cells(ActiveCell.Row, 11).Formula ' Ju�zo na reda��o Espaider
    strJuizo = BuscaJuizo(strJuizo)
    
    strTextoObs = plan.Cells(contLinha, 7).Text
    strExpressaoInicio = "Conta judicial "
    strExpressaoFinal = ". Saldo de capital"
    intExtensaoExprInicio = Len(strExpressaoInicio)
    intInicio = InStr(1, strTextoObs, strExpressaoInicio)
    intFim = InStr(1, strTextoObs, strExpressaoFinal)
    If intInicio <> 0 And intFim <> 0 Then strNumeroConta = Mid(strTextoObs, intInicio + intExtensaoExprInicio, intFim - intInicio - intExtensaoExprInicio)
    
    strExpressaoInicio = ". Saldo de capital original: "
    strExpressaoFinal = ". Saldo atualizado"
    intExtensaoExprInicio = Len(strExpressaoInicio)
    intInicio = InStr(1, strTextoObs, strExpressaoInicio)
    intFim = InStr(1, strTextoObs, strExpressaoFinal)
    If intInicio <> 0 And intFim <> 0 Then cValorPenhora = CCur(Mid(strTextoObs, intInicio + intExtensaoExprInicio, intFim - intInicio - intExtensaoExprInicio))
    
    '". Saldo atualizado"

    GerarPeticaoSimples strJuizo, strProcesso, strAdverso, "liberarpenhora", cValorPenhora, , strNumeroConta
End Sub

Sub JuntadaRespostaRpv(Optional control As IRibbonControl)
    Dim strProcesso As String, idDeposito As String
    Dim cValor As Currency
    Dim expectativaPagamento As Date
    Dim plan As Worksheet
    Dim contLinha As Long
    
    Set plan = ActiveSheet
    contLinha = ActiveCell.Row
    
    strProcesso = plan.Cells(contLinha, 1).Text
    idDeposito = plan.Cells(contLinha, 4).Text
    cValor = CCur(plan.Cells(contLinha, 3).Text)
    expectativaPagamento = Format(plan.Cells(contLinha, 5).Text, "dd/mm/yyyy")
    
    GerarRespostaRpv strProcesso, idDeposito, cValor, expectativaPagamento
End Sub

Sub ManifestacaoLiminar(control As IRibbonControl)
    Dim strJuizo As String, strProcesso As String, strAdverso As String
    Dim plan As Worksheet
    Dim contLinha As Long
    
    Set plan = ActiveSheet
    contLinha = ActiveCell.Row
    
    strProcesso = plan.Cells(contLinha, 1).Text
    strAdverso = plan.Cells(contLinha, 2).Text
    ' Pega o Ju�zo na reda��o do Espaider, depois pega o ju�zo na reda��o para ficar na planilha.
    strJuizo = ActiveSheet.Cells(ActiveCell.Row, 11).Formula ' Ju�zo na reda��o Espaider
    strJuizo = BuscaJuizo(strJuizo)

    GerarPeticaoSimples strJuizo, strProcesso, strAdverso, "liminar"
End Sub

Sub JuntadaPreparo(control As IRibbonControl)
    Dim strJuizo As String, strProcesso As String, strAdverso As String
    Dim plan As Worksheet
    Dim contLinha As Long
    
    Set plan = ActiveSheet
    contLinha = ActiveCell.Row
    
    strProcesso = plan.Cells(contLinha, 1).Text
    strAdverso = plan.Cells(contLinha, 2).Text
    ' Pega o Ju�zo na reda��o do Espaider, depois pega o ju�zo na reda��o para ficar na planilha.
    strJuizo = ActiveSheet.Cells(ActiveCell.Row, 11).Formula ' Ju�zo na reda��o Espaider
    strJuizo = BuscaJuizo(strJuizo)

    GerarPeticaoSimples strJuizo, strProcesso, strAdverso, "preparo"
End Sub

Sub RequerimentoAlvara(control As IRibbonControl)
    Dim strJuizo As String, strProcesso As String, strAdverso As String
    Dim plan As Worksheet
    Dim contLinha As Long
    
    Set plan = ActiveSheet
    contLinha = ActiveCell.Row
    
    strProcesso = plan.Cells(contLinha, 1).Text
    strAdverso = plan.Cells(contLinha, 2).Text
    ' Pega o Ju�zo na reda��o do Espaider, depois pega o ju�zo na reda��o para ficar na planilha.
    strJuizo = ActiveSheet.Cells(ActiveCell.Row, 11).Formula ' Ju�zo na reda��o Espaider
    strJuizo = BuscaJuizo(strJuizo)

    GerarPeticaoSimples strJuizo, strProcesso, strAdverso, "alvar�"
End Sub

Sub RequerimentoExecucao(control As IRibbonControl)
    Dim strJuizo As String, strProcesso As String, strAdverso As String
    Dim plan As Worksheet
    Dim contLinha As Long
    
    Set plan = ActiveSheet
    contLinha = ActiveCell.Row
    
    strProcesso = plan.Cells(contLinha, 1).Text
    strAdverso = plan.Cells(contLinha, 2).Text
    ' Pega o Ju�zo na reda��o do Espaider, depois pega o ju�zo na reda��o para ficar na planilha.
    strJuizo = ActiveSheet.Cells(ActiveCell.Row, 11).Formula ' Ju�zo na reda��o Espaider
    strJuizo = BuscaJuizo(strJuizo)

    GerarPeticaoSimples strJuizo, strProcesso, strAdverso, "execu��o"
End Sub

Sub RequerimentoCertidao(control As IRibbonControl)
    Dim strJuizo As String, strProcesso As String, strAdverso As String
    Dim plan As Worksheet
    Dim contLinha As Long
    
    Set plan = ActiveSheet
    contLinha = ActiveCell.Row
    
    strProcesso = plan.Cells(contLinha, 1).Text
    strAdverso = plan.Cells(contLinha, 2).Text
    ' Pega o Ju�zo na reda��o do Espaider, depois pega o ju�zo na reda��o para ficar na planilha.
    strJuizo = ActiveSheet.Cells(ActiveCell.Row, 11).Formula ' Ju�zo na reda��o Espaider
    strJuizo = BuscaJuizo(strJuizo)

    GerarPeticaoSimples strJuizo, strProcesso, strAdverso, "certid�o de daje"
End Sub

Sub GerarEmailSolicitacao(strModalidade As String, strProcesso As String, strAdverso As String, strMatricula As String, _
    Optional strComarca As String = "", Optional dtPrazo As Date = Empty, Optional strCausaPedir As String = "", _
    Optional strAudiencia As String = "", Optional rsAlcada As Currency = CCur(0))

    Dim appword As Object
    Dim wdDocPeticao As Word.Document
    Dim strModelo As String, strCont As String
    
    ' Selecionar modelo
    Select Case strModalidade
        Case "Acordo"
            strModelo = "Proposta-Al�ada-Acordo.dotx"
        Case "Ibametro"
            strModelo = "Pedido-Laudo-Ibametro.dotx"
        Case "Preparo"
            strModelo = "Pedido-Pagamento-Custas.dotx"
        Case "Cumprimento"
            strModelo = "Pedido-Cumprimento-Sentenca.dotx"
            
            'Ajusta e exibe o formul�rio
            Set form = New frmSolicitaFazer
            form.Show
            If form.chbDeveGerar.Value = False Then Exit Sub
    
            'Colhe os t�picos
            If form.chbRefat.Value = True Then strCont = strCont & "- Refaturar meses: " & form.txtMesesRef.Text & " para " & form.txtValorRefat.Text & Chr(13)
            If form.chbCancelarCobranca.Value = True Then strCont = strCont & "- Cancelar cobran�a de " & form.cmbCobrancaACancelar.Text & " nos meses: " & form.txtMesesCancelar.Text & Chr(13)
            If form.chbQuitar.Value = True Then strCont = strCont & "- Quitar faturas dos meses: " & form.txtMesesQuitar.Text & " pelos dep�sitos judiciais" & Chr(13)
            If form.chbExcluirSPC.Value = True Then strCont = strCont & "- Excluir Autor do SPC" & Chr(13)
            If form.chbDesvincularContrato.Value = True Then strCont = strCont & "- Desvincular contrato do nome do Autor" & Chr(13)
            If form.chbReligar.Value = True Then strCont = strCont & "- Religar liga��o" & Chr(13)
            If form.chbDesligar.Value = True Then strCont = strCont & "- Desligar liga��o" & Chr(13)
            If form.chbDesmembrar.Value = True Then strCont = strCont & "- Desmembrar liga��o" & Chr(13)
            If form.chbSubsHidrometro.Value = True Then strCont = strCont & "- Substituir hidr�metro (ou instalar, se n�o houver)" & Chr(13)
            If form.chbRealizarLigacao.Value = True Then strCont = strCont & "- Realizar liga��o" & Chr(13)
            If Trim(form.txtOutros.Text) <> "" Then strCont = strCont & "- " & form.txtOutros.Text & Chr(13)
            If Trim(form.txtObsGeral.Text) <> "" Then strCont = strCont & Chr(13) & "Obs.: " & form.txtObsGeral
            Do While Right(strCont, 1) = Chr(13)
                strCont = Left(strCont, Len(strCont) - 1) 'Apaga enter no final
            Loop
            
        Case "Subsidios"
            Select Case strCausaPedir
                Case "Negativa��o no SPC"
                    strModelo = "Pedido-Solicita-Subsidios-Negativacao.dotx"
                Case "Corte no fornecimento"
                    strModelo = "Pedido-Solicita-Subsidios-Corte.dotx"
                Case Else
                    strModelo = "Pedido-Solicita-Subsidios.dotx"
            End Select
    End Select
    
    ' Criar o documento a partir do modelo
    
    Set appword = CreateObject("Word.Application")
    appword.Visible = True
    Set wdDocPeticao = appword.Documents.Add(BuscarCaminhoPrograma & "modelos-automaticos\" & strModelo)
    appword.Activate

    ' Realizar as substitui��es no documento
    
    With appword.Selection.Find
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
    
        If strModalidade = "Cumprimento" Then
            .Text = "<matr�cula>"
            .Replacement.Text = strMatricula
            .Execute Replace:=wdReplaceAll
            
            .Text = "<obriga��es>"
            .Replacement.Text = ""
            .Execute Replace:=wdReplaceOne
            appword.Selection.Range.Text = strCont 'Tem que ser colado desta forma, pois se substituir desformata os par�grafos.

            
        Else
            .Text = "<email>"
            .Replacement.Text = PegarEmail(strModalidade)
            .Execute Replace:=wdReplaceAll
    
            .Text = "<processo>"
            .Replacement.Text = strProcesso
            .Execute Replace:=wdReplaceAll
    
            .Text = "<adverso>"
            .Replacement.Text = strAdverso
            .Execute Replace:=wdReplaceAll
    
            .Text = "<matr�cula>"
            .Replacement.Text = strMatricula
            .Execute Replace:=wdReplaceAll
            
            If strComarca <> "" Then
                .Text = "<comarca>"
                .Replacement.Text = strComarca
                .Execute Replace:=wdReplaceAll
            End If
    
            If Not IsEmpty(dtPrazo) Then
                .Text = "<prazo>"
                .Replacement.Text = Format(dtPrazo, "dd/mm/yy")
                .Execute Replace:=wdReplaceAll
            End If
            
            If strAudiencia <> "" Then
                .Text = "<audi�ncia>"
                .Replacement.Text = strAudiencia
                .Execute Replace:=wdReplaceAll
            End If
    
            If rsAlcada <> 0 Then
                .Text = "<al�ada>"
                .Replacement.Text = rsAlcada
                .Execute Replace:=wdReplaceAll
            End If
    
        End If
    End With
    
    ' Copiar para a �rea de transfer�ncia
    
    wdDocPeticao.Content.Copy
    
    ' Fechar o documento
    
    wdDocPeticao.Close savechanges:=wdDoNotSaveChanges
    
    If appword.Documents.Count = 0 Then appword.Quit
    
End Sub

Function PegarEmail(strModalidade As String)
''
'' Retorna o e-mail � direita do nome da modalidade na planilha "cfConfigura��es".
''
    Dim intPrimeiraLinha As Integer, intUltimaLinha As Integer
    Dim strBusca As String
    Dim plan As Worksheet
    Set plan = ThisWorkbook.Worksheets("cfConfigura��es")
    
    'Definir a express�o de busca
    Select Case strModalidade
        Case "Acordo"
            strBusca = "Proposta de acordo"
        Case "Ibametro"
            strBusca = "Solicita��o de laudos Ibametro"
        Case "Preparo"
            strBusca = "Solicita��o de pagamento de custas"
        Case "Pautista"
            strBusca = "Solicita��o de advogado pautista"
    End Select
        
    'Ir � c�lula � direita. Retornar e-mail
    PegarEmail = plan.Cells().Find(what:=strBusca, lookat:=xlWhole, searchorder:=xlByColumns, MatchCase:=False).Offset(0, 1).Text
    
End Function

Sub GerarRespostaRpv(numProcesso As String, idDeposito As String, valor As Currency, expectativaPagamento As Date)
    Dim appword As Object
    Dim wdDocPeticao As Word.Document
    Dim strDocOrigem As String
    Dim strCont As Variant
    Dim cCont As Currency
    
    ' Criar o documento a partir do modelo
    
    Set appword = CreateObject("Word.Application")
    appword.Visible = True
    Set wdDocPeticao = appword.Documents.Add(BuscarCaminhoPrograma & "modelos-automaticos\PPJCM Modelo.dotx")
    strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99920 Resposta RPV.docx"
    If Not strDocOrigem = "" Then InserirArquivo strDocOrigem, wdDocPeticao
    
    ' Realizar as substitui��es no documento
    
    With wdDocPeticao.Content.Find
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        
        .Text = "<numeroProcesso>"
        .Replacement.Text = numProcesso
        .Execute Replace:=wdReplaceAll
        
        .Text = "<valor>"
        .Replacement.Text = Format(valor, "#,##0.00")
        .Execute Replace:=wdReplaceAll
        
        .Text = "<idDeposito>"
        .Replacement.Text = idDeposito
        .Execute Replace:=wdReplaceAll
        
        .Text = "<expectativaPagamento>"
        .Replacement.Text = IIf(expectativaPagamento = 0, "", expectativaPagamento)
        .Execute Replace:=wdReplaceAll
    End With
    
    wdDocPeticao.Paragraphs.Last.Range.Delete
    appword.Activate
    
    ' Salvar o documento
    strCont = "Resposta RPV"
    
    ' Salvar
    If bolSsfPrazosBotaoPdfPressionado Then 'Gerar como PDF
        wdDocPeticao.ExportAsFixedFormat OutputFilename:=BuscarCaminhoPrograma & "01 " & strCont & " - " & numProcesso & " " & idDeposito & ".pdf", _
            ExportFormat:=wdExportFormatPDF, OptimizeFor:=wdExportOptimizeForOnScreen, CreateBookmarks:=wdExportCreateHeadingBookmarks, BitmapMissingFonts:=False
        'MsgBox DeterminarTratamento & ", o PDF foi gerado com sucesso.", vbInformation + vbOKOnly, "S�sifo - Documento gerado"
        wdDocPeticao.Close wdDoNotSaveChanges
        If appword.Documents.Count = 0 Then appword.Quit
        
    Else ' Gerar como Word
        wdDocPeticao.SaveAs BuscarCaminhoPrograma & Format(Date, "yyyy.mm.dd") & " " & strCont & " - " & numProcesso & ".docx"
        wdDocPeticao.GoTo wdGoToPage, wdGoToFirst
        MsgBox DeterminarTratamento & ", o documento Word foi gerado com sucesso.", vbInformation + vbOKOnly, "S�sifo - Documento gerado"
    End If
End Sub

Sub GerarPeticaoSimples(strJuizo As String, strProcesso As String, strAdverso As String, strPeticao As String, Optional curCondenacao As Currency = 0, _
    Optional curDebitoMatricula As Currency = 0, Optional strContaJudicial As String = "")

    Dim appword As Object
    Dim wdDocPeticao As Word.Document
    Dim strDocOrigem As String
    Dim strCont As Variant
    Dim cCont As Currency
    
    ' Criar o documento a partir do modelo
    
    Set appword = CreateObject("Word.Application")
    appword.Visible = True
    Set wdDocPeticao = appword.Documents.Add(BuscarCaminhoPrograma & "modelos-automaticos\PPJCM Modelo.dotx")
    
    ' Colocar o cabe�alho
    Select Case strPeticao
    Case "pagamento"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99901 Juntada cumprimento Cabe�alho.docx"
    Case "compensa��o"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99904 Juntada pagamento e compensa��o Cabe�alho.docx"
    Case "fazer"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99901 Juntada cumprimento Cabe�alho.docx"
    Case "liminar"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99915 Manifesta��o liminar Cabe�alho.docx"
    Case "alvar�"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99908 Requerimento alvar� Cabe�alho.docx"
    Case "liberarpenhora"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99917 Requerimento libera��o de penhora Cabe�alho.docx"
    Case "preparo"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99906 Juntada preparo Cabe�alho.docx"
    Case "execu��o"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99910 Requerimento execu��o Cabe�alho.docx"
    Case "certid�o de daje"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99912 Requerimento certid�o DAJEs.docx"
    End Select
    
    InserirArquivo strDocOrigem, wdDocPeticao
 
    ' Colocar o pedido de habilita�ao de advogado
    Select Case strPeticao
    Case "pagamento", "fazer", "preparo", "alvar�"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99913 Habilita��o advogado Pequena.docx"
    Case "compensa��o", "execu��o", "liminar", "liberarpenhora"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99913 Habilita��o advogado Grande.docx"
    Case "certid�o de daje"
        strDocOrigem = ""
    End Select
    
    If Not strDocOrigem = "" Then InserirArquivo strDocOrigem, wdDocPeticao
    
    ' Colocar a conclus�o
    Select Case strPeticao
    Case "pagamento"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99903 Juntada pagamento Conclus�o.docx"
    Case "compensa��o"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99905 Juntada pagamento e compensa��o Corpo.docx"
    Case "fazer"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99902 Juntada cumprimento Conclus�o.docx"
    Case "liminar"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99916 Manifesta��o liminar Corpo.docx"
    Case "preparo"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99907 Juntada preparo Conclus�o.docx"
    Case "alvar�"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99909 Requerimento alvar� Conclus�o.docx"
    Case "liberarpenhora"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99918 Requerimento libera��o de penhora Corpo.docx"
    Case "execu��o"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99911 Requerimento execu��o Corpo.docx"
    Case "certid�o de daje"
        strDocOrigem = ""
    End Select
    
    If Not strDocOrigem = "" Then InserirArquivo strDocOrigem, wdDocPeticao
    
    ' Colocar o rodap�
    
    strDocOrigem = BuscarCaminhoPrograma & "Frankenstein\00099 Pedidos - Rodap�.docx"
    InserirArquivo strDocOrigem, wdDocPeticao
    
    ' Realizar as substitui��es no documento
    
    With wdDocPeticao.Content.Find
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        
        .Text = "<ju�zo>"
        .Replacement.Text = UCase(strJuizo)
        .Execute Replace:=wdReplaceAll
        'Se o ju�zo n�o foi encontrado no S�sifo, avisa para colocar manualmente.
        If strJuizo = "" Then MsgBox "Ju�zo n�o encontrado na base de dados no S�sifo. Lembre-se de acrescent�-lo manualmente no endere�amento da peti��o.", _
            vbCritical + vbOKOnly, "Alerta - ju�zo n�o encontrado"

        
        .Text = "<processo>"
        .Replacement.Text = strProcesso
        .Execute Replace:=wdReplaceAll
        
        .Text = "<adverso>"
        .Replacement.Text = strAdverso
        .Execute Replace:=wdReplaceAll
        
        If strPeticao = "alvar�" Then
            .Text = "<eventos>"
            strCont = Application.InputBox(DeterminarTratamento & ", quais eventos t�m dep�sitos judiciais para solicitar alvar� para a Embasa?", "Informa��es sobre alvar�", "11, 26 e 35", Type:=2)
            If strCont = False Then strCont = ""
            .Replacement.Text = strCont
            .Execute Replace:=wdReplaceAll
        End If
        
        If strPeticao = "compensa��o" Then
            .Text = "<d�bito-embasa>"
            .Replacement.Text = Format(curCondenacao, "#,##0.00")
            .Execute Replace:=wdReplaceAll
        
            .Text = "<d�bito-autor>"
            .Replacement.Text = Format(curDebitoMatricula, "#,##0.00")
            .Execute Replace:=wdReplaceAll
        
            .Text = "<diferen�a-para-autor>"
            .Replacement.Text = Format(curCondenacao - curDebitoMatricula, "#,##0.00")
            .Execute Replace:=wdReplaceAll
        End If
        
        If strPeticao = "liberarpenhora" Then
            .Text = "<valor-penhora>"
            .Replacement.Text = Format(curCondenacao, "#,##0.00")
            .Execute Replace:=wdReplaceAll
            
            .Text = "<conta-judicial>"
            .Replacement.Text = strContaJudicial
            .Execute Replace:=wdReplaceAll
        End If
        
        If strPeticao = "execu��o" Then
            .Text = "<CPF-Autor>"
            strCont = Application.InputBox(DeterminarTratamento & ", qual o CPF/CNPJ do executado?", "Informa��es sobre execu��o", "", Type:=2)
            If strCont = False Then strCont = ""
            .Replacement.Text = strCont
            .Execute Replace:=wdReplaceAll
            
            .Text = "<natureza-condena��o>"
            strCont = Application.InputBox(DeterminarTratamento & ", a que se refere o valor exeq�endo?", "Informa��es sobre execu��o", "honor�rios sucumbenciais", Type:=2)
            If strCont = False Then strCont = ""
            .Replacement.Text = strCont
            .Execute Replace:=wdReplaceAll
            
            .Text = "<valor-condena��o>"
            cCont = Application.InputBox(DeterminarTratamento & ", qual o valor da condena��o (antes da atualiza��o)?", "Informa��es sobre execu��o", "", Type:=2)
            If cCont = False Then cCont = 0
            .Replacement.Text = Format(cCont, "#.##0,00")
            .Execute Replace:=wdReplaceAll
            
            .Text = "<termo-inicial-atualiza��o>"
            strCont = Application.InputBox(DeterminarTratamento & ", qual foi o termo inicial da atualiza��o?", "Informa��es sobre execu��o", "10/10/2019", Type:=2)
            If strCont = False Then strCont = ""
            .Replacement.Text = strCont
            .Execute Replace:=wdReplaceAll
            
            .Text = "<termo-inicial-juros>"
            strCont = Application.InputBox(DeterminarTratamento & ", qual foi o termo inicial dos juros de mora?", "Informa��es sobre execu��o", "10/10/2019", Type:=2)
            If strCont = False Then strCont = ""
            .Replacement.Text = strCont
            .Execute Replace:=wdReplaceAll
            
            .Text = "<valor-atualizado>"
            cCont = CCur(Application.InputBox(DeterminarTratamento & ", qual o valor da condena��o atualizado?", "Informa��es sobre execu��o", "", Type:=2))
            If cCont = False Then cCont = 0
            .Replacement.Text = Format(cCont, "#.##0,00")
            .Execute Replace:=wdReplaceAll
            
            .Text = "<valor-multa>"
            .Replacement.Text = Format(cCont * 0.1, "#.##0,00")
            .Execute Replace:=wdReplaceAll
            
            .Text = "<valor-com-multa-e-honor�rios>"
            .Replacement.Text = Format(cCont * 1.2, "#.##0,00")
            .Execute Replace:=wdReplaceAll
            
        End If
        
        If strPeticao = "certid�o de daje" Then
            .Text = "<data-pagamento>"
            strCont = Application.InputBox(DeterminarTratamento & ", quando foi feito o pagamento dos DAJEs?", "Informa��es sobre DAJEs", Format(Date - 13, "dd/mm/yyyy"), Type:=2)
            If strCont = False Then strCont = ""
            .Replacement.Text = strCont
            .Execute Replace:=wdReplaceAll
        End If
        
        .Text = "<data>"
        .Replacement.Text = Format(Date, "dd \de mmmm \de yyyy")
        .Execute Replace:=wdReplaceAll

    End With
    
    wdDocPeticao.Paragraphs.Last.Range.Delete
    appword.Activate
    
    ' Salvar o documento
    Select Case strPeticao
    Case "pagamento"
        strCont = "Juntada pagamento"
    Case "compensa��o"
        strCont = "Pagamento e compensacao"
    Case "fazer"
        strCont = "Juntada fazer"
    Case "liminar"
        strCont = "Manifestacao liminar"
    Case "preparo"
        strCont = "Juntada preparo"
    Case "alvar�"
        strCont = "Requerimento alvara"
    Case "liberarpenhora"
        strCont = "Liberacao de valor"
    Case "execu��o"
        strCont = "Requerimento de execucao"
    Case "certid�o de daje"
        strCont = "Requerimento certidao"
    End Select
    
    ' Salvar
    
    
    If bolSsfPrazosBotaoPdfPressionado Then 'Gerar como PDF
        wdDocPeticao.ExportAsFixedFormat OutputFilename:=BuscarCaminhoPrograma & "01 " & strCont & " - " & SeparaPrimeirosNomes(strAdverso, 2) & ".pdf", _
            ExportFormat:=wdExportFormatPDF, OptimizeFor:=wdExportOptimizeForOnScreen, CreateBookmarks:=wdExportCreateHeadingBookmarks, BitmapMissingFonts:=False
        'MsgBox DeterminarTratamento & ", o PDF foi gerado com sucesso.", vbInformation + vbOKOnly, "S�sifo - Documento gerado"
        wdDocPeticao.Close wdDoNotSaveChanges
        If appword.Documents.Count = 0 Then appword.Quit
        
    Else ' Gerar como Word
        wdDocPeticao.SaveAs BuscarCaminhoPrograma & Format(Date, "yyyy.mm.dd") & " " & strCont & " - " & SeparaPrimeirosNomes(strAdverso, 2) & ".docx"
        wdDocPeticao.GoTo wdGoToPage, wdGoToFirst
        MsgBox DeterminarTratamento & ", o documento Word foi gerado com sucesso.", vbInformation + vbOKOnly, "S�sifo - Documento gerado"
    End If
    
End Sub


