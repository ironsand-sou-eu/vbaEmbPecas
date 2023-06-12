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
    strAudiencia = InputBox(DeterminarTratamento & ", qual a data e hora da audiência?", "Informações sobre acordo", plan.Cells(contLinha, 7).Text)
    rsAlcada = CCur(InputBox(DeterminarTratamento & ", qual a alçada para o acordo?", "Informações sobre acordo", "3.000,00"))
    
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
    ' Pega o Juízo na redação do Espaider, depois pega o juízo na redação para ficar na planilha.
    strJuizo = ActiveSheet.Cells(ActiveCell.Row, 11).Formula ' Juízo na redação Espaider
    strJuizo = BuscaJuizo(strJuizo)

    'Ajusta e exibe o formulário
    Set form = New frmJuntPagamento
    form.Show
    If form.chbDeveGerar.Value = False Then Exit Sub
    
    If form.chbExisteDebito.Value = False Then
        'Juntada simples
        GerarPeticaoSimples strJuizo, strProcesso, strAdverso, "pagamento"
    Else
        'Juntada com compensação
        GerarPeticaoSimples strJuizo, strProcesso, strAdverso, "compensação", CCur(form.txtValCondenacao.Text), CCur(form.txtDebMatricula.Text)
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
    ' Pega o Juízo na redação do Espaider, depois pega o juízo na redação para ficar na planilha.
    strJuizo = ActiveSheet.Cells(ActiveCell.Row, 11).Formula ' Juízo na redação Espaider
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
    ' Pega o Juízo na redação do Espaider, depois pega o juízo na redação para ficar na planilha.
    strJuizo = ActiveSheet.Cells(ActiveCell.Row, 11).Formula ' Juízo na redação Espaider
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
    ' Pega o Juízo na redação do Espaider, depois pega o juízo na redação para ficar na planilha.
    strJuizo = ActiveSheet.Cells(ActiveCell.Row, 11).Formula ' Juízo na redação Espaider
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
    ' Pega o Juízo na redação do Espaider, depois pega o juízo na redação para ficar na planilha.
    strJuizo = ActiveSheet.Cells(ActiveCell.Row, 11).Formula ' Juízo na redação Espaider
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
    ' Pega o Juízo na redação do Espaider, depois pega o juízo na redação para ficar na planilha.
    strJuizo = ActiveSheet.Cells(ActiveCell.Row, 11).Formula ' Juízo na redação Espaider
    strJuizo = BuscaJuizo(strJuizo)

    GerarPeticaoSimples strJuizo, strProcesso, strAdverso, "alvará"
End Sub

Sub RequerimentoExecucao(control As IRibbonControl)
    Dim strJuizo As String, strProcesso As String, strAdverso As String
    Dim plan As Worksheet
    Dim contLinha As Long
    
    Set plan = ActiveSheet
    contLinha = ActiveCell.Row
    
    strProcesso = plan.Cells(contLinha, 1).Text
    strAdverso = plan.Cells(contLinha, 2).Text
    ' Pega o Juízo na redação do Espaider, depois pega o juízo na redação para ficar na planilha.
    strJuizo = ActiveSheet.Cells(ActiveCell.Row, 11).Formula ' Juízo na redação Espaider
    strJuizo = BuscaJuizo(strJuizo)

    GerarPeticaoSimples strJuizo, strProcesso, strAdverso, "execução"
End Sub

Sub RequerimentoCertidao(control As IRibbonControl)
    Dim strJuizo As String, strProcesso As String, strAdverso As String
    Dim plan As Worksheet
    Dim contLinha As Long
    
    Set plan = ActiveSheet
    contLinha = ActiveCell.Row
    
    strProcesso = plan.Cells(contLinha, 1).Text
    strAdverso = plan.Cells(contLinha, 2).Text
    ' Pega o Juízo na redação do Espaider, depois pega o juízo na redação para ficar na planilha.
    strJuizo = ActiveSheet.Cells(ActiveCell.Row, 11).Formula ' Juízo na redação Espaider
    strJuizo = BuscaJuizo(strJuizo)

    GerarPeticaoSimples strJuizo, strProcesso, strAdverso, "certidão de daje"
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
            strModelo = "Proposta-Alçada-Acordo.dotx"
        Case "Ibametro"
            strModelo = "Pedido-Laudo-Ibametro.dotx"
        Case "Preparo"
            strModelo = "Pedido-Pagamento-Custas.dotx"
        Case "Cumprimento"
            strModelo = "Pedido-Cumprimento-Sentenca.dotx"
            
            'Ajusta e exibe o formulário
            Set form = New frmSolicitaFazer
            form.Show
            If form.chbDeveGerar.Value = False Then Exit Sub
    
            'Colhe os tópicos
            If form.chbRefat.Value = True Then strCont = strCont & "- Refaturar meses: " & form.txtMesesRef.Text & " para " & form.txtValorRefat.Text & Chr(13)
            If form.chbCancelarCobranca.Value = True Then strCont = strCont & "- Cancelar cobrança de " & form.cmbCobrancaACancelar.Text & " nos meses: " & form.txtMesesCancelar.Text & Chr(13)
            If form.chbQuitar.Value = True Then strCont = strCont & "- Quitar faturas dos meses: " & form.txtMesesQuitar.Text & " pelos depósitos judiciais" & Chr(13)
            If form.chbExcluirSPC.Value = True Then strCont = strCont & "- Excluir Autor do SPC" & Chr(13)
            If form.chbDesvincularContrato.Value = True Then strCont = strCont & "- Desvincular contrato do nome do Autor" & Chr(13)
            If form.chbReligar.Value = True Then strCont = strCont & "- Religar ligação" & Chr(13)
            If form.chbDesligar.Value = True Then strCont = strCont & "- Desligar ligação" & Chr(13)
            If form.chbDesmembrar.Value = True Then strCont = strCont & "- Desmembrar ligação" & Chr(13)
            If form.chbSubsHidrometro.Value = True Then strCont = strCont & "- Substituir hidrômetro (ou instalar, se não houver)" & Chr(13)
            If form.chbRealizarLigacao.Value = True Then strCont = strCont & "- Realizar ligação" & Chr(13)
            If Trim(form.txtOutros.Text) <> "" Then strCont = strCont & "- " & form.txtOutros.Text & Chr(13)
            If Trim(form.txtObsGeral.Text) <> "" Then strCont = strCont & Chr(13) & "Obs.: " & form.txtObsGeral
            Do While Right(strCont, 1) = Chr(13)
                strCont = Left(strCont, Len(strCont) - 1) 'Apaga enter no final
            Loop
            
        Case "Subsidios"
            Select Case strCausaPedir
                Case "Negativação no SPC"
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

    ' Realizar as substituições no documento
    
    With appword.Selection.Find
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
    
        If strModalidade = "Cumprimento" Then
            .Text = "<matrícula>"
            .Replacement.Text = strMatricula
            .Execute Replace:=wdReplaceAll
            
            .Text = "<obrigações>"
            .Replacement.Text = ""
            .Execute Replace:=wdReplaceOne
            appword.Selection.Range.Text = strCont 'Tem que ser colado desta forma, pois se substituir desformata os parágrafos.

            
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
    
            .Text = "<matrícula>"
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
                .Text = "<audiência>"
                .Replacement.Text = strAudiencia
                .Execute Replace:=wdReplaceAll
            End If
    
            If rsAlcada <> 0 Then
                .Text = "<alçada>"
                .Replacement.Text = rsAlcada
                .Execute Replace:=wdReplaceAll
            End If
    
        End If
    End With
    
    ' Copiar para a área de transferência
    
    wdDocPeticao.Content.Copy
    
    ' Fechar o documento
    
    wdDocPeticao.Close savechanges:=wdDoNotSaveChanges
    
    If appword.Documents.Count = 0 Then appword.Quit
    
End Sub

Function PegarEmail(strModalidade As String)
''
'' Retorna o e-mail à direita do nome da modalidade na planilha "cfConfigurações".
''
    Dim intPrimeiraLinha As Integer, intUltimaLinha As Integer
    Dim strBusca As String
    Dim plan As Worksheet
    Set plan = ThisWorkbook.Worksheets("cfConfigurações")
    
    'Definir a expressão de busca
    Select Case strModalidade
        Case "Acordo"
            strBusca = "Proposta de acordo"
        Case "Ibametro"
            strBusca = "Solicitação de laudos Ibametro"
        Case "Preparo"
            strBusca = "Solicitação de pagamento de custas"
        Case "Pautista"
            strBusca = "Solicitação de advogado pautista"
    End Select
        
    'Ir à célula à direita. Retornar e-mail
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
    
    ' Realizar as substituições no documento
    
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
        'MsgBox DeterminarTratamento & ", o PDF foi gerado com sucesso.", vbInformation + vbOKOnly, "Sísifo - Documento gerado"
        wdDocPeticao.Close wdDoNotSaveChanges
        If appword.Documents.Count = 0 Then appword.Quit
        
    Else ' Gerar como Word
        wdDocPeticao.SaveAs BuscarCaminhoPrograma & Format(Date, "yyyy.mm.dd") & " " & strCont & " - " & numProcesso & ".docx"
        wdDocPeticao.GoTo wdGoToPage, wdGoToFirst
        MsgBox DeterminarTratamento & ", o documento Word foi gerado com sucesso.", vbInformation + vbOKOnly, "Sísifo - Documento gerado"
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
    
    ' Colocar o cabeçalho
    Select Case strPeticao
    Case "pagamento"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99901 Juntada cumprimento Cabeçalho.docx"
    Case "compensação"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99904 Juntada pagamento e compensação Cabeçalho.docx"
    Case "fazer"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99901 Juntada cumprimento Cabeçalho.docx"
    Case "liminar"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99915 Manifestação liminar Cabeçalho.docx"
    Case "alvará"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99908 Requerimento alvará Cabeçalho.docx"
    Case "liberarpenhora"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99917 Requerimento liberação de penhora Cabeçalho.docx"
    Case "preparo"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99906 Juntada preparo Cabeçalho.docx"
    Case "execução"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99910 Requerimento execução Cabeçalho.docx"
    Case "certidão de daje"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99912 Requerimento certidão DAJEs.docx"
    End Select
    
    InserirArquivo strDocOrigem, wdDocPeticao
 
    ' Colocar o pedido de habilitaçao de advogado
    Select Case strPeticao
    Case "pagamento", "fazer", "preparo", "alvará"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99913 Habilitação advogado Pequena.docx"
    Case "compensação", "execução", "liminar", "liberarpenhora"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99913 Habilitação advogado Grande.docx"
    Case "certidão de daje"
        strDocOrigem = ""
    End Select
    
    If Not strDocOrigem = "" Then InserirArquivo strDocOrigem, wdDocPeticao
    
    ' Colocar a conclusão
    Select Case strPeticao
    Case "pagamento"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99903 Juntada pagamento Conclusão.docx"
    Case "compensação"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99905 Juntada pagamento e compensação Corpo.docx"
    Case "fazer"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99902 Juntada cumprimento Conclusão.docx"
    Case "liminar"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99916 Manifestação liminar Corpo.docx"
    Case "preparo"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99907 Juntada preparo Conclusão.docx"
    Case "alvará"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99909 Requerimento alvará Conclusão.docx"
    Case "liberarpenhora"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99918 Requerimento liberação de penhora Corpo.docx"
    Case "execução"
        strDocOrigem = BuscarCaminhoPrograma & "modelos-automaticos\99911 Requerimento execução Corpo.docx"
    Case "certidão de daje"
        strDocOrigem = ""
    End Select
    
    If Not strDocOrigem = "" Then InserirArquivo strDocOrigem, wdDocPeticao
    
    ' Colocar o rodapé
    
    strDocOrigem = BuscarCaminhoPrograma & "Frankenstein\00099 Pedidos - Rodapé.docx"
    InserirArquivo strDocOrigem, wdDocPeticao
    
    ' Realizar as substituições no documento
    
    With wdDocPeticao.Content.Find
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        
        .Text = "<juízo>"
        .Replacement.Text = UCase(strJuizo)
        .Execute Replace:=wdReplaceAll
        'Se o juízo não foi encontrado no Sísifo, avisa para colocar manualmente.
        If strJuizo = "" Then MsgBox "Juízo não encontrado na base de dados no Sísifo. Lembre-se de acrescentá-lo manualmente no endereçamento da petição.", _
            vbCritical + vbOKOnly, "Alerta - juízo não encontrado"

        
        .Text = "<processo>"
        .Replacement.Text = strProcesso
        .Execute Replace:=wdReplaceAll
        
        .Text = "<adverso>"
        .Replacement.Text = strAdverso
        .Execute Replace:=wdReplaceAll
        
        If strPeticao = "alvará" Then
            .Text = "<eventos>"
            strCont = Application.InputBox(DeterminarTratamento & ", quais eventos têm depósitos judiciais para solicitar alvará para a Embasa?", "Informações sobre alvará", "11, 26 e 35", Type:=2)
            If strCont = False Then strCont = ""
            .Replacement.Text = strCont
            .Execute Replace:=wdReplaceAll
        End If
        
        If strPeticao = "compensação" Then
            .Text = "<débito-embasa>"
            .Replacement.Text = Format(curCondenacao, "#,##0.00")
            .Execute Replace:=wdReplaceAll
        
            .Text = "<débito-autor>"
            .Replacement.Text = Format(curDebitoMatricula, "#,##0.00")
            .Execute Replace:=wdReplaceAll
        
            .Text = "<diferença-para-autor>"
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
        
        If strPeticao = "execução" Then
            .Text = "<CPF-Autor>"
            strCont = Application.InputBox(DeterminarTratamento & ", qual o CPF/CNPJ do executado?", "Informações sobre execução", "", Type:=2)
            If strCont = False Then strCont = ""
            .Replacement.Text = strCont
            .Execute Replace:=wdReplaceAll
            
            .Text = "<natureza-condenação>"
            strCont = Application.InputBox(DeterminarTratamento & ", a que se refere o valor exeqüendo?", "Informações sobre execução", "honorários sucumbenciais", Type:=2)
            If strCont = False Then strCont = ""
            .Replacement.Text = strCont
            .Execute Replace:=wdReplaceAll
            
            .Text = "<valor-condenação>"
            cCont = Application.InputBox(DeterminarTratamento & ", qual o valor da condenação (antes da atualização)?", "Informações sobre execução", "", Type:=2)
            If cCont = False Then cCont = 0
            .Replacement.Text = Format(cCont, "#.##0,00")
            .Execute Replace:=wdReplaceAll
            
            .Text = "<termo-inicial-atualização>"
            strCont = Application.InputBox(DeterminarTratamento & ", qual foi o termo inicial da atualização?", "Informações sobre execução", "10/10/2019", Type:=2)
            If strCont = False Then strCont = ""
            .Replacement.Text = strCont
            .Execute Replace:=wdReplaceAll
            
            .Text = "<termo-inicial-juros>"
            strCont = Application.InputBox(DeterminarTratamento & ", qual foi o termo inicial dos juros de mora?", "Informações sobre execução", "10/10/2019", Type:=2)
            If strCont = False Then strCont = ""
            .Replacement.Text = strCont
            .Execute Replace:=wdReplaceAll
            
            .Text = "<valor-atualizado>"
            cCont = CCur(Application.InputBox(DeterminarTratamento & ", qual o valor da condenação atualizado?", "Informações sobre execução", "", Type:=2))
            If cCont = False Then cCont = 0
            .Replacement.Text = Format(cCont, "#.##0,00")
            .Execute Replace:=wdReplaceAll
            
            .Text = "<valor-multa>"
            .Replacement.Text = Format(cCont * 0.1, "#.##0,00")
            .Execute Replace:=wdReplaceAll
            
            .Text = "<valor-com-multa-e-honorários>"
            .Replacement.Text = Format(cCont * 1.2, "#.##0,00")
            .Execute Replace:=wdReplaceAll
            
        End If
        
        If strPeticao = "certidão de daje" Then
            .Text = "<data-pagamento>"
            strCont = Application.InputBox(DeterminarTratamento & ", quando foi feito o pagamento dos DAJEs?", "Informações sobre DAJEs", Format(Date - 13, "dd/mm/yyyy"), Type:=2)
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
    Case "compensação"
        strCont = "Pagamento e compensacao"
    Case "fazer"
        strCont = "Juntada fazer"
    Case "liminar"
        strCont = "Manifestacao liminar"
    Case "preparo"
        strCont = "Juntada preparo"
    Case "alvará"
        strCont = "Requerimento alvara"
    Case "liberarpenhora"
        strCont = "Liberacao de valor"
    Case "execução"
        strCont = "Requerimento de execucao"
    Case "certidão de daje"
        strCont = "Requerimento certidao"
    End Select
    
    ' Salvar
    
    
    If bolSsfPrazosBotaoPdfPressionado Then 'Gerar como PDF
        wdDocPeticao.ExportAsFixedFormat OutputFilename:=BuscarCaminhoPrograma & "01 " & strCont & " - " & SeparaPrimeirosNomes(strAdverso, 2) & ".pdf", _
            ExportFormat:=wdExportFormatPDF, OptimizeFor:=wdExportOptimizeForOnScreen, CreateBookmarks:=wdExportCreateHeadingBookmarks, BitmapMissingFonts:=False
        'MsgBox DeterminarTratamento & ", o PDF foi gerado com sucesso.", vbInformation + vbOKOnly, "Sísifo - Documento gerado"
        wdDocPeticao.Close wdDoNotSaveChanges
        If appword.Documents.Count = 0 Then appword.Quit
        
    Else ' Gerar como Word
        wdDocPeticao.SaveAs BuscarCaminhoPrograma & Format(Date, "yyyy.mm.dd") & " " & strCont & " - " & SeparaPrimeirosNomes(strAdverso, 2) & ".docx"
        wdDocPeticao.GoTo wdGoToPage, wdGoToFirst
        MsgBox DeterminarTratamento & ", o documento Word foi gerado com sucesso.", vbInformation + vbOKOnly, "Sísifo - Documento gerado"
    End If
    
End Sub


