Attribute VB_Name = "modTradXML"
Option Explicit

Enum EstiloCaractere
    ssfItalico
    ssfNegrito
    ssfSublinhado
    ssfCorVermelha
    ssfCorVerde
End Enum

Sub TesteTraducao()
    SalvaremXMLSisifo BuscarCaminhoPrograma & "consumo rateado\Cont2.docx"
End Sub

Sub SalvaremXMLSisifo(strArquivoWordComCaminho As String)
''
'' Abre um documento do word e o traduz para o XML do Sísifo.
''
    Dim appword As Object
    Dim wdDocTraducao As Word.Document
    Dim rngCont As Word.Range
    Dim lngCont As Long
    Dim strNomeArquivo As String, strCaminho As String, strTitulo As String
    
    ' Criar o documento novo
    Set appword = New Word.Application
    appword.Visible = True
    Set wdDocTraducao = appword.Documents.Add(BuscarCaminhoPrograma & "modelos-automaticos\PPJCM Modelo.dotx")
    
    ' Copiar texto do original para o novo
    InserirArquivo strArquivoWordComCaminho, wdDocTraducao
    
    If Len(wdDocTraducao.Paragraphs.Last.Range.Text) = 1 Then
        If Asc(wdDocTraducao.Paragraphs.Last.Range.Text) = 13 Then
            wdDocTraducao.Paragraphs.Last.Range.Delete
        End If
    End If
    
    ''''''''''''''''''''''''''
    ' Fazer as substituições '
    ''''''''''''''''''''''''''
    
    RealizarSubstituicoes wdDocTraducao
    
    '''''''''''''''''''
    ' Salvar e fechar '
    '''''''''''''''''''
    
    ' Ajustar nome do arquivo
    strTitulo = "Sísifo - Salvar arquivo XML"
    strNomeArquivo = strCaminho & "xml\" & strNomeArquivo
    strNomeArquivo = Replace(strNomeArquivo, ".docx", ".ssfx")
    'strNomeArquivo = PerguntarNomeArquivo(strTitulo, strCaminho & "xml\")
    
    ' Salvar
    If strNomeArquivo <> "" Then
        'wdDocTraducao.SaveAs2 Filename:=strNomeArquivo, FileFormat:=xlTextWindows, AddToRecentFiles:=False, Encoding:=msoEncodingUTF8, InsertLineBreaks:=False, AllowSubstitutions:=True
        SalvarTXT wdDocTraducao, strNomeArquivo
    Else
        MsgBox DeterminarTratamento & ", o arquivo não pôde ser salvo porque o caminho e/ou nome escolhido foi inválido. Favor tentar novamente.", vbOKOnly + vbCritical, "Sísifo - Erro ao salvar"
    End If
    
    ' Fechar
    wdDocTraducao.Close savechanges:=wdDoNotSaveChanges
    
End Sub

Sub RealizarSubstituicoes(wdDoc As Word.Document)
''
'' Substitui os estilos de parágrafo, caractere, comentários e formas por tags XML
''
    ' Substitui os estilos de parágrafo
    SustituirFormatoEstiloPorTag wdDoc, "PPJCM Título", "<par:Titulo>"
    SustituirFormatoEstiloPorTag wdDoc, "PPJCM Subtítulo", "<par:Subtitulo>"
    SustituirFormatoEstiloPorTag wdDoc, "PPJCM Tipo de petição", "<par:TipoPeticao>"
    SustituirFormatoEstiloPorTag wdDoc, "PPJCM Cabecalho", "<par:Cabecalho>"
    SustituirFormatoEstiloPorTag wdDoc, "PPJCM Citação", "<par:Citacao>"
    SustituirFormatoEstiloPorTag wdDoc, "PPJCM Lista", "<par:Lista>"
    SustituirFormatoEstiloPorTag wdDoc, "PPJCM Peticao", "<par:Peticao>"
    SustituirFormatoEstiloPorTag wdDoc, "PPJCM Rodapé", "<par:Rodape>"
    SustituirFormatoEstiloPorTag wdDoc, "PPJCM Tabelas", "<par:Tabelas>"
    
    ' Substitui os estilos de caractere
    SustituirFormatoCaracterePorTag wdDoc, ssfCorVerde
    SustituirFormatoCaracterePorTag wdDoc, ssfItalico
    SustituirFormatoCaracterePorTag wdDoc, ssfNegrito
    SustituirFormatoCaracterePorTag wdDoc, ssfSublinhado
    SustituirFormatoCaracterePorTag wdDoc, ssfCorVermelha
    
    ' Substitui os comentários
    If wdDoc.Comments.Count >= 1 Then
        For lngCont = 1 To wdDoc.Comments.Count Step 1
            SustituirComentarioPorTag wdDoc, lngCont
        Next lngCont
    End If
    
    ' Substitui as formas
    If wdDoc.Shapes.Count >= 1 Then
        For lngCont = wdDoc.Shapes.Count To 1 Step -1
            SustituirFormaPorTag wdDoc, lngCont
        Next lngCont
    End If
    
    If Len(wdDoc.Paragraphs.Last.Range.Text) = 1 Then
        If Asc(wdDoc.Paragraphs.Last.Range.Text) = 13 Then
            wdDoc.Paragraphs.Last.Range.Delete
        End If
    End If
    
End Sub

Sub SustituirFormatoEstiloPorTag(wdDoc As Word.Document, strEstiloOrigem As String, strTagAbertura As String)

    Dim strTagFechamento As String
    strTagFechamento = Replace(strTagAbertura, "<", "</")
    
    ' Substitui
    With wdDoc.Range.Find
        .ClearFormatting
        .Style = wdDoc.Styles(strEstiloOrigem)
        .Replacement.ClearFormatting
        .Replacement.Style = wdDoc.Styles("Normal")
        .Text = ""
        .Replacement.Text = strTagAbertura & "^&" & strTagFechamento & "^p"
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

    ' Tira a quebra de parágrafo
    With wdDoc.Range.Find
        .ClearFormatting
        .Style = wdDoc.Styles("Normal")
        .Replacement.ClearFormatting
        .Replacement.Style = wdDoc.Styles("Normal")
        .Text = "^p" & strTagFechamento
        .Replacement.Text = strTagFechamento
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

End Sub

Sub SustituirFormatoCaracterePorTag(wdDoc As Word.Document, ecFormato As EstiloCaractere)

    Dim strTagFechamento As String
    strTagFechamento = Replace(strTagAbertura, "<", "</")
    
    ' Substitui
    With wdDoc.Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        Select Case ecFormato
        Case ssfItalico
            .Font.Italic = True
            .Replacement.Font.Italic = False
            .Replacement.Text = "<fnt:i>^&</i>"
            .Replacement.Font.Italic = False
        Case ssfNegrito
            .Font.Bold = True
            .Replacement.Font.Bold = False
            .Replacement.Text = "<fnt:b>^&</b>"
        Case ssfSublinhado
            .Font.Underline = wdUnderlineSingle
            .Replacement.Font.Underline = wdUnderlineNone
            .Replacement.Text = "<fnt:u>^&</u>"
        Case ssfCorVermelha
            .Font.ColorIndex = wdRed
            .Replacement.Font.ColorIndex = wdAuto
            .Replacement.Text = "<fnt:Vermelho>^&</fnt:Vermelho>"
        Case ssfCorVerde
            .Font.ColorIndex = wdGreen
            .Replacement.Font.ColorIndex = wdAuto
            .Replacement.Text = "<fnt:Verde>^&</fnt:Verde>"
        End Select
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    
End Sub

Sub SustituirComentarioPorTag(wdDoc As Word.Document, lngNumComentario As Long)

    Dim strTagAbertura As String, strTagFechamento As String
    
    strTagFechamento = "</cmt:Comentario" & intnumcometario & ">"
    
    ' Substitui
    With wdDoc.Comments(lngNumComentario)
        strTagAbertura = "<cmt:Comentario" & lngNumComentario & " TextoComentario=""" & .Range & """>"
        .Scope.InsertBefore (strTagAbertura)
        .Scope.InsertAfter (strTagFechamento)
        .DeleteRecursively
        
    End With
    
End Sub

Sub SustituirFormaPorTag(wdDoc As Word.Document, lngNumForma As Long)

    Dim rngCont As Range
    Dim strCont As String
    Dim strTagAbertura As String, strTagFechamento As String
    
    strTagFechamento = "</frm:Forma" & lngNumForma & ">"
    
    ' Substitui
    With wdDoc.Shapes(lngNumForma)
        strCont = " Type=" & .AutoShapeType & " Left=" & .Left & " Top=" & .Top & " Width=" & .Width & " Height=" & .Height & _
            " BackgroundStyle=" & .BackgroundStyle & " Fill.ForeColor.RGB=" & .Fill.ForeColor.RGB & " Name=" & .Name & """"
        strTagAbertura = "<frm:Forma" & lngNumForma & strCont & ">"
        .Anchor.InsertBefore (strTagAbertura)
        .Anchor.InsertAfter (strTagFechamento)
        .Delete
        
    End With
    
End Sub

Sub SalvarTXT(wdDoc As Word.Document, strNomeArquivo As String)
''
''
''
    Dim parCont As Word.Paragraph
    Dim intNumArq As Integer
    
    intNumArq = FreeFile
    Open strNomeArquivo For Random Lock Write As #intNumArq
    Close #intNumArq
    
    intNumArq = FreeFile
    Open strNomeArquivo For Output Lock Write As #intNumArq
    
    For Each parCont In wdDoc.Paragraphs
        Write #intNumArq, parCont.Range.Text
    Next parCont
    
    Close #intNumArq
    
End Sub

