Attribute VB_Name = "Módulo1"
Sub AjustarHconSci()
    
    Dim plan As Worksheet
    Dim lnUltimaLinha As Long
    Dim lnUltimaColuna As Long
    
    Set plan = ActiveSheet
    
    With plan
        If .Cells(1, 6).MergeArea.Address <> "$F$1:$G$1" Then .Range("$F$1:$G$1").Merge
        If .Cells(2, 1).MergeArea.Address <> "$A$2:$D$2" Then .Range("$A$2:$D$2").Merge
        .Rows(3).RowHeight = 30
        .Columns(1).ColumnWidth = 9
        .Columns(2).ColumnWidth = 3
        .Columns(3).ColumnWidth = 10
        .Columns(4).ColumnWidth = 12
        .Columns(6).ColumnWidth = 12
        .Columns(8).ColumnWidth = 8
        .Columns(9).ColumnWidth = 3
        
        lnUltimaLinha = .UsedRange.Rows.Count
        If .Cells(lnUltimaLinha, 1).Formula = "" Then .Rows(lnUltimaLinha).Delete
        
        lnUltimaLinha = .UsedRange.Rows.Count
        lnUltimaColuna = 10
        .Range("$A$3:$J$" & lnUltimaLinha).Cells.Borders(xlInsideVertical).LineStyle = 1
        .Range("$A$3:$J$" & lnUltimaLinha).Cells.Borders(xlInsideHorizontal).LineStyle = 1
        .Range("$A$3:$J$" & lnUltimaLinha).Cells.Borders(xlEdgeRight).LineStyle = 1
        .Range("$A$3:$J$" & lnUltimaLinha).Cells.Borders(xlEdgeLeft).LineStyle = 1
        .Range("$A$3:$J$" & lnUltimaLinha).Cells.Borders(xlEdgeTop).LineStyle = 1
        .Range("$A$3:$J$" & lnUltimaLinha).Cells.Borders(xlEdgeBottom).LineStyle = 1
        
        .PageSetup.CenterHorizontally = True
        .PageSetup.PrintTitleRows = "$3:$3"
        
    End With
    
End Sub

Sub GerarTodasRespostasRpvsEmPdf()
    Dim rng As Excel.Range
    Dim stateBotaoPdfPressionado As Boolean
    
    Set rng = ActiveCell
    
    stateBotaoPdfPressionado = bolSsfPrazosBotaoPdfPressionado
    bolSsfPrazosBotaoPdfPressionado = True
    Do Until rng.Formula = ""
        rng.Activate
        JuntadaRespostaRpv
        Set rng = rng.Offset(1, 0)
    Loop
    bolSsfPrazosBotaoPdfPressionado = stateBotaoPdfPressionado
    
End Sub
