VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmContClasTarif 
   Caption         =   "Sísifo - Mestre, fale-me um pouco mais sobre o processo!"
   ClientHeight    =   11925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11325
   OleObjectBlob   =   "frmContClasTarif.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmContClasTarif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btGerar_Click()
    
    chbDeveGerar.Value = True
    Me.Hide
    
End Sub

Private Sub chbIncompetenciaTerritorial_Change()

    lblComarca.Visible = chbIncompetenciaTerritorial.Value
    txtComarcaCompetente.Visible = chbIncompetenciaTerritorial.Value
    
End Sub

Private Sub UserForm_Initialize()

        Me.Height = 498
        Me.frmPedidos.Top = 318
        Me.btGerar.Top = 440
        Me.frmCalcExemplo.Visible = False
        
End Sub

Private Sub chbApresentarCalcExemplo_Change()

    If chbApresentarCalcExemplo.Value = True Then
        Me.Height = Me.Height + 120
        frmPedidos.Top = frmPedidos.Top + 120
        btGerar.Top = btGerar.Top + 120
        frmCalcExemplo.Visible = True
        
    Else
        Me.Height = Me.Height - 120
        frmPedidos.Top = frmPedidos.Top - 120
        btGerar.Top = btGerar.Top - 120
        frmCalcExemplo.Visible = False
        
    End If

End Sub

Private Sub chbPeriodoMenor_Change()

        Label25.Enabled = chbPeriodoMenor.Value
        Label11.Enabled = chbPeriodoMenor.Value
        Label23.Enabled = chbPeriodoMenor.Value
        Label12.Enabled = chbPeriodoMenor.Value
        txtClassifOriginal.Enabled = chbPeriodoMenor.Value
        txtDataAlteracao1.Enabled = chbPeriodoMenor.Value
        txtRefAlteracao1.Enabled = chbPeriodoMenor.Value
        txtClassifAlteracao1.Enabled = chbPeriodoMenor.Value
        chbTemAlteracao2.Enabled = chbPeriodoMenor.Value
        
        If chbPeriodoMenor.Value = False Then
            Label14.Enabled = False
            Label24.Enabled = False
            Label13.Enabled = False
            txtDataAlteracao2.Enabled = False
            txtRefAlteracao2.Enabled = False
            txtClassifAlteracao2.Enabled = False
        Else
            Label14.Enabled = chbTemAlteracao2.Value
            Label24.Enabled = chbTemAlteracao2.Value
            Label13.Enabled = chbTemAlteracao2.Value
            txtDataAlteracao2.Enabled = chbTemAlteracao2.Value
            txtRefAlteracao2.Enabled = chbTemAlteracao2.Value
            txtClassifAlteracao2.Enabled = chbTemAlteracao2.Value
        End If

End Sub

Private Sub chbReconvencao_Change()
    
    Label27.Visible = chbReconvencao.Value
    txtTotalReconvencao.Visible = chbReconvencao.Value
    
End Sub

Private Sub chbTemAlteracao2_Change()

        Label14.Enabled = chbTemAlteracao2.Value
        Label24.Enabled = chbTemAlteracao2.Value
        Label13.Enabled = chbTemAlteracao2.Value
        txtDataAlteracao2.Enabled = chbTemAlteracao2.Value
        txtRefAlteracao2.Enabled = chbTemAlteracao2.Value
        txtClassifAlteracao2.Enabled = chbTemAlteracao2.Value

End Sub

Private Sub txtClassifAlteracao1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    If ValidaNumeros(KeyAscii, "/") = False Then KeyAscii = 0
    
End Sub

Private Sub txtClassifAlteracao1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode <> 8 Or KeyCode <> 46 Then 'Se não foi backspace ou delete
        If InserePontos(txtClassifAlteracao1.Text) = True Then txtClassifAlteracao1.Text = txtClassifAlteracao1.Text & "."
    End If
    
End Sub

Private Sub txtClassifAlteracao2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    If ValidaNumeros(KeyAscii, "/") = False Then KeyAscii = 0
    
End Sub

Private Sub txtClassifAlteracao2_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode <> 8 Or KeyCode <> 46 Then 'Se não foi backspace ou delete
        If InserePontos(txtClassifAlteracao2.Text) = True Then txtClassifAlteracao2.Text = txtClassifAlteracao2.Text & "."
    End If
    
End Sub

Private Sub txtClassifOriginal_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    If ValidaNumeros(KeyAscii, "/") = False Then KeyAscii = 0
    
End Sub

Private Sub txtClassifOriginal_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode <> 8 Or KeyCode <> 46 Then 'Se não foi backspace ou delete
        If InserePontos(txtClassifOriginal.Text) = True Then txtClassifOriginal.Text = txtClassifOriginal.Text & "."
    End If
    
End Sub

Private Sub txtDataAlteracao1_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim strCont As String

    strCont = Replace(txtDataAlteracao1.Text, " ", "")
    strCont = Replace(strCont, "-", "")
    strCont = Replace(strCont, ".", "")
    strCont = Replace(strCont, "/", "")
    
    If IsNumeric(strCont) Then ' Se forem só números
        If Len(strCont) = 6 Then 'Dia, mês e ano com dois dígitos
            strCont = Format(strCont, "00/00/00")
            strCont = Left(strCont, 6) & "20" & Mid(strCont, 7)
        ElseIf Len(strCont) = 8 Then
            strCont = Format(strCont, "00/00/0000")
        End If
    
    Else ' Se não forem só números
        strCont = Trim(txtDataAlteracao1.Text)
    
    End If
    
    txtDataAlteracao1.Text = strCont
    
    If strCont <> "" And Not IsDate(strCont) Then MsgBox DeterminarTratamento & ", o valor """ & strCont & """ não parece ser uma data. " & _
            "O programa rodará assim mesmo, mas decidi alertar a vossa Infalibilíssima Graça, apenas caso queirais conferir " & _
            "alguma eventualidade fugaz no vosso perene conhecimento.", vbOKOnly, "Sísifo - Data não reconhecida"

End Sub

Private Sub txtDataAlteracao2_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim strCont As String

    strCont = Replace(txtDataAlteracao2.Text, " ", "")
    strCont = Replace(strCont, "-", "")
    strCont = Replace(strCont, ".", "")
    strCont = Replace(strCont, "/", "")
    
    If IsNumeric(strCont) Then ' Se forem só números
        If Len(strCont) = 6 Then 'Dia, mês e ano com dois dígitos
            strCont = Format(strCont, "00/00/00")
            strCont = Left(strCont, 6) & "20" & Mid(strCont, 7)
        ElseIf Len(strCont) = 8 Then
            strCont = Format(strCont, "00/00/0000")
        End If
    
    Else ' Se não forem só números
        strCont = Trim(txtDataAlteracao2.Text)
    
    End If
    
    txtDataAlteracao2.Text = strCont
    
    If strCont <> "" And Not IsDate(strCont) Then MsgBox DeterminarTratamento & ", o valor """ & strCont & """ não parece ser uma data. " & _
            "O programa rodará assim mesmo, mas decidi alertar a vossa Infalibilíssima Graça, apenas caso queirais conferir " & _
            "alguma eventualidade fugaz no vosso perene conhecimento.", vbOKOnly, "Sísifo - Data não reconhecida"

End Sub

Private Sub txtConsExemplo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If ValidaNumeros(KeyAscii) = False Then KeyAscii = 0

End Sub

Private Sub txtExFaturado1Cat_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If ValidaNumeros(KeyAscii) = False Then KeyAscii = 0

End Sub

Private Sub txtExFaturado1Cat_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode <> 8 Or KeyCode <> 46 Then 'Se não foi backspace ou delete
        If Len(Me.txtExFaturado1Cat.Text) = 1 And InserePontos(Me.txtExFaturado1Cat.Text) = True Then Me.txtExFaturado1Cat.Text = Me.txtExFaturado1Cat.Text & "."
        If Len(Me.txtExFaturado1Cat.Text) = 3 Then Me.txtExFaturado1Economias.SetFocus
    End If
    
End Sub

Private Sub txtExFaturado1Economias_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If ValidaNumeros(KeyAscii) = False Then KeyAscii = 0

End Sub

Private Sub chbTemDoisTipos_Change()

    Me.txtExFaturado2Cat.Enabled = chbTemDoisTipos.Value
    Me.txtExFaturado2Economias.Enabled = chbTemDoisTipos.Value
    Me.txtExFaturado2Cat.SetFocus

End Sub

Private Sub txtExFaturado2Cat_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If ValidaNumeros(KeyAscii) = False Then KeyAscii = 0

End Sub

Private Sub txtExFaturado2Cat_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode <> 8 Or KeyCode <> 46 Then 'Se não foi backspace ou delete
        If Len(Me.txtExFaturado2Cat.Text) = 1 And InserePontos(Me.txtExFaturado2Cat.Text) = True Then Me.txtExFaturado2Cat.Text = Me.txtExFaturado2Cat.Text & "."
        If Len(Me.txtExFaturado2Cat.Text) = 3 Then Me.txtExFaturado2Economias.SetFocus
    End If
    
End Sub

Private Sub txtExFaturado2Economias_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If ValidaNumeros(KeyAscii) = False Then KeyAscii = 0

End Sub

Private Sub txtExPretCat_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If ValidaNumeros(KeyAscii) = False Then KeyAscii = 0

End Sub

Private Sub txtExPretCat_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode <> 8 Or KeyCode <> 46 Then 'Se não foi backspace ou delete
        If Len(Me.txtExPretCat.Text) = 1 And InserePontos(Me.txtExPretCat.Text) = True Then Me.txtExPretCat.Text = Me.txtExPretCat.Text & "."
        If Len(Me.txtExPretCat.Text) = 3 Then Me.txtExPretEconomias.SetFocus
    End If
    
End Sub

Private Sub txtExPretEconomias_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If ValidaNumeros(KeyAscii) = False Then KeyAscii = 0

End Sub

Private Sub txtMesRefExemplo_Change()
''
'' Alerta se for julho ou agosto
''
    
    Select Case Left(Me.txtMesRefExemplo.Text, 2)
    Case "07", "08"
        MsgBox DeterminarTratamento & ", os meses de julho e agosto não são interessantes para o exemplo de cálculo, pois costumam abranger o período " & _
                "de reajuste tarifário, resultando num cálculo mais complicado, com duas tarifas diferentes. Favor escolher outro período.", _
                vbOKOnly + vbExclamation, "Sísifo - Mês para a fatura de exemplo de cálculo"
        Me.txtMesRefExemplo.Text = ""
    End Select
        
End Sub

Private Sub txtMesRefExemplo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
''Corrige ano com dois dígitos

    If Len(Me.txtMesRefExemplo.Text) = 5 Then Me.txtMesRefExemplo.Text = Left(Me.txtMesRefExemplo.Text, 2) & "/20" & Right(Me.txtMesRefExemplo.Text, 2)

End Sub

Private Sub txtMesRefExemplo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    If ValidaNumeros(KeyAscii) = False Then KeyAscii = 0

End Sub

Private Sub txtMesRefExemplo_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode <> 8 Or KeyCode <> 46 Then 'Se não foi backspace ou delete
        Select Case Len(Me.txtMesRefExemplo.Text)  ' Quantidade de caracteres da categoria atual.
        Case 2 ' Se for 2, é hora de colocar uma barra, pois estamos após o mês.
            Me.txtMesRefExemplo.Text = Me.txtMesRefExemplo.Text & "/"
        End Select
    End If
    
End Sub

Private Sub txtRefAlteracao1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
''Corrige ano com dois dígitos

    If Len(Me.txtRefAlteracao1.Text) = 5 Then Me.txtRefAlteracao1.Text = Left(Me.txtRefAlteracao1.Text, 2) & "/20" & Right(Me.txtRefAlteracao1.Text, 2)

End Sub

Private Sub txtRefAlteracao1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    If ValidaNumeros(KeyAscii) = False Then KeyAscii = 0

End Sub

Private Sub txtRefAlteracao1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode <> 8 Or KeyCode <> 46 Then 'Se não foi backspace ou delete
        Select Case Len(Me.txtRefAlteracao1.Text)  ' Quantidade de caracteres da categoria atual.
        Case 2 ' Se for 2, é hora de colocar uma barra, pois estamos após o mês.
            Me.txtRefAlteracao1.Text = Me.txtRefAlteracao1.Text & "/"
        End Select
    End If
    
End Sub

Private Sub txtRefAlteracao2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
''Corrige ano com dois dígitos

    If Len(Me.txtRefAlteracao2.Text) = 5 Then Me.txtRefAlteracao2.Text = Left(Me.txtRefAlteracao2.Text, 2) & "/20" & Right(Me.txtRefAlteracao2.Text, 2)

End Sub

Private Sub txtRefAlteracao2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    If ValidaNumeros(KeyAscii) = False Then KeyAscii = 0

End Sub

Private Sub txtRefAlteracao2_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode <> 8 Or KeyCode <> 46 Then 'Se não foi backspace ou delete
        Select Case Len(Me.txtRefAlteracao2.Text)  ' Quantidade de caracteres da categoria atual.
        Case 2 ' Se for 2, é hora de colocar uma barra, pois estamos após o mês.
            Me.txtRefAlteracao2.Text = Me.txtRefAlteracao2.Text & "/"
        End Select
    End If
    
End Sub

Private Sub chbDanMat_Click()

    chbDanMatSemProvas.Visible = chbDanMat.Value
    chbValorLucroCessante.Visible = chbDanMat.Value
    chbDanMatSemProvas.Value = False
    chbValorLucroCessante.Value = False
    
End Sub

Private Sub chbDanMorCorte_Click()

    If chbDanMorCorte.Value = True Then chbDanMorMeraCobranca.Value = False
    
End Sub

Private Sub chbDanMorMeraCobranca_Click()

    If chbDanMorMeraCobranca.Value = True Then
        chbDanMorCorte.Value = False
        chbDanMorNegativacao.Value = False
    End If
    
End Sub

Private Sub chbDanMorNegativacao_Click()

    If chbDanMorNegativacao.Value = True Then chbDanMorMeraCobranca.Value = False
    
End Sub

Private Sub chbDanoMoral_Change()
    
    chbDanMorCorte.Visible = chbDanoMoral.Value
    chbDanMorNegativacao.Visible = chbDanoMoral.Value
    chbDanMorMeraCobranca.Visible = chbDanoMoral.Value
    optAutorPF.Enabled = chbDanoMoral.Value
    optAutorCondominio.Enabled = chbDanoMoral.Value
    optAutorOutrosPJ.Enabled = chbDanoMoral.Value
    Label5.Enabled = chbDanoMoral.Value
    
    chbDanMorCorte.Value = False
    chbDanMorNegativacao.Value = False
    chbDanMorMeraCobranca.Value = False
    
End Sub

Private Sub chbDevolDobro_Change()

    cmbPagamento.Enabled = chbDevolDobro.Value

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Cancel = 1
    chbDeveGerar.Value = False
    Me.Hide
    
End Sub
