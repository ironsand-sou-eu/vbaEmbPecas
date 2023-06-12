VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRIDesabastecimento 
   Caption         =   "Sísifo - Mestre, fale-me um pouco mais sobre o processo!"
   ClientHeight    =   12630
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11340
   OleObjectBlob   =   "frmRIDesabastecimento.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRIDesabastecimento"
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

Private Sub cmbCorresponsavel_Change()

    If cmbCorresponsavel.Text = "" Or cmbCorresponsavel.Text = "Não houve outro responsável" Then
        chbExcluiuCorresp.Value = False
        chbExcluiuCorresp.Visible = False
        
    Else
        chbExcluiuCorresp.Caption = "Excluiu " & cmbCorresponsavel.Text
        chbExcluiuCorresp.Visible = True
        
    End If

End Sub

Private Sub txtValorCondenacao_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    txtValorCondenacao.Text = Format(txtValorCondenacao.Text, "#,##0.00")
    
End Sub

Private Sub UserForm_Initialize()
        Me.Height = 581.25
        Me.frmDesabastecimento.Height = 204
        Me.frmPedidos.Top = 306
        Me.frmInstSent.Top = 426
        Me.btGerar.Top = 528
        Me.txtProcessosSimilares.Visible = False
        
End Sub

Private Sub chbMultiplosProcessos_Change()
    
    If Me.chbMultiplosProcessos.Value = True Then ' Mostrar caixa de texto
        Me.Height = Me.Height + 72
        btGerar.Top = btGerar.Top + 72
        frmInstSent.Top = frmInstSent.Top + 72
        frmPedidos.Top = frmPedidos.Top + 72
        frmDesabastecimento.Height = frmDesabastecimento.Height + 72
        txtProcessosSimilares.Visible = True
        
    Else ' Ocultar caixa de texto
        Me.Height = Me.Height - 72
        btGerar.Top = btGerar.Top - 72
        frmInstSent.Top = frmInstSent.Top - 72
        frmPedidos.Top = frmPedidos.Top - 72
        frmDesabastecimento.Height = frmDesabastecimento.Height - 72
        txtProcessosSimilares.Visible = False
        
    End If
    
End Sub

Private Sub txtDataInicio_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim strCont As String
    
    strCont = Replace(txtDataInicio.Text, " ", "")
    strCont = Replace(strCont, "/", "")
    
    If IsNumeric(strCont) Then ' Se forem só números
        If Len(strCont) = 6 Then 'Dia, mês e ano com dois dígitos
            strCont = Format(strCont, "00/00/00")
            strCont = Left(strCont, 6) & "20" & Mid(strCont, 7)
        ElseIf Len(strCont) = 8 Then
            strCont = Format(strCont, "00/00/0000")
        End If
        
    Else ' Se não forem só números
        strCont = Trim(txtDataInicio.Text)
    
    End If
    
    txtDataInicio.Text = strCont
    
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
