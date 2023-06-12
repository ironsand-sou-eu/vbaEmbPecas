VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmContRevConsElevado 
   Caption         =   "S�sifo - Mestre, fale-me um pouco mais sobre o processo!"
   ClientHeight    =   10230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11655
   OleObjectBlob   =   "frmContRevConsElevado.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmContRevConsElevado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chbEsgotoAposCorte_Change()

    If chbEsgotoAposCorte.Value = True Then
        chbDanMorCorte.Value = True
    End If
    
End Sub

Private Sub chbIncompetenciaTerritorial_Change()

    lblComarca.Visible = chbIncompetenciaTerritorial.Value
    txtComarcaCompetente.Visible = chbIncompetenciaTerritorial.Value
    
End Sub

Private Sub chbMediaCorreta_Click()
    
    chbMediaConsRetificado.Visible = chbMediaCorreta.Value
    chbMediaConsRetificado.Value = chbMediaCorreta.Value
    
End Sub

Private Sub btGerar_Click()
    
    chbDeveGerar.Value = True
    Me.Hide
    
End Sub

Private Sub cmbAferHidrometro_Change()
' Se houver aferi��o, desmarca a caixa de pedir aferi��o
    If cmbAferHidrometro.Text <> "N�o h�" Then chbRequerAfericao.Value = False
    
End Sub

Private Sub cmbPadraoConsumo_Change()
' A depender do caso, marca a caixa de pedir aferi��o
    If cmbPadraoConsumo.Text = "H� padr�o, consumo impugnado exorbitou e continua alterado" Or _
    cmbPadraoConsumo.Text = "H� padr�o, mas o consumo impugnado � razoavelmente compat�vel com a m�dia" Or _
    cmbPadraoConsumo.Text = "N�o h� padr�o definido, consumo cheio de altos e baixos" Then
        chbRequerAfericao.Value = True
    Else
        chbRequerAfericao.Value = False
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
