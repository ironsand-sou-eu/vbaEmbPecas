VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmContNegativacao 
   Caption         =   "Sísifo - Mestre, fale-me um pouco mais sobre o processo!"
   ClientHeight    =   11550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11325
   OleObjectBlob   =   "frmContNegativacao.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmContNegativacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btGerar_Click()
    
    chbDeveGerar.Value = True
    Me.Hide
    
End Sub

Private Sub chbComprovResidenciaDeTerceiro_Click()

    If Me.chbComprovResidenciaDeTerceiro.Value = True Then Me.chbSemComprovResidencia.Value = False

End Sub

Private Sub chbIncompetenciaTerritorial_Change()

    lblComarca.Visible = chbIncompetenciaTerritorial.Value
    txtComarcaCompetente.Visible = chbIncompetenciaTerritorial.Value
    
End Sub

Private Sub chbSemComprovResidencia_Change()

    If Me.chbSemComprovResidencia.Value = True Then Me.chbComprovResidenciaDeTerceiro.Value = False

End Sub

Private Sub cmbPerfilContrato_Change()

    If Me.cmbPerfilContrato.Value = "Houve uso e pagamentos regulares até certa data" Then
        Me.Label11.Visible = True
        Me.txtMesFinalUsoRegular.Visible = True
        
    Else
        Me.Label11.Visible = False
        Me.txtMesFinalUsoRegular.Visible = False
        
    End If
    
End Sub

Private Sub txtMatricula1_Change()

    txtMatricula2.Text = txtMatricula1.Text
    txtMatricula3.Text = txtMatricula1.Text
    
End Sub

Private Sub txtMesFinalUsoRegular_Exit(ByVal Cancel As MSForms.ReturnBoolean)
''Corrige ano com dois dígitos

If Len(Me.txtMesFinalUsoRegular.Text) = 5 Then Me.txtMesFinalUsoRegular.Text = Left(Me.txtMesFinalUsoRegular.Text, 2) & "/20" & Right(Me.txtMesFinalUsoRegular.Text, 2)

End Sub

Private Sub txtMesFinalUsoRegular_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    If ValidaNumeros(KeyAscii) = False Then KeyAscii = 0

End Sub

Private Sub txtMesFinalUsoRegular_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode <> 8 Or KeyCode <> 46 Then 'Se não foi backspace ou delete
        Select Case Len(Me.txtMesFinalUsoRegular.Text)  ' Quantidade de caracteres da categoria atual.
        Case 2 ' Se for 2, é hora de colocar uma barra, pois estamos após o mês.
            Me.txtMesFinalUsoRegular.Text = Me.txtMesFinalUsoRegular.Text & "/"
        End Select
    End If
    
End Sub


Private Sub txtVencim1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    AjustarVencimento Me.txtVencim1, Me.txtaaaamm1

End Sub

Private Sub txtVencim2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    AjustarVencimento Me.txtVencim2, Me.txtaaaamm2

End Sub

Private Sub txtVencim3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    AjustarVencimento Me.txtVencim3, Me.txtaaaamm3

End Sub

Private Sub AjustarVencimento(txt As MSForms.TextBox, txtMesAno As MSForms.TextBox)

    txt.Text = Trim(txt.Text)
    txt.Text = Replace(txt.Text, " ", "")
    txt.Text = Replace(txt.Text, "/", "")
    
    If Len(txt.Text) = 6 Then 'Dia, mês e ano com dois dígitos
        txt.Text = Format(txt.Text, "00/00/00")
        txt.Text = Left(txt.Text, 6) & "20" & Mid(txt.Text, 7)
        txtMesAno.Text = Right(txt.Text, 4) & Mid(txt.Text, 4, 2)
    ElseIf Len(txt.Text) = 8 Then
        txt.Text = Format(txt.Text, "00/00/0000")
        txtMesAno.Text = Right(txt.Text, 4) & Mid(txt.Text, 4, 2)
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

