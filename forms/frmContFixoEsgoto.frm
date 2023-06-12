VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmContFixoEsgoto 
   Caption         =   "Sísifo - Mestre, fale-me um pouco mais sobre o processo!"
   ClientHeight    =   9510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11295
   OleObjectBlob   =   "frmContFixoEsgoto.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmContFixoEsgoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbhHouveVicio_Click()

    chbVicioConsertado.Visible = cbhHouveVicio.Value
    chbVicioConsertado.Value = False
    
End Sub

Private Sub chbIncompetenciaTerritorial_Change()

    lblComarca.Visible = chbIncompetenciaTerritorial.Value
    txtComarcaCompetente.Visible = chbIncompetenciaTerritorial.Value
    
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

Private Sub txtConsExemplo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If ValidaNumeros(KeyAscii) = False Then KeyAscii = 0

End Sub

Private Sub txtExCat1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If ValidaNumeros(KeyAscii) = False Then KeyAscii = 0

End Sub

Private Sub txtExCat1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode <> 8 Or KeyCode <> 46 Then 'Se não foi backspace ou delete
        If Len(Me.txtExCat1.Text) = 1 And InserePontos(Me.txtExCat1.Text) = True Then Me.txtExCat1.Text = Me.txtExCat1.Text & "."
        If Len(Me.txtExCat1.Text) = 3 Then Me.txtExEconomias1.SetFocus
    End If
    
End Sub

Private Sub txtExEconomias1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If ValidaNumeros(KeyAscii) = False Then KeyAscii = 0

End Sub

Private Sub chbTemDoisTipos_Change()

    Me.txtExCat2.Enabled = chbTemDoisTipos.Value
    Me.txtExEconomias2.Enabled = chbTemDoisTipos.Value
    Me.txtExCat2.SetFocus

End Sub

Private Sub txtExCat2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If ValidaNumeros(KeyAscii) = False Then KeyAscii = 0

End Sub

Private Sub txtExCat2_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode <> 8 Or KeyCode <> 46 Then 'Se não foi backspace ou delete
        If Len(Me.txtExCat2.Text) = 1 And InserePontos(Me.txtExCat2.Text) = True Then Me.txtExCat2.Text = Me.txtExCat2.Text & "."
        If Len(Me.txtExCat2.Text) = 3 Then Me.txtExEconomias2.SetFocus
    End If
    
End Sub

Private Sub txtExEconomias2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If ValidaNumeros(KeyAscii) = False Then KeyAscii = 0

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

Private Sub btGerar_Click()
    
    chbDeveGerar.Value = True
    Me.Hide
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Cancel = 1
    chbDeveGerar.Value = False
    Me.Hide
    
End Sub
