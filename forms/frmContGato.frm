VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmContGato 
   Caption         =   "S�sifo - Mestre, fale-me um pouco mais sobre o processo!"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11325
   OleObjectBlob   =   "frmContGato.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmContGato"
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

Private Sub txtDataRetiradaGato_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim strCont As String

    strCont = Replace(txtDataRetiradaGato.Text, " ", "")
    strCont = Replace(strCont, "-", "")
    strCont = Replace(strCont, ".", "")
    strCont = Replace(strCont, "/", "")
    
    If IsNumeric(strCont) Then ' Se forem s� n�meros
        If Len(strCont) = 6 Then 'Dia, m�s e ano com dois d�gitos
            strCont = Format(strCont, "00/00/00")
            strCont = Left(strCont, 6) & "20" & Mid(strCont, 7)
        ElseIf Len(strCont) = 8 Then
            strCont = Format(strCont, "00/00/0000")
        End If
    
    Else ' Se n�o forem s� n�meros
        strCont = Trim(txtDataRetiradaGato.Text)
    
    End If
    
    txtDataRetiradaGato.Text = strCont
    
    If strCont <> "" And Not IsDate(strCont) Then
        MsgBox DeterminarTratamento & ", o valor """ & strCont & """ n�o parece ser uma data. " & _
            "O programa rodar� assim mesmo, mas decidi alertar a vossa Infalibil�ssima Gra�a, apenas caso queirais conferir " & _
            "alguma eventualidade fugaz no vosso perene conhecimento.", vbOKOnly, "S�sifo - Data n�o reconhecida"
    ElseIf strCont <> "" And IsDate(strCont) Then
        If Trim(Me.txtMesRefRegulaConsumo.Text) = "" Then
            strCont = CStr(Format((CDate(strCont) + 31), "dd/mm/yyyy"))
            strCont = Right(strCont, 7)
            Me.txtMesRefRegulaConsumo.Text = strCont
        End If
    End If

End Sub

Private Sub txtMesRefRegulaConsumo_Change()
'' Prev� o m�s de aplica��o das san��es
    Dim strCont As String
    
    If Len(Me.txtMesRefRegulaConsumo.Text) = 7 Then 'Se tiver 7 d�gitos, � o m�s-e-ano completo.
        If Trim(Me.txtMesRefMulta.Text) = "" Then
            strCont = Left(Me.txtMesRefRegulaConsumo.Text, 2)
            strCont = CStr(Format(CInt(strCont) + 2, "00"))
            Me.txtMesRefMulta.Text = strCont & Right(Me.txtMesRefRegulaConsumo.Text, 5)
        End If
    End If

End Sub

Private Sub txtMesRefRegulaConsumo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'' Corrige ano com dois d�gitos

    If Len(Me.txtMesRefRegulaConsumo.Text) = 5 Then Me.txtMesRefRegulaConsumo.Text = Left(Me.txtMesRefRegulaConsumo.Text, 2) & "/20" & Right(Me.txtMesRefRegulaConsumo.Text, 2)

End Sub

Private Sub txtMesRefRegulaConsumo_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'' Insere barra ao digitar o m�s
    If KeyCode <> 8 Or KeyCode <> 46 Then 'Se n�o foi backspace ou delete
        Select Case Len(Me.txtMesRefRegulaConsumo.Text)  ' Quantidade de caracteres da caixa de texto.
        Case 2 ' Se for 2, � hora de colocar uma barra, pois estamos ap�s o m�s.
            Me.txtMesRefRegulaConsumo.Text = Me.txtMesRefRegulaConsumo.Text & "/"
        End Select
    End If
    
End Sub

Private Sub txtMesRefRegulaConsumo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'' Somente n�meros
    If ValidaNumeros(KeyAscii) = False Then KeyAscii = 0

End Sub

Private Sub txtMesRefMulta_Exit(ByVal Cancel As MSForms.ReturnBoolean)
''Corrige ano com dois d�gitos

    If Len(Me.txtMesRefMulta.Text) = 5 Then Me.txtMesRefMulta.Text = Left(Me.txtMesRefMulta.Text, 2) & "/20" & Right(Me.txtMesRefMulta.Text, 2)

End Sub

Private Sub txtMesRefMulta_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode <> 8 Or KeyCode <> 46 Then 'Se n�o foi backspace ou delete
        Select Case Len(Me.txtMesRefMulta.Text)  ' Quantidade de caracteres da caixa de texto.
        Case 2 ' Se for 2, � hora de colocar uma barra, pois estamos ap�s o m�s.
            Me.txtMesRefMulta.Text = Me.txtMesRefMulta.Text & "/"
        End Select
    End If
    
End Sub

Private Sub txtMesRefMulta_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If ValidaNumeros(KeyAscii) = False Then KeyAscii = 0

End Sub

Private Sub txtValorMulta_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If ValidaNumeros(KeyAscii, ",") = False Then KeyAscii = 0

End Sub

Private Sub txtValorMulta_Exit(ByVal Cancel As MSForms.ReturnBoolean)
''
'' Totaliza as san��es, escreve em txtTotalSancoes e formata
''
    Dim X As Single, y As Single, z As Single
    
    X = IIf(Trim(Me.txtValorMulta.Text = ""), 0, Me.txtValorMulta.Text)
    y = IIf(Trim(Me.txtValorRecCons.Text = ""), 0, Me.txtValorRecCons.Text)
    z = IIf(Trim(Me.txtValorReparo.Text = ""), 0, Me.txtValorReparo.Text)
    
    Me.txtTotalSancoes.Text = Format(X + y + z, "#,##0.00")
    Me.txtValorMulta.Text = Format(Me.txtValorMulta.Text, "#,##0.00")

End Sub

Private Sub txtValorRecCons_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If ValidaNumeros(KeyAscii, ",") = False Then KeyAscii = 0

End Sub

Private Sub txtValorRecCons_Exit(ByVal Cancel As MSForms.ReturnBoolean)
''
'' Totaliza as san��es, escreve em txtTotalSancoes e formata
''
    Dim X As Single, y As Single, z As Single
    
    X = IIf(Trim(Me.txtValorMulta.Text = ""), 0, Me.txtValorMulta.Text)
    y = IIf(Trim(Me.txtValorRecCons.Text = ""), 0, Me.txtValorRecCons.Text)
    z = IIf(Trim(Me.txtValorReparo.Text = ""), 0, Me.txtValorReparo.Text)
    
    Me.txtTotalSancoes.Text = Format(X + y + z, "#,##0.00")
    Me.txtValorRecCons.Text = Format(Me.txtValorRecCons.Text, "#,##0.00")

End Sub

Private Sub txtValorReparo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If ValidaNumeros(KeyAscii, ",") = False Then KeyAscii = 0

End Sub

Private Sub txtValorReparo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
''
'' Totaliza as san��es, escreve em txtTotalSancoes e formata
''
    Dim X As Single, y As Single, z As Single
    
    X = IIf(Trim(Me.txtValorMulta.Text = ""), 0, Me.txtValorMulta.Text)
    y = IIf(Trim(Me.txtValorRecCons.Text = ""), 0, Me.txtValorRecCons.Text)
    z = IIf(Trim(Me.txtValorReparo.Text = ""), 0, Me.txtValorReparo.Text)
    
    Me.txtTotalSancoes.Text = Format(X + y + z, "#,##0.00")
    Me.txtValorReparo.Text = Format(Me.txtValorReparo.Text, "#,##0.00")

End Sub

Private Sub chbTemAlteracao2_Change()
'''''''''
'''''
'''''
        Label14.Enabled = chbTemAlteracao2.Value
        Label24.Enabled = chbTemAlteracao2.Value
        Label13.Enabled = chbTemAlteracao2.Value
        txtDataAlteracao2.Enabled = chbTemAlteracao2.Value
        txtRefAlteracao2.Enabled = chbTemAlteracao2.Value
        txtClassifAlteracao2.Enabled = chbTemAlteracao2.Value

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
