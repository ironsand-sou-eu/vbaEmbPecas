VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRICorte 
   Caption         =   "Sísifo - Mestre, fale-me um pouco mais sobre o processo!"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6825
   OleObjectBlob   =   "frmRICorte.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRICorte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub btGerar_Click()
    
    chbDeveGerar.Value = True
    Me.Hide
    
End Sub

Private Sub chbDanosMateriais_Change()
    
    If chbDanosMateriais.Value = False Then chbCondenouDanosMateriais.Value = False
    
End Sub

Private Sub chbDevDobro_Click()

    If chbDevDobro.Value = False Then chbCondenouDevDobro.Value = False
    
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

Private Sub frmSentenca_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    txtValorCondenacao.Text = Format(txtValorCondenacao.Text, "#,##0.00")
    
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

Private Sub txtValorCondenacao_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    txtValorCondenacao.Text = Format(txtValorCondenacao.Text, "#,##0.00")
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Cancel = 1
    chbDeveGerar.Value = False
    Me.Hide
    
End Sub
