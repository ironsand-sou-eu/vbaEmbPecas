VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCRRIDesabApagXingu2018 
   Caption         =   "Sísifo - Mestre, fale-me um pouco mais sobre o processo!"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8295
   OleObjectBlob   =   "frmCRRIDesabApagXingu2018.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCRRIDesabApagXingu2018"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btGerar_Click()
    
    chbDeveGerar.Value = True
    Me.Hide
    
End Sub

Private Sub frmSentenca_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    txtValorCondenacao.Text = Format(txtValorCondenacao.Text, "#,##0.00")
    
End Sub

Private Sub optProcedenteEmParte_Change()

    txtValorCondenacao.Enabled = optProcedenteEmParte.Value
    'if optprocedenteemparte.Value = true then

End Sub

Private Sub txtValorCondenacao_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    txtValorCondenacao.Text = Format(txtValorCondenacao.Text, "#,##0.00")
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Cancel = 1
    chbDeveGerar.Value = False
    Me.Hide
    
End Sub
