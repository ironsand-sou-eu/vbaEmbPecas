VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmJuntPagamento 
   Caption         =   "Sísifo - Mestre, fale-me um pouco mais sobre o pagamento!"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6330
   OleObjectBlob   =   "frmJuntPagamento.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmJuntPagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btGerar_Click()
    
    chbDeveGerar.Value = True
    Me.Hide
    
End Sub

Private Sub chbExisteDebito_Click()
' Torna visíveis as caixas de texto com o valor do débito e da condenação
    
    If chbExisteDebito.Value Then
        Label2.Enabled = True
        txtDebMatricula.Enabled = True
        Label3.Enabled = True
        txtValCondenacao.Enabled = True
    Else
        Label2.Enabled = False
        txtDebMatricula.Enabled = False
        Label3.Enabled = False
        txtValCondenacao.Enabled = False
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Cancel = 1
    chbDeveGerar.Value = False
    Me.Hide
    
End Sub
