VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSolicitaFazer 
   Caption         =   "Sísifo - Mestre, explique-me a obrigação de fazer!"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8100
   OleObjectBlob   =   "frmSolicitaFazer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSolicitaFazer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btGerar_Click()
    
    chbDeveGerar.Value = True
    Me.Hide
    
End Sub


Private Sub chbCancelarCobranca_Click()
' Habilita as caixas de texto com os meses e as cobranças a cancelar
    
    If chbCancelarCobranca.Value Then
        txtMesesCancelar.Enabled = True
        cmbCobrancaACancelar.Enabled = True
        txtMesesCancelar.Text = "05/2017 e 01/2018"

    Else
        txtMesesCancelar.Enabled = False
        cmbCobrancaACancelar.Enabled = False
        txtMesesCancelar.Text = ""
    End If
End Sub

Private Sub chbQuitar_Click()
' Habilita as caixas de texto com os meses a refaturar
    
    If chbQuitar.Value Then
        txtMesesQuitar.Enabled = True
        txtMesesQuitar.Text = "02, 03, 04 e 06/2017"
    Else
        txtMesesQuitar.Enabled = False
        txtMesesQuitar.Text = ""
    End If
End Sub

Private Sub chbRefat_Click()
' Habilita as caixas de texto com os meses e o valor a refaturar

    If chbRefat.Value Then
        txtMesesRef.Enabled = True
        txtValorRefat.Enabled = True
        txtMesesRef.Text = "07 a 12/2017"
        txtValorRefat.Text = "10 m3"
    Else
        txtMesesRef.Enabled = False
        txtValorRefat.Enabled = False
        txtMesesRef.Text = ""
        txtValorRefat.Text = ""
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Cancel = 1
    chbDeveGerar.Value = False
    Me.Hide
    
End Sub
