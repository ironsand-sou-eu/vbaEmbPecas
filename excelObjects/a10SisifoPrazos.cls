VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a10SisifoPrazos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
''
'' Confere se o endere�o da pasta PPJCM est� correto. Se n�o estiver, configura.
''

    Dim strCont As String
    Dim bolConseguiuSalvar As Boolean
    
Reinicio:
    strCont = Dir(PegarDiretorioPPJCM, vbDirectory)
    If strCont <> "." Then ' Se n�o existe a pasta, pergunta
        MsgBox DeterminarTratamento & ", a pasta PPJCM armazenada na minha mem�ria n�o existe. Suplico que me mostre onde fica a pasta PPJCM " & _
        "neste computador!", vbCritical + vbOKOnly, "S�sifo - Pasta n�o encontrada"
        strCont = PerguntarDiretorio("S�sifo - Informar pasta das peti��es", False)
        If Right(strCont, 1) <> "\" Then strCont = strCont & "\"
        bolConseguiuSalvar = SalvarDiretorioPPJCM(strCont)
        GoTo Reinicio
        
    End If
    
End Sub
