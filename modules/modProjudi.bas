Attribute VB_Name = "modProjudi"
Option Explicit

Sub ProjudiAbrirTelaPeticionar()
''
'' Com o Projudi aberto e logado no Internet Explorer, abre a tela de peticionar de um processo no Projudi
''
    Dim IE As InternetExplorer
    Dim DocHTML As HTMLDocument
    Dim strNumeroProcesso As String, strCont As String
    Dim intCont As Integer
    Dim tbTabela As HTMLTable
    
    strNumeroProcesso = ActiveCell.Formula
    
    Set IE = New InternetExplorer
    
    'Pegar link pelo n�mero CNJ
    strCont = ProjudiPegaLinkPeticionar(strNumeroProcesso, IE, DocHTML)
    
    If strCont = "Sess�o expirada" Then
        MsgBox DeterminarTratamento & ", a sess�o expirou. Fa�a login no Projudi e tente novamente.", vbCritical + vbOKOnly, "S�sifo - Sess�o do Projudi expirada"
        Exit Sub
    ElseIf strCont = "Processo n�o encontrado" Then
        MsgBox DeterminarTratamento & ", o processo n�o foi encontrado. Verifique se o n�mero est� correto e tente novamente.", vbCritical + vbOKOnly, "S�sifo - Processo n�o encontrado"
        Exit Sub
    ElseIf strCont = "N�o abriu por demora" Then
        MsgBox DeterminarTratamento & ", o processo n�o abriu por demora. Provavelmente, a conex�o est� muito lenta. Tente novamente daqui a pouco.", vbCritical + vbOKOnly, "S�sifo - Tempo de espera expirado"
        Exit Sub
    End If
    
    IE.navigate strCont
    
    Do
        DoEvents
    Loop Until IE.readyState = READYSTATE_COMPLETE
    
    
    
End Sub

Function ProjudiPegaLinkPeticionar(ByVal strNumeroCNJ As String, ByRef IE As InternetExplorer, ByRef DocHTML As HTMLDocument) As String
''
'' Retorna o link da p�gina principal do processo strNumeroCNJ.
'' DEVO LIDAR COM O ERRO DE N�O ESTAR LOGADO!!!!!!!
''

    Dim lnkPeticionar As HTMLLinkElement
    Dim intCont As Integer

    IE.Visible = True
    IE.navigate sfURLBuscaProjudiAdvogado
        
    Set IE = RecuperarIE(sfURLBuscaProjudiAdvogado)
    If IE Is Nothing Then
        On Error Resume Next
        IE.Quit
        On Error GoTo 0
        ProjudiPegaLinkPeticionar = "N�o abriu por demora"
        Exit Function
    End If
    IE.Visible = True
    
    On Error GoTo Volta1
Volta1:
    Do
        DoEvents
    Loop Until IE.Document.readyState = "complete"
    
    Do
        DoEvents
    Loop Until IE.Document.getElementsByTagName("body")(0).Children(2).Children(0).Children(0).Children(0).Children(1).Children(0).innerText = "N�mero Processo"
    On Error GoTo 0
    
    ' Preenche o n�mero do processo na busca e submete o formul�rio
    Set DocHTML = IE.Document
    
    If DocHTML.Title = "Sistema CNJ - A sess�o expirou" Then
        ProjudiPegaLinkPeticionar = "Sess�o expirada"
        Exit Function
    End If
    
    DocHTML.getElementById("numeroProcesso").Value = strNumeroCNJ
    DocHTML.forms("busca").submit
    
    'Esperar 1
    ' No futuro: observar a requisi��o, para ver que valores j� voltam preenchidos e quais s�o criados de forma ass�ncrona, a� testar bom base em algum ass�ncrono.
    On Error GoTo Volta2
Volta2:
    Do
        DoEvents
    Loop Until IE.readyState = 4
    
    'Do
    '    DoEvents
    'Loop Until DocHTML.getElementsByTagName("body")(0).Children(0).Children(0).innerText = "Processos Obtidos Por Busca"
    
    Do
        intCont = DocHTML.getElementsByTagName("a").length - 1
        For intCont = 0 To intCont Step 1
            If Trim(DocHTML.getElementsByTagName("a")(intCont).innerText) = "Peticionar" Then Set lnkPeticionar = DocHTML.getElementsByTagName("a")(intCont)
        Next intCont
    Loop While lnkPeticionar Is Nothing
    On Error GoTo 0
    
    'COLOCAR UM TIMEOUT AQUI
    
    ' Procura pelo link
    If DocHTML.getElementsByTagName("a")(2) Is Nothing Then
        PegaLinkProcessoProjudi = ""
    End If
    
    ProjudiPegaLinkPeticionar = lnkPeticionar.href
    
End Function


