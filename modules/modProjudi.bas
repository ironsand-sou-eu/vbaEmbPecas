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
    
    'Pegar link pelo número CNJ
    strCont = ProjudiPegaLinkPeticionar(strNumeroProcesso, IE, DocHTML)
    
    If strCont = "Sessão expirada" Then
        MsgBox DeterminarTratamento & ", a sessão expirou. Faça login no Projudi e tente novamente.", vbCritical + vbOKOnly, "Sísifo - Sessão do Projudi expirada"
        Exit Sub
    ElseIf strCont = "Processo não encontrado" Then
        MsgBox DeterminarTratamento & ", o processo não foi encontrado. Verifique se o número está correto e tente novamente.", vbCritical + vbOKOnly, "Sísifo - Processo não encontrado"
        Exit Sub
    ElseIf strCont = "Não abriu por demora" Then
        MsgBox DeterminarTratamento & ", o processo não abriu por demora. Provavelmente, a conexão está muito lenta. Tente novamente daqui a pouco.", vbCritical + vbOKOnly, "Sísifo - Tempo de espera expirado"
        Exit Sub
    End If
    
    IE.navigate strCont
    
    Do
        DoEvents
    Loop Until IE.readyState = READYSTATE_COMPLETE
    
    
    
End Sub

Function ProjudiPegaLinkPeticionar(ByVal strNumeroCNJ As String, ByRef IE As InternetExplorer, ByRef DocHTML As HTMLDocument) As String
''
'' Retorna o link da página principal do processo strNumeroCNJ.
'' DEVO LIDAR COM O ERRO DE NÃO ESTAR LOGADO!!!!!!!
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
        ProjudiPegaLinkPeticionar = "Não abriu por demora"
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
    Loop Until IE.Document.getElementsByTagName("body")(0).Children(2).Children(0).Children(0).Children(0).Children(1).Children(0).innerText = "Número Processo"
    On Error GoTo 0
    
    ' Preenche o número do processo na busca e submete o formulário
    Set DocHTML = IE.Document
    
    If DocHTML.Title = "Sistema CNJ - A sessão expirou" Then
        ProjudiPegaLinkPeticionar = "Sessão expirada"
        Exit Function
    End If
    
    DocHTML.getElementById("numeroProcesso").Value = strNumeroCNJ
    DocHTML.forms("busca").submit
    
    'Esperar 1
    ' No futuro: observar a requisição, para ver que valores já voltam preenchidos e quais são criados de forma assíncrona, aí testar bom base em algum assíncrono.
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


