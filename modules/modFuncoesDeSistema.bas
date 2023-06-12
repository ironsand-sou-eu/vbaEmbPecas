Attribute VB_Name = "modFuncoesDeSistema"
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)
 
Function CaminhoDesktop() As String
    CaminhoDesktop = CreateObject("WScript.Shell").specialfolders("desktop")
End Function

Sub AoCarregarRibbon(ByVal plPlan As Worksheet, Ribbon As IRibbonUI)
    ' Guarda a referência ao objeto Ribbon
    Dim lnRibbon As LongPtr
    lnRibbon = ObjPtr(Ribbon)
    plPlan.Cells().Find(what:="Ponteiro do Ribbon", lookat:=xlWhole).Offset(0, 1).Formula = "'" & lnRibbon
    
End Sub

Function RecuperarObjetoPorReferencia(arq As Workbook) As IRibbonUI
    ' Recupera o objeto pela referência
    Dim lnRefObjeto As LongPtr, rbObjeto As IRibbonUI
    
    lnRefObjeto = arq.Sheets("cfConfigurações").Cells().Find(what:="Ponteiro do Ribbon", lookat:=xlWhole).Offset(0, 1).Text
    CopyMemory rbObjeto, lnRefObjeto, 6
    Set RecuperarObjetoPorReferencia = rbObjeto
End Function

Sub LiberarEdicao(arq As Workbook)
' Apresentar dados para as planilhas de configurações serem alterados pelo usuário
    Dim rib As IRibbonUI
    
    arq.IsAddin = False
    Set rib = RecuperarObjetoPorReferencia(arq)
    
    Select Case Right(arq.CodeName, Len(arq.CodeName) - 9)
    Case "Intimacoes"
        rib.InvalidateControl "btFechaConfigIntimacoes"
    Case "Prazos"
        rib.InvalidateControl "btFechaConfigPrazos"
    End Select
        
    MsgBox "Planilhas de configuração liberadas para edição. Tenha cuidado, e só realize alterações conforme as " & _
            "instruções fornecidas em cada planilha.", vbInformation + vbOKOnly, "Sísifo - Liberando alterações"
End Sub

Function DeterminarTratamento() As String
''
'' Vai à planilha cfTratamentos, pega um adjetivo ou superlativo e um pronome/substantivo de tratamento.
''

    Dim plan As Worksheet
    Dim rngCont As Range
    Dim intCont As Integer
    Dim strTratamento As String
    
    Set plan = ThisWorkbook.Sheets("cfTratamentos")
    
    ' Define se será superlativo ou normal. intChance 1 = normal; 2 = superlativo
    Randomize
    If CInt(100 * Rnd + 1) <= 60 Then intChance = 1 Else intChance = 2
    
    ' Conta os adjetivos e escolhe um aleatoriamente
    Set rngCont = plan.Cells().Find(IIf(intChance = 1, "Adjetivos", "Superlativos"), lookat:=xlWhole).Offset(1, 0)
    Set rngCont = Range(rngCont, rngCont.End(xlDown))
    Randomize
    intCont = CInt((rngCont.Cells().Count) * Rnd + 1)
    strTratamento = rngCont.Cells(intCont).Text
        
    ' Conta os substantivos e escolhe um aleatoriamente
    Set rngCont = plan.Cells().Find("Substantivos", lookat:=xlWhole).Offset(1, 0)
    Set rngCont = Range(rngCont, rngCont.End(xlDown))
    Randomize
    intCont = CInt((rngCont.Cells().Count) * Rnd + 1)
    strTratamento = strTratamento & " " & rngCont.Cells(intCont).Text
    
    DeterminarTratamento = strTratamento
    
End Function

Function RecuperarIE(strTrechoURLProcurada As String) As InternetExplorer
''
'' Reatribui o objeto InternetExplorer para a variável IE, perdida por causa da saída da intranet.
''

    Dim Shell As Shell32.Shell
    Dim CadaIE As Variant
    Dim snInicioTimer As Single
    
    snInicioTimer = Timer
IEVazio:
    'Do
    Set Shell = New Shell32.Shell
    For Each CadaIE In Shell.Windows
        If InStr(1, CadaIE.LocationURL, strTrechoURLProcurada) <> 0 Then Exit For
    Next CadaIE
    'Loop
    
    If Timer >= snInicioTimer + 10 Then GoTo TempoEsgotado
    If CadaIE = Empty Then GoTo IEVazio

    Set RecuperarIE = CadaIE
    Exit Function
    
TempoEsgotado:
    Set RecuperarIE = Nothing
    
End Function

Sub RestringirEdicaoRibbon(arq As Workbook, Optional ByVal controle As IRibbonControl)
    arq.IsAddin = True
    Application.DisplayAlerts = False
    arq.Save
    'arq.SaveAs Filename:=arq.FullName, FileFormat:=xlOpenXMLAddIn
    Application.DisplayAlerts = True
    
    If Not controle Is Nothing Then
        Select Case Right(arq.CodeName, Len(arq.CodeName) - 9)
        Case "Intimacoes"
            FechaConfigVisivel "btFechaConfigIntimacoes", controle
        Case "Prazos"
            FechaConfigVisivel "btFechaConfigPrazos", controle
        End Select
    End If
    
End Sub

Sub FechaConfigVisivel(strControle As String, Optional ByVal controle As IRibbonControl, Optional ByRef returnedVal)
    Dim rib As IRibbonUI
    
    Set rib = RecuperarObjetoPorReferencia(ThisWorkbook)
    returnedVal = Not ThisWorkbook.IsAddin
    rib.InvalidateControl strControle
    
End Sub
