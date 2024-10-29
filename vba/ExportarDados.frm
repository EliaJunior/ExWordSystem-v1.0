VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExportarDados 
   Caption         =   "UserForm1"
   ClientHeight    =   3588
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9012.001
   OleObjectBlob   =   "ExportarDados.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExportarDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0
#Const EarlyBind = False
Public dic As Object

Sub ConverterWordParaPDF(caminhoArquivo As String, caminhoPDF As String)
    ' Err Handler
    On Error GoTo escape
    
    ' Inicia o Word de forma invisível
    #If EarlyBind Then
        Dim wordApp As New Word.Application
        Dim doc     As Word.Document
    #Else
        Dim wordApp As Object
        Dim doc     As Object
        
        Set wordApp = CreateObject("Word.Application")
    #End If

    wordApp.Visible = False
    
    ' Abre o documento .docx
    Set doc = wordApp.Documents.Open(caminhoArquivo)
    
    ' Exporta o documento como PDF
    doc.ExportAsFixedFormat OutputFileName:=caminhoPDF, ExportFormat:=17
        

    ' Fecha o documento e o Word
    doc.Close False
escape:
    wordApp.Quit

    ' Libera memória
    Set doc = Nothing
    Set wordApp = Nothing
End Sub
Private Sub SubstituirCamposArquivoWord(pathModel As String, fields As Object, dirOut As String, fileName As String)
    '-------------------------------------------------------------------------------------
    'Substitui os campos de um arquivo .docx ou .xml com os itens de um objeto Scripting.Dictionary
    '   Parâmetros:
    '       pathModel -> Caminho do arquivo modelo [...\Arquivo.docx ou ...\Arquivo.xml]
    '       fields    -> Scripting.Dictionary onde as chaves são os campos a serem substituidos [Chave:Valor]
    '       dirOut    -> Diretório onde os arquivos serão salvos
    '       fileName  -> Nome dos arquivos de saída (sem extensão)
    '   Retorno:
    '       Retorna um arquivo .docx e um arquivo .pdf no diretório de saída
    '-------------------------------------------------------------------------------------
    
    #If EarlyBind Then
        Dim fso As New Scripting.FileSystemObject
        Dim xmlFile As TextStream
    #Else
        Dim fso As Object, xmlFile As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
    #End If
    
    Dim tempFolder As String: tempFolder = Environ("Temp") & "\DocxTemp\"
    Dim modelExt As String: modelExt = fso.GetExtensionName(pathModel)
    Dim xmlFilePath As String, tempZipPath As String, zipOutput As String, docOutput As String, pdfOutput As String
    Dim xmlContent As String, campo As Variant
    
    ' Limpa a pasta temporária, se existir, e configura os caminhos
    If fso.FolderExists(tempFolder) Then fso.DeleteFolder fso.GetFolder(tempFolder), True
    fso.CreateFolder tempFolder
    
    dirOut = dirOut & "\"
    zipOutput = dirOut & fileName & ".zip"
    docOutput = dirOut & fileName & ".docx"
    pdfOutput = dirOut & fileName & ".pdf"
    xmlFilePath = IIf(modelExt = "docx", tempFolder & "word\document.xml", tempFolder & "TempXML.xml")
    
    ' Copia e extrai o modelo para a pasta temporária
    If modelExt = "docx" Then
        tempZipPath = tempFolder & fso.GetFileName(pathModel)
        fso.CopyFile pathModel, Replace(tempZipPath, ".docx", ".zip")
        ZipUnzipFile Replace(tempZipPath, ".docx", ".zip"), tempFolder
        Application.Wait Now + TimeValue("00:00:03")
        fso.DeleteFile (Replace(tempZipPath, ".docx", ".zip"))
    Else
        fso.CopyFile pathModel, xmlFilePath, True
    End If
    
    ' Substitui os campos no XML e salva
    If fso.FileExists(xmlFilePath) Then
        Set xmlFile = fso.OpenTextFile(xmlFilePath, 1, False, -2)
        xmlContent = xmlFile.ReadAll: xmlFile.Close
        
        For Each campo In fields.Keys: xmlContent = Replace(xmlContent, campo, StringToEntities(fields(campo))): Next campo

        Set xmlFile = fso.CreateTextFile(xmlFilePath, True)
        xmlFile.Write xmlContent: xmlFile.Close
        
        ' Compacta novamente para .docx ou converte diretamente para PDF
        If modelExt = "docx" Then
            Open zipOutput For Output As #1: Close #1
            ZipUnzipFile tempFolder, zipOutput
            Application.Wait Now + TimeValue("00:00:03")
            fso.CopyFile zipOutput, docOutput
            fso.DeleteFile zipOutput
            ConverterWordParaPDF docOutput, pdfOutput
        Else
            ConverterWordParaPDF xmlFilePath, pdfOutput
        End If
    Else
        MsgBox "Erro: Arquivo XML não encontrado!", vbExclamation
    End If

    fso.DeleteFolder fso.GetFolder(tempFolder)
    Set fso = Nothing: Set xmlFile = Nothing
End Sub

Private Sub exportarDados_Click()
    ' Verificações
    If chaves.Value = "" Or Not IsRangeValido(chaves.Value) Or valores.Value = "" Or Not IsRangeValido(valores.Value) Or _
       dirOut.Value = "" Or pathTemplateWord.Value = "" Then
        MsgBox "Necessário preenchimento de todos os campos para prosseguir!", vbExclamation: Exit Sub
    End If
    
    ' Inicia o dicionario
    IniciarDicionario
    
    ' Inicia o procedimento de exportação
    Dim nomeArquivo As String
    nomeArquivo = "exported_file_" & Format(Now, "ddmmyyyy_hhmm")
    SubstituirCamposArquivoWord pathTemplateWord, dic, dirOut.Value, nomeArquivo
    
    ' Conclusão
    MsgBox "Procedimento concluído, o arquivo foi exportado para: " & Chr(10) & nomeArquivo & ".pdf", vbInformation
End Sub

Private Sub selecionarDiretorioSaida_Click()
    Dim path As String
    
    ' Seleciona o caminho do arquivo word (docx ou xml)
    path = PegarCaminho(msoFileDialogFolderPicker, "Selecionar o diretório de saída", False)
    
    ' Verifica se a seleção do arquivo foi bem sucedida
    If path = "" Then MsgBox "Procedimento cancelado, nenhuma pasta foi selecionada", vbExclamation: Exit Sub: dirOut.Value = ""
    
    ' Escreve na text box
    dirOut.Value = path
End Sub

Private Sub selecionarModeloWord_Click()
    Dim pathModel As String
    
    ' Seleciona o caminho do arquivo word (docx ou xml)
    pathModel = PegarCaminho(msoFileDialogFilePicker, "Selecionar template Word", False, Array("Arquivos Word", "*.docx,*.xml"))
    
    ' Verifica se a seleção do arquivo foi bem sucedida
    If pathModel = "" Then MsgBox "Procedimento cancelado, nenhum arquivo selecionado", vbExclamation: Exit Sub: pathTemplateWord.Value = ""
    
    ' Escreve na text box
    pathTemplateWord.Value = pathModel
End Sub

Private Sub selectCamposChaves_Click()
    Dim rngSelecionado As range
    
    ' Solicita ao usuário que selecione um intervalo
    On Error Resume Next
    Set rngSelecionado = Application.InputBox("Selecione o intervalo que contém as chaves:", Type:=8)
    On Error GoTo 0
    
    ' Verifica se o usuário fez uma seleção
    If Not rngSelecionado Is Nothing Then
        chaves.Value = "'" & rngSelecionado.Parent.Name & "'!" & rngSelecionado.Address(False, False)
    Else
        MsgBox "Nenhum intervalo foi selecionado.", vbExclamation
    End If
End Sub

Private Sub selectCamposValues_Click()
    Dim rngSelecionado As range
    
    ' Verifica se o range de chaves ja foi preenchido
    If (chaves.Value = "") Then MsgBox "Necessário definir o campo 'Chaves'", vbExclamation: Exit Sub
    
    ' Verifica se o campo chaves é um range valido
    If Not IsRangeValido(chaves.Value) Then MsgBox "O intervalo do campo 'Chaves' não é válido!", vbExclamation: Exit Sub
    
    ' Solicita ao usuário que selecione um intervalo
    On Error Resume Next
    Set rngSelecionado = Application.InputBox("Selecione o intervalo que contém os valores:", Type:=8)
    On Error GoTo 0
    
    ' Verifica se o usuário fez uma seleção
    If Not rngSelecionado Is Nothing Then
        ' Verifica se o tamanho do range de valores é igual ao tamanho do range de chaves
        
        If UBound(range(chaves.Value).Value) <> UBound(rngSelecionado.Value) Then MsgBox "O intervalo 'Valores' deve ser do mesmo tamanho do intervalo de 'Chaves'", vbExclamation: Exit Sub: valores.Value = ""
        
        valores.Value = "'" & rngSelecionado.Parent.Name & "'!" & rngSelecionado.Address(False, False)
    Else
        MsgBox "Nenhum intervalo foi selecionado.", vbExclamation
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Formata o formulário
    FormatarFormulario Me
    
    ' Inicia os campos
    On Error Resume Next
    With ThisWorkbook.Worksheets("__tempSheet__")
        pathTemplateWord.Value = .range("A1").Value
        dirOut.Value = CorrigirEndereco(.range("A2").Value)
        chaves.Value = CorrigirEndereco(.range("A3").Value)
        valores.Value = CorrigirEndereco(.range("A4").Value)
    End With
    On Error GoTo 0
    
    ' Diretório de saída padrão
    If dirOut.Value = "" Then dirOut.Value = ThisWorkbookPath
    
    ' Caption
    Me.Caption = "ExWord System v1.0"
End Sub
Private Sub IniciarDicionario()
    Dim arChaves As Variant
    Dim arValores As Variant
    Dim i As Long, j As Long
    Dim key As String, vl As String
    
    ' Reseta o dicionario
    Set dic = Nothing
    
    ' Inicia instancia do dicionario
    #If EarlyBind Then
        Set dic = New Scripting.Dictionary
    #Else
        Set dic = CreateObject("Scripting.Dictionary")
    #End If
    
    ' Atribui as chaves e valores
    arChaves = range(chaves.Value)
    arValores = range(valores.Value)
    
    j = LBound(arChaves, 2)
    
    For i = LBound(arChaves) To UBound(arChaves)
        key = CStr(arChaves(i, j))
        vl = CStr(arValores(i, j))
        If Not dic.Exists(key) Then dic.Add key, vl
    Next i
End Sub
Public Sub ZipUnzipFile(ByVal sZipFile, ByVal sDestinationPath)
    On Error GoTo Error_Handler
    #If EarlyBind = True Then
        Dim oShell            As Shell32.Shell
        Dim oFolderSrc        As Shell32.Folder
        Dim oFolderDest       As Shell32.Folder
        Dim oFolderSrcItems   As Shell32.folderItems

        Set oShell = New Shell32.Shell
    #Else
        Dim oShell            As Object
        Dim oFolderSrc        As Object
        Dim oFolderDest       As Object
        Dim oFolderSrcItems   As Variant

        Set oShell = CreateObject("Shell.Application")
    #End If

    Set oFolderSrcItems = oShell.Namespace(sZipFile).Items()
    If Not oFolderSrcItems Is Nothing Then
        Set oFolderDest = oShell.Namespace(sDestinationPath)
        If Not oFolderDest Is Nothing Then
            If oFolderSrcItems.Count <> 0 Then
                oFolderDest.CopyHere oFolderSrcItems, 4
            End If
        End If
    End If

Error_Handler_Exit:
    On Error Resume Next
    Set oFolderSrcItems = Nothing
    Set oFolderDest = Nothing
    Set oShell = Nothing
    Exit Sub

Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Source: Shell_UnZipFile" & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Sub
Private Function StringToEntities(texto As String) As String
    Dim i As Integer
    Dim codigoChar As Long
    Dim resultado As String

    ' Percorre cada caractere na string
    For i = 1 To Len(texto)
        codigoChar = Asc(Mid(texto, i, 1))
        
        ' Verifica se o caractere está fora do intervalo ASCII comum (0-127)
        If codigoChar < 32 Or codigoChar > 126 Then
            ' Transforma em entidade XML
            resultado = resultado & "&#" & codigoChar & ";"
        Else
            ' Adiciona o caractere normalmente
            resultado = resultado & Mid(texto, i, 1)
        End If
    Next i
    StringToEntities = resultado
End Function
Sub FormatarFormulario(xFrame As Object)
    Dim c As Control, cTipo As String
    Dim CorBorda As Long, CorFundoFrame As Long, corPreto As Long, corBranco As Long, corFonte As Long, corFundoLabel As Long
    Dim stTAG As String
    
    CorBorda = 15904775
    CorFundoFrame = 4538690
    corPreto = RGB(0, 0, 0)
    corBranco = RGB(255, 255, 255)
    corFonte = corBranco
    corFundoLabel = RGB(255, 255, 255)
    
    ' Fundo do objeto
    xFrame.BackColor = CorFundoFrame

    For Each c In xFrame.Controls
        cTipo = TypeName(c)
        With c
            If c.Tag = "LockedNotEnabled" Then c.Locked = True: c.Enabled = False
            Select Case cTipo
                    
                Case "Label"
                    .Font.Size = 11
                    .Font.Name = "Roboto"
                    .ForeColor = corFonte
                    .BackStyle = fmBackStyleTransparent

                Case "ComboBox"
                    .Font.Size = 11
                    .Font.Name = "Roboto"
                    .ShowDropButtonWhen = fmShowDropButtonWhenFocus
                    .BorderStyle = fmBorderStyleSingle
                    .BorderColor = CorBorda
                    .Height = 18
                    .MatchEntry = fmMatchEntryNone
                    .MatchRequired = True
                    .Style = fmStyleDropDownCombo

                Case "TextBox"
                    .Font.Size = 11
                    .Font.Name = "Roboto"
                    .BorderStyle = fmBorderStyleSingle
                    .BorderColor = CorBorda
                    .Height = 18

    
                Case "CommandButton"
                    .BackColor = corBranco
                    
                Case "OptionButton"
                    .Font.Size = 11
                    .Font.Name = "Roboto"
                    .ForeColor = corFonte
                    .BackStyle = fmBackStyleTransparent
                    .TextAlign = fmTextAlignLeft
                    
                Case "Frame"
                    .BackColor = CorFundoFrame
                    .BorderStyle = fmBorderStyleSingle
                    .Caption = ""
                    .BorderColor = CorBorda
            End Select
        End With
    Next c
End Sub
Private Function IsRangeValido(strRange As String) As Boolean
    Dim testeRange As range
    
    On Error Resume Next
    Set testeRange = range(strRange)
    IsRangeValido = Not testeRange Is Nothing
    On Error GoTo 0
End Function
Function PegarCaminho(MsoTipo As Long, titulo As String, SelecaoMultipla As Boolean, Optional arExtensoes As Variant) As Variant
    '-----------------------------------------------------------------------------
    'Retorna o caminho do arquivo ou pasta selecionada
    '   Parâmetros:
    '       MsoTipo  -> Pasta ou Arquivo
    '       MsoTipo  -> titulo
    '       MsoTipo  -> permitir seleção multipla
    '       MsoTipo  -> Extensões possíveis
    'Autor: Elias Junior||eng.eliasocjunior@gmail.com||(86)99993-0217
    'Última alteração 21/02/2024
    '------------------------------------------------------------------------------
    Dim pathOut As String
    Dim i As Long
    Dim curDesc As String, curExt As String
    
    With Application.FileDialog(MsoTipo)
        .Title = titulo
        .AllowMultiSelect = SelecaoMultipla
        
        If Not IsMissing(arExtensoes) Then
            curDesc = arExtensoes(0) 'Descrição da extensão
            curExt = arExtensoes(1)  'Extensão
            With .Filters
                .Clear
                .Add curDesc, curExt, 1
            End With
        End If
        
        .show
        If .SelectedItems.Count = 0 Then
            PegarCaminho = ""
            Exit Function
        End If
        If Not SelecaoMultipla Then
            PegarCaminho = .SelectedItems(1)
        Else
            PegarCaminho = .SelectedItems
        End If
    End With
End Function
Private Function ThisWorkbookFullPath() As String
    Dim oneDrivePart As String
    Dim FullPath As String
    
    FullPath = ThisWorkbook.FullName
    FullPath = VBA.Replace(FullPath, "/", "\")
    oneDrivePart = "https:\\d.docs.live.net\"
    If VBA.InStr(FullPath, oneDrivePart) Then
        FullPath = VBA.Replace(FullPath, oneDrivePart, "")
        FullPath = Right(FullPath, Len(FullPath) - VBA.InStr(1, FullPath, "\"))
        FullPath = Environ$("OneDriveConsumer") & "\" & FullPath
    End If
    ThisWorkbookFullPath = FullPath
End Function
Private Function ThisWorkbookPath() As String
    Dim oneDrivePart As String
    Dim xPath As String
    
    xPath = ThisWorkbook.path
    xPath = VBA.Replace(xPath, "/", "\")
    oneDrivePart = "https:\\d.docs.live.net\"
    If VBA.InStr(xPath, oneDrivePart) Then
        xPath = VBA.Replace(xPath, oneDrivePart, "")
        xPath = Right(xPath, Len(xPath) - VBA.InStr(1, xPath, "\"))
        xPath = Environ$("OneDriveConsumer") & "\" & xPath
    End If
    ThisWorkbookPath = xPath
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Cria uma planilha, se não existir, temporária com os campos deste userform
    Dim wsAtiva As Worksheet
    Dim wsTemp As Worksheet
    
    Set wsAtiva = ActiveSheet

    On Error Resume Next
    If sheetIndex("__tempSheet__") = 0 Then
        Set wsTemp = ThisWorkbook.Worksheets.Add
        wsTemp.Name = "__tempSheet__"
        With wsTemp
            .range("A1").Value = pathTemplateWord.Value
            .range("A2").Value = dirOut.Value
            .range("A3").Value = chaves.Value
            .range("A4").Value = valores.Value
        End With
    Else
        Set wsTemp = ThisWorkbook.Worksheets("__tempSheet__")
        With wsTemp
            .range("A1").Value = pathTemplateWord.Value
            .range("A2").Value = dirOut.Value
            .range("A3").Value = chaves.Value
            .range("A4").Value = valores.Value
        End With
    End If
    
    wsAtiva.Activate
End Sub
Private Function sheetIndex(wsNome As String) As Long
    On Error Resume Next
    sheetIndex = ThisWorkbook.Worksheets(wsNome).Index
    If Err.Number <> 0 Then sheetIndex = 0: Exit Function
End Function
Private Function CorrigirEndereco(strRange As String) As String
    ' Adiciona uma aspa simples no início, caso esteja faltando
    If Left(strRange, 1) <> "'" And InStr(strRange, "!") > 0 Then
        CorrigirEndereco = "'" & strRange
    Else
        CorrigirEndereco = strRange
    End If
End Function
