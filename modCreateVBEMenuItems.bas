Attribute VB_Name = "modCreateVBEMenuItems"
Option Explicit

'VERSaO  1.01.0
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Feito Por: Ronan Vico
'Descricao: Este modulo possui Rotinas para criacao do botao na Barra de Comandos do VBE (Visual Basic Editor)
'           e necessario toda vez que iniciar a aplicacao instanciar a barra novamente ,pois ela funciona com eventos
'           Tambem e possivel rodar manualmente a rotina InitVBRVTool.
'Como usar?: Apenas rode InitVBRVTool e ela instanciara a barra de comando.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private MenuEvent As CVBECommandHandler
Private CmdBarItem As CommandBarButton 'CommandBarControl
Private cbBarTOOL As Office.CommandBarPopup
Public EventHandlers As New Collection
Private cmbar As Office.CommandBar


Private Const C_TAG = "MY_VBE_TAG"
Private Const C_RV_TOOLS_BAR As String = "RV"
Public Const C_APPNAME As String = "RVTool"
Public Const C_SECTION_CopyText As String = "CopyText"
Public Const C_SECTION_PasteText As String = "CopyText"



Private Enum ControlsType
        msoControlButton = 1
        msoControlDropdown = 3
        msoControlComboBox = 4
        msoControlPopup = 10
End Enum

Sub InitVBRVTool()
Dim cpop                               As CommandBarPopup
Dim cpopPasteText                      As CommandBarPopup
Dim cpopCopyText                       As CommandBarPopup
Dim i                                  As Long
Dim settingPasteText                   As String
Dim FaceID                             As Long
Dim DicButtons                         As New Scripting.Dictionary


    Call DeleteMenuItems
    '''''''''''''''''''''''''''''''''''''''''''''''''
    ' Delete any existing event handlers.
    '''''''''''''''''''''''''''''''''''''''''''''''''
    Do Until EventHandlers.Count = 0
        EventHandlers.Remove 1
    Loop

    '''''''''''''''''''''''''''''''''''''''''''''''''
    ' add the first control to the Tools menu.
    '''''''''''''''''''''''''''''''''''''''''''''''''
    On Error Resume Next
    Set cmbar = Application.VBE.CommandBars("Barra de menus")
    If cmbar Is Nothing Then
        Set cmbar = Application.VBE.CommandBars(1)
    End If
    On Error GoTo 0

    Set cbBarTOOL = cmbar.FindControl(tag:=C_RV_TOOLS_BAR)
    If cbBarTOOL Is Nothing Then
        With cmbar.Controls.Add(10, , , cmbar.Controls.Count + 1, False)
            .tag = C_RV_TOOLS_BAR
            .CAption = "RV&Tools"
            .BeginGroup = True
            .Visible = True
        End With
    End If




    With DicButtons
    
If PT_BR() Then
        .Add "Snippets", "Auto Completar Code Snippe&t"
        '--
        .Add "Inserir/Editar", "&Inserir e Editar"
        .Add "InsertProcedureHeader", "Inserir &Cabecalho"
        .Add "insertErrorTreatment", "Inserir &Error Treatment"
        .Add "InsertLineNumber", "Inserir &Numeracao nas Linhas"
        .Add "RemoveLineNumber", "&Remover Numeracao nas Linhas"
        .Add "IndentVariables", "&Identar Variaveis"
        '--
        .Add "AuxText", "Aux Textos"
        .Add "toUpperCase", "Texto Selecionado para &Maiusculo aA"
        .Add "toLowerCase", "Texto Selecionado para Mi&nusculo Aa"
        
        
        .Add "CopyText", "Copiar_&C"
        .Add "PasteText", "Colar_&V"
        .Add "CopyTextDesc", "Copiar texto para area "
        .Add "CleanPasteText", "Limpar Tudo"
        '--
        .Add "CheckVariablesNotUsedInProcedure", "Verificar Variaveis nao Utilizadas"
        '--
        .Add "Listar", "Listar Procedures"
        .Add "GetFunctionAndSubNames", "Imprimir TUDO"
        .Add "GetFunctionAndSubNameAtual", "Imprimir Modulo Atual"
        .Add "CloseProjectExplorer", "Fecha&r Project Explorer"
        .Add "CloseAllWindowsCodeModule", "Fechar All VBE &Windows"
        .Add "Hook", "Desbloquear All VBE's"
        '--
        .Add "CoresEditor", "&Alterar Cores Editor"
        .Add "Change_color_Dark_Theme", "DARK THEME"
        .Add "Change_color_White_Theme", "WHITE THEME DEFAULT"
        '--
        .Add "Atualizar_RVTool", "Update RVTOOL"
        '--
        .Add "aboutme", "About Creator"
        '--
        .Add "IndentarProcedure", "Indentar &Procedure"
        .Add "Change_Region", "Change Tool Lenguage to English"
Else
        .Add "Snippets", "Auto Complete Code Snippe&t"
        '--
        .Add "Inserir/Editar", "Edit / &Insert"
        .Add "InsertProcedureHeader", "Insert &Header"
        .Add "InsertLineNumber", "Insert &Error Treatment"
        .Add "RemoveLineNumber", "Insert Line &Number"
        .Add "IndentVariables", "&Ident Variables"
        '--
        .Add "AuxText", "Aux Texts"
        .Add "toUpperCase", "Selected Text TO UPPER CASE aA"
        .Add "toLowerCase", "Selected Text TO LOWER CASE aA"
        
        
        .Add "CopyText", "&Copy"
        .Add "PasteText", "&Paste"
        .Add "CopyTextDesc", "Copy selected Text to &"
        .Add "CleanPasteText", "Clean All"
        '--
        .Add "CheckVariablesNotUsedInProcedure", "Check &Unused Variables"
        '--
        .Add "Listar", "Print Procedures"
        .Add "GetFunctionAndSubNames", "Debug Print &All"
        .Add "GetFunctionAndSubNameAtual", "Debug Print Active &Module"
        .Add "CloseProjectExplorer", "Close Project Explo&rer"
        .Add "CloseAllWindowsCodeModule", "Close All Code &Windows"
        .Add "Hook", "Unlock All VBE's"
        '--
        .Add "CoresEditor", "Change &Editor Colors"
        .Add "Change_color_Dark_Theme", "DARK THEME"
        .Add "Change_color_White_Theme", "WHITE THEME DEFAULT"
        '--
        .Add "Atualizar_RVTool", "Update RVTOOL"
        '--
        .Add "aboutme", "About Creator"
        '--
        .Add "IndentarProcedure", "Ident &Procedure"
        .Add "Change_Region", "Trocar Tool para Portugues"
End If
    End With
    
    Set cbBarTOOL = cmbar.FindControl(tag:=C_RV_TOOLS_BAR)
    Call AddMenuButton(DicButtons("Snippets"), True, "Snippets", 7581)
    '-------------------------------------------------------------------------------
    Set cpop = cbBarTOOL.Controls.Add(10)
    With cpop
        .tag = C_TAG
        .BeginGroup = True
        .CAption = DicButtons("Inserir/Editar")
        .ToolTipText = "Listar"
    End With
    Call AddMenuButton(DicButtons("InsertProcedureHeader"), True, "InsertProcedureHeader", 12, cpop)
    Call AddMenuButton(DicButtons("insertErrorTreatment"), False, "insertErrorTreatment", 464, cpop)
    Call AddMenuButton(DicButtons("InsertLineNumber"), True, "InsertLineNumber", 9680, cpop)
    Call AddMenuButton(DicButtons("RemoveLineNumber"), False, "RemoveLineNumber", 4171, cpop) '66
    Call AddMenuButton(DicButtons("IndentVariables"), True, "IndentVariables", 123, cpop)
    '--------------------------------------------
    With cbBarTOOL
        'Strings
        Set cpop = .Controls.Add(10)
        With cpop
            .BeginGroup = True
            .tag = C_TAG
            .CAption = DicButtons("AuxText")
            .ToolTipText = "Textus"
        End With

        Call AddMenuButton(DicButtons("toUpperCase"), False, "toUpperCase", 311, cpop)
        Call AddMenuButton(DicButtons("toLowerCase"), False, "toLowerCase", 310, cpop)

        'CopyText
        Set cpopCopyText = .Controls.Add(10)
        With cpopCopyText
            .tag = C_TAG
            .CAption = DicButtons("CopyText")
            .ToolTipText = "CopyText"
        End With

        'PasteText
        Set cpopPasteText = .Controls.Add(10)
        With cpopPasteText
            .tag = C_TAG
            .CAption = DicButtons("PasteText")
            .ToolTipText = "PasteText"
        End With

        'CopyText e PasteText sendo criados os botões do menu
        For i = 1 To 10 ', 6766,6735
            settingPasteText = VBA.GetSetting(C_APPNAME, C_SECTION_CopyText, i)
            Call AddMenuButton(DicButtons("CopyTextDesc") & i, False, "CopyText", IIf(settingPasteText = "", 1132, 7992), cpopCopyText, i, , "CopyOrPaste")
            Call AddMenuButton(IIf(settingPasteText = "", "PasteText " & i, VBA.Left$(settingPasteText, 50) & IIf(VBA.Len(settingPasteText) > 49, "...", "")), _
                                False, "PasteText", 1, cpopPasteText, i, (settingPasteText <> ""), "CopyOrPaste")
        Next i

        Call AddMenuButton(DicButtons("CleanPasteText"), True, "CleanPasteText", 450, cpopPasteText, "limpa todos PasteTexts")
    End With


    '-------------------------------------------
    Call AddMenuButton(DicButtons("CheckVariablesNotUsedInProcedure"), True, "CheckVariablesNotUsedInProcedure", 202)
    Set cpop = cbBarTOOL.Controls.Add(10)
    With cpop
        .BeginGroup = False
        .tag = C_TAG
        .CAption = DicButtons("Listar")
        .ToolTipText = "Listar"
    End With

    Call AddMenuButton(DicButtons("GetFunctionAndSubNames"), False, "GetFunctionAndSubNames", 2045, cpop)
    Call AddMenuButton(DicButtons("GetFunctionAndSubNameAtual"), False, "GetFunctionAndSubNameAtual", 2046, cpop)
    '----------------------------------------------------------------
    Call AddMenuButton(DicButtons("CloseProjectExplorer"), True, "CloseProjectExplorer", 2477)
    Call AddMenuButton(DicButtons("CloseAllWindowsCodeModule"), False, "CloseAllWindowsCodeModule", 2477)
    Call AddMenuButton(DicButtons("Hook"), False, "Hook", 650)
    '-----------------------------------------------------------------------------
    Set cpop = cbBarTOOL.Controls.Add(10)
    With cpop
        .tag = C_TAG
        .BeginGroup = True
        .CAption = DicButtons("CoresEditor")
        .ToolTipText = "CoresEditor"
    End With
    Call AddMenuButton(DicButtons("Change_color_Dark_Theme"), True, "Change_color_Dark_Theme", 9534, cpop)
    Call AddMenuButton(DicButtons("Change_color_White_Theme"), False, "Change_color_White_Theme", 9535, cpop)
    '----------------------------------------------------------
    Call AddMenuButton(DicButtons("Atualizar_RVTool"), True, "Atualizar_RVTool", 37) '654
    '----------------------------------------------------------
    Call AddMenuButton(DicButtons("IndentarProcedure"), True, "IndentarProcedure", 1556)  '66
    '----------------------------------------------------------
    Call AddMenuButton(DicButtons("Change_Region"), True, "Change_Region", 5765) '66
    '-----------------------------------------------------------------
    Call AddMenuButton(DicButtons("aboutme"), True, "aboutme", 59) '66

End Sub


Sub AddMenuButton(ByVal CAption As String, _
                    BeginGroup As Boolean, _
                    OnACtion As String, _
                    FaceID As Long, _
                    Optional ByVal cbar As Object = Nothing, _
                    Optional ByVal DescriptionText As String = "", _
                    Optional Enabled As Boolean = True, _
                    Optional ToolTipText As String)
                    
                    
                    
    If cbar Is Nothing Then
        Set cbar = cbBarTOOL
    End If
    With cbar
        Set CmdBarItem = .Controls.Add
        With CmdBarItem
            '.Type = 1
            .FaceID = FaceID
            .CAption = CAption
            .BeginGroup = BeginGroup
            .OnACtion = OnACtion
            .ToolTipText = ToolTipText
            .tag = C_TAG
            .Enabled = Enabled
            .DescriptionText = DescriptionText
        End With
    End With
    
    Set MenuEvent = New CVBECommandHandler
    Set MenuEvent.EvtHandler = Application.VBE.Events.CommandBarEvents(CmdBarItem)
    EventHandlers.Add MenuEvent
    
End Sub

Sub DeleteMenuItems()
On Error GoTo f
'''''''''''''''''''''''''''''''''''''''''''''''''''''
' This procedure deletes all controls that have a
' tag of C_TAG.
'''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Ctrl As Office.CommandBarControl
    Set Ctrl = Application.VBE.CommandBars.FindControl(tag:=C_TAG)
    Do Until Ctrl Is Nothing
        Ctrl.Delete
        Set Ctrl = Application.VBE.CommandBars.FindControl(tag:=C_TAG)
    Loop
    
    Set Ctrl = Application.VBE.CommandBars.FindControl(tag:=C_RV_TOOLS_BAR)
    Do Until Ctrl Is Nothing
        Ctrl.Delete
        Set Ctrl = Application.VBE.CommandBars.FindControl(tag:=C_RV_TOOLS_BAR)
    Loop
    
    Set Ctrl = Application.VBE.CommandBars.FindControl(tag:="TECNUN")
    Do Until Ctrl Is Nothing
        Ctrl.Delete
        Set Ctrl = Application.VBE.CommandBars.FindControl(tag:=C_RV_TOOLS_BAR)
    Loop
f:
End Sub


Public Sub ChangeRegistry_AccessVBOM()

    'Made by Ronan Vico
    'helped by Rabaquim
    'helpde by Fernando
    Dim shl
    Dim Key As String
    Key = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Excel ecurity\AccessVBOM"
    Set shl = CreateObject("WScript.Shell")
     Call shl.RegWrite(Key, 1, "REG_DWORD")
End Sub

Public Sub MOSTRAR_ERRO(ByVal ERR_DESC As String, ByVal ERR_Number As String, ByVal Rotina As String)
    
End Sub



