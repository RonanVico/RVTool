Attribute VB_Name = "modCreateVBEMenuItems"
Option Explicit
'VERSaO  1.01.0
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Feito Por: Ronan Vico
'Descricao: Este módulo possui Rotinas para criacao do botao na Barra de Comandos do VBE (Visual Basic Editor)
'           é necessario toda vez que iniciar a aplicacao instanciar a barra novamente ,pois ela funciona com eventos
'           Também é possivel rodar manualmente a rotina InitVBRVTool.
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
Public Const C_SECTION_COPIAR As String = "COPIAR"
Public Const C_SECTION_COLAR As String = "COPIAR"



Private Enum ControlsType
        msoControlButton = 1
        msoControlDropdown = 3
        msoControlComboBox = 4
        msoControlPopup = 10
End Enum

Sub InitVBRVTool()
Dim cpop                               As CommandBarPopup
Dim cpopColar                          As CommandBarPopup
Dim cpopCopiar                         As CommandBarPopup
Dim i                                  As Long
Dim settingColar                       As String
Dim FaceID                             As Long


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

    Set cbBarTOOL = cmbar.FindControl(tag:=C_RV_TOOLS_BAR)
    Call AddMenuButton("Complete Code Snippe&t", True, "Snippets", 7581)
    '-------------------------------------------------------------------------------
    Set cpop = cbBarTOOL.Controls.Add(10)
    With cpop
        .tag = C_TAG
        .BeginGroup = True
        .CAption = "&Inserir e Editar"
        .TooltipText = "Listar"
    End With
    
    Call AddMenuButton("Inserir &Cabecalho", True, "InserirCabecalhoNaProc", 12, cpop)
    Call AddMenuButton("Inserir &Error Treatment", False, "inserirTratamentoDeErro", 464, cpop)
    Call AddMenuButton("Inserir &Númeracao nas Linhas", True, "inserirNumeracaoDeLinha", 9680, cpop)
    Call AddMenuButton("Remover &Númeracao nas Linhas", False, "RetirarNumeraCaoDeLinhas", 4171, cpop) '66
    Call AddMenuButton("Identar &Variaveis", True, "IdentaVariaveis", 123, cpop)
    '--------------------------------------------
    With cbBarTOOL
        'Strings
        Set cpop = .Controls.Add(10)
        With cpop
            .BeginGroup = True
            .tag = C_TAG
            .CAption = "Aux Textos"
            .TooltipText = "Textus"
        End With

        Call AddMenuButton("Selection TO UPPER CASE aA", False, "toUpperCase", 311, cpop)
        Call AddMenuButton("Selection TO LOWER CASE Aa", False, "toLowerCase", 310, cpop)
        
        'Copiar
        Set cpopCopiar = .Controls.Add(10)
        With cpopCopiar
            .tag = C_TAG
            .CAption = "Copiar"
            .TooltipText = "Copiar"
        End With

        'Colar
        Set cpopColar = .Controls.Add(10)
        With cpopColar
            .tag = C_TAG
            .CAption = "Colar"
            .TooltipText = "Colar"
        End With

        'Copiar e colar sendo criados os botões do menu
        For i = 1 To 10 ', 6766,6735
            settingColar = VBA.GetSetting(C_APPNAME, C_SECTION_COPIAR, i)
            Call AddMenuButton("Copiar para area " & i, False, "Copiar", IIf(settingColar = "", 1132, 7992), cpopCopiar, i)
            Call AddMenuButton(IIf(settingColar = "", "Colar " & i, VBA.Left$(settingColar, 50) & IIf(VBA.Len(settingColar) > 49, "...", "")), _
                                False, "Colar", 1, cpopColar, i, (settingColar <> ""))
        Next i

        Call AddMenuButton("Limpar Tudo", True, "LimparColar", 450, cpopColar, "limpa todos colars")
    End With


    '-------------------------------------------
    Call AddMenuButton("Verificar Variaveis nao Utilizadas", True, "Verifica_Variaveis", 202)
    Set cpop = cbBarTOOL.Controls.Add(10)
    With cpop
        .BeginGroup = False
        .tag = C_TAG
        .CAption = "Listar Procedures"
        .TooltipText = "Listar"
    End With
    
    Call AddMenuButton("Imprimir TUDO", False, "GetFunctionAndSubNames", 2045, cpop)
    Call AddMenuButton("Imprimir Módulo Atual", False, "GetFunctionAndSubNameAtual", 2046, cpop)
    '----------------------------------------------------------------
    Call AddMenuButton("Fecha&r Project Explorer", True, "FecharProjectExplorer", 2477)
    Call AddMenuButton("Fechar All VBE &Windows", False, "FecharTodasJanelas", 2477)
    Call AddMenuButton("Desbloquear All VBE's", False, "Hook", 650)
    '-----------------------------------------------------------------------------
    Set cpop = cbBarTOOL.Controls.Add(10)
    With cpop
        .tag = C_TAG
        .BeginGroup = True
        .CAption = "&Alterar Cores Editor"
        .TooltipText = "Listar"
    End With
    Call AddMenuButton("DARK THEME", True, "Change_color_Dark_Theme", 9534, cpop)
    Call AddMenuButton("WHITE THEME DEFAULT", False, "Change_color_White_Theme", 9535, cpop)
    '----------------------------------------------------------
    Call AddMenuButton("Atualizar RV_TOOLS", False, "Atualizar_RVTool", 654)
    '-----------------------------------------------------------------
    Call AddMenuButton("About Creator", True, "aboutme", 59) '66
    '--
    'Call AddMenuButton("DEV", True, "dev", 1556) '66
    
    
End Sub

Sub AddMenuButton(ByVal CAption As String, _
                    BeginGroup As Boolean, _
                    OnACtion As String, _
                    FaceID As Long, _
                    Optional ByVal cbar As Object = Nothing, _
                    Optional ByVal DescriptionText As String = "", _
                    Optional Enabled As Boolean = True)
                    
                    
                    
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
            '.OnAction = "'" & ThisWorkbook.Name & "'!Procedure_One"
            '.OnACtion = "'" & ThisWorkbook.Name & "'!" & OnACtion
            .OnACtion = OnACtion
            .TooltipText = OnACtion
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
    Key = "HKEY_CURRENT_USER oftware\Microsoft\Office\" & Application.Version & "\Excel ecurity\AccessVBOM"
    Set shl = CreateObject("WScript.Shell")
 
     'Debug.Print shl.regRead(key)
'     Call shl.regWrite("HKEY_CURRENT_USER oftware\Microsoft\Office\16.0\Common\Graphics\DisableAnimations", 1, "REG_DWORD")
     Call shl.RegWrite(Key, 1, "REG_DWORD")
End Sub





Public Sub MOSTRAR_ERRO(ByVal ERR_DESC As String, ByVal ERR_Number As String, ByVal Rotina As String)
    
End Sub

