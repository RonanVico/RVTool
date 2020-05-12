Attribute VB_Name = "mod_WindowsRegistry"

Option Explicit

'---------------------------------------------------------------------------------------
' Autor.....: RONAN VICO
' Contato...: ronanvico@hotmail.com.br - Empresa: Ronan Vico - Rotina: Public Sub CreateContextMenuForExcel()
' Data......: 5/12/2020 d m y
' Descricao.: Made By Ronan Vico
'---------------------------------------------------------------------------------------
Public Sub CreateContextMenuForExcel()

    Dim DicXls      As New Scripting.Dictionary
    Dim DicCommands As New Scripting.Dictionary
    Dim ext, command
    With DicXls
        .Add "xlsm", "xlsm"
        .Add "xlsx", "xlsx"
        .Add "xls", "xls"
        .Add "xla", "xla"
        .Add "xlam", "xlam"
        .Add "xlsb", "xlsb"
        .Add "csv", "csv"
    End With
    
    With DicCommands
        .Add "ThisWorkbook Open New Instance", "excel.exe ""%1"" /x"
        .Add "ThisWorkbook Open as Read Only", "excel.exe ""%1"" /r"
        .Add "ThisWorkbook Open as Template", "excel.exe ""%1"" /t"
        .Add "ThisWorkbook Open as SafeMode", "excel.exe ""%1"" /s"
        .Add "New Workbook  New Instance", "excel.exe /x"
        .Add "New Workbook  Safe Mode", "excel.exe /s /x"
        .Add "About Creator", "explorer ""https://www.linkedin.com/in/ronan-vico/"""
        .Add "WARNING! Kill All Excels (TaskKill) ", "taskkill /f /im excel.exe"
    End With
    
    For Each ext In DicXls
        On Error Resume Next
        For Each command In DicCommands
            Call CreateContextMenu(VBA.CStr(ext), VBA.CStr(command), VBA.CStr(DicCommands(command)))
        Next command
    Next ext
    
    If PT_BR Then
        MsgBox "Menu Criado Com Sucesso!", vbInformation, "Pronto! " & VERSAO & " - RVTools"
    Else
        MsgBox "Menu Created Sucefully!", vbInformation, "Ready! " & VERSAO & " - RVTools"
    End If
End Sub

Private Sub CreateContextMenu(sExt As String, CommandCaption As String, sCommand As String)
    On Error GoTo Error_Handler
    Dim sClass                As String
    Dim sRegKey               As String
    
    'Find associated class
    sClass = RegKeyRead("HKEY_CLASSES_ROOT\." & sExt & "\")
    '******************************************************************************************************
    ' Do Not change the name of the following key to ensure compatibility with future update of this tool!
    '******************************************************************************************************
    sRegKey = "HKEY_CURRENT_USER\Software\Classes\" & sClass & "\shell\RVTools\"
    If RegKeyExists(sRegKey) = False Then Call RegKeyCreate(sRegKey, "")
    sRegKey = "HKEY_CURRENT_USER\Software\Classes\" & sClass & "\shell\RVTools\shell\"
    If RegKeyExists(sRegKey) = False Then Call RegKeyCreate(sRegKey, "")
    'Add required strings to basic structure
    'MUIVerb - Actual name of the menu
    sRegKey = "HKEY_CURRENT_USER\Software\Classes\" & sClass & "\shell\RVTools\MUIVerb"
    If RegKeyExists(sRegKey) = False Then Call RegKeySave(sRegKey, "RVTools", "REG_SZ")
    'SubCommands
    sRegKey = "HKEY_CURRENT_USER\Software\Classes\" & sClass & "\shell\RVTools\SubCommands"
    If RegKeyExists(sRegKey) = False Then Call RegKeySave(sRegKey, "", "REG_SZ")

    If Not IsNull(CommandCaption) Then
        sRegKey = "HKEY_CURRENT_USER\Software\Classes\" & sClass & "\shell\RVTools\shell\" & CommandCaption & "\"
        Call RegKeyCreate(sRegKey, CommandCaption)
        sRegKey = "HKEY_CURRENT_USER\Software\Classes\" & sClass & "\shell\RVTools\shell\" & CommandCaption & "\command\"
        Call RegKeyCreate(sRegKey, sCommand)
    End If

Error_Handler_Exit:
    On Error Resume Next
    Exit Sub

Error_Handler:
    MsgBox "The following error has occured" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Sub

'Self-Healing oWS Variable -> WScript.Shell
'   The property will automatically create the oWS variable if it does not exist whenever the variable is used.
'   No need to Dim or Set.
'   Even if you Set it to Nothing, the next time you use it, it will automatically be recreated.
Property Get oWS() As Object
    On Error GoTo Err_Handler:
    Static WS                  As Object

    If WS.SpecialFolders("Desktop") = WS.SpecialFolders("Desktop") Then Set oWS = WS

Exit_Procedure:
    Exit Property

Err_Handler:
    Select Case Err.Number
        Case 91, 3284, 3265
            Set WS = CreateObject("WScript.Shell")
            Resume
        Case Else
            MsgBox Err.Number & ": " & Err.Description
            Resume Exit_Procedure
    End Select
    Resume
End Property

'reads the value for the registry key i_RegKey
'if the key cannot be found, the return value is ""
Private Function RegKeyRead(i_RegKey As String) As String
'    Dim oWS                  As Object

    On Error Resume Next
    'access Windows scripting
'    Set oWS = CreateObject("WScript.Shell")
    'read key from registry
    RegKeyRead = oWS.RegRead(i_RegKey)
End Function


Private Sub RegKeyCreate(i_RegKey As String, _
               i_Value As String)
'    Dim oWS                  As Object

    'access Windows scripting
'    Set oWS = CreateObject("WScript.Shell")
    'write registry key
    oWS.RegWrite i_RegKey, i_Value
End Sub

'sets the registry key i_RegKey to the
'value i_Value with type i_Type
'if i_Type is omitted, the value will be saved as string
'if i_RegKey wasn't found, a new registry key will be created

' change REG_DWORD to the correct key type
Private Sub RegKeySave(i_RegKey As String, _
               i_Value As String, _
               Optional i_Type As String = "REG_DWORD")
'    Dim oWS                  As Object

    'access Windows scripting
'    Set oWS = CreateObject("WScript.Shell")
    'write registry key
    oWS.RegWrite i_RegKey, i_Value, i_Type
End Sub

'returns True if the registry key i_RegKey was found
'and False if not
Private Function RegKeyExists(i_RegKey As String) As Boolean
'    Dim oWS                  As Object

    On Error GoTo ErrorHandler
    'access Windows scripting
'    Set oWS = CreateObject("WScript.Shell")
    'try to read the registry key
    oWS.RegRead i_RegKey
    'key was found
    RegKeyExists = True
    Exit Function

ErrorHandler:
    'key was not found
    RegKeyExists = False
End Function


'returns True if the registry key i_RegKey was found
'and False if not
Private Function RegKeyDelete(i_RegKey As String) As Boolean
'    Dim oWS                  As Object

    On Error GoTo ErrorHandler
    'access Windows scripting
'    Set oWS = CreateObject("WScript.Shell")
    'try to read the registry key
    oWS.RegDelete i_RegKey
    'key was found
    RegKeyDelete = True
    Exit Function

ErrorHandler:
    'key was not found
    RegKeyDelete False
End Function


  sSplit = VBA.Split(sSplit, "'")(0)
           If IsLinhaMatch(sSplit, "(End (Function|Sub|Property))") Then
                Call Application.VBE.ActiveCodePane.CodeModule.InsertLines _
                                (nLinha, _
                             PARAM_ERROR_HANDLER_DEFAULT)
                Call Application.VBE.ActiveCodePane.CodeModule.InsertLines _
                     (pInfo.ProcBodyLine + 1 + linQuebrada, _
                    "on error got