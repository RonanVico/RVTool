VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTOOLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Teste ASD
'VERSaƒO  1.01.6

'*MANUTENa‡a•ES / MELHORIAS*
'Criado por: Ronan Raphael Vico // ronanvico@hotmail.com // https://br.linkedin.com/in/ronan-vico
'   Descria§a£o:Ma³dulo utilizado para programar o IDE VBE, facilitando criaa§a£o de rotinas e manutena§a£o de ca³digos.
'       as rotinas sera£o utilizadas em botaµes programaveis na barra de comandos do VBE dentro do EXCEL
'
'
'
' Ca³digos utilizados como ajuda e fornecedores
' 1- GetProcedureDeclaration  e ProcedureInfo - Made By CPearson
'----------------------------------------------------------------------------------------------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As LongPtr, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As LongPtr, ByVal lpfnCB As LongPtr) As Long
#Else
     Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If


Private Enum ProcScope
    ScopePrivate = 1
    ScopePublic = 2
    ScopeFriend = 3
    ScopeDefault = 4
End Enum
 
Private Enum LineSplits
    LineSplitRemove = 0
    LineSplitKeep = 1
    LineSplitConvert = 2
End Enum
 
Private Type ProcInfo
    ProcName                As String
    procKind                As VBIDE.vbext_ProcKind
    ProcStartLine           As Long
    ProcBodyLine            As Long
    ProcCountLines          As Long
    ProcScope               As ProcScope
    ProcDeclaration         As String
End Type
 

 
Private pInfo           As ProcInfo
       
       
        
'-------- PARAMETROS
'-------- ABAIXO ESTaƒO OS PARAMETROS UTILIZADOS PARA PROGRAMAa‡aƒO
Private Const QUEBRA_DE_LINHA                As String = "_VBNEWLINE!"
Private Const tagVarInit                     As String = "[V@_"
Private Const tagVarEnd                      As String = "@]"


'---------------------------------------------------------------------------------------
' Modulo....: cTOOLS \ CodeModule
' Rotina....: Private Property Get PARAM_HEADER_DEFAULT() As String
' Autor.....: RONAN VICO
' Contato...: RONANVICO@hotmail.com
' Data......: 14/05/2019
' Descria§a£o.: Retorna o cabea§alho
'---------------------------------------------------------------------------------------
Private Property Get PARAM_HEADER_DEFAULT() As String
pInfo = ProcedureInfo(ActiveProcedure, Application.VBE.ActiveCodePane.CodeModule, pInfo.procKind)
PARAM_HEADER_DEFAULT = "'---------------------------------------------------------------------------------------" & _
    vbNewLine & "' Autor.....: " & PARAM_PROGRAMADOR & _
    vbNewLine & "' Contato...: " & PARAM_PROGRAMADOR_MAIL & " - Empresa: " & PARAM_EMPRESA & " - Rotina: " & VBA.Replace(pInfo.ProcDeclaration, "_" & vbNewLine, "") & _
    vbNewLine & "' Data......: " & VBA.CStr(VBA.Date) & _
    vbNewLine & "' Descria§a£o.: " & _
    vbNewLine & "'---------------------------------------------------------------------------------------"
'PARAM_HEADER_DEFAULT = "'---------------------------------------------------------------------------------------" & _
    vbNewLine & "' Modulo....: " & Application.VBE.ActiveCodePane.CodeModule & " \ " & VBA.TypeName(Application.VBE.ActiveCodePane.CodeModule) & _
    vbNewLine & "' Rotina....: " & VBA.Replace(pInfo.ProcDeclaration, "_" & vbNewLine, "") & _
    vbNewLine & "' Autor.....: " & PARAM_PROGRAMADOR & _
    vbNewLine & "' Contato...: " & PARAM_PROGRAMADOR_MAIL & _
    vbNewLine & "' Data......: " & VBA.CStr(VBA.Date) & _
    vbNewLine & "' Empresa...: " & PARAM_EMPRESA & _
    vbNewLine & "' Descria§a£o.: " & _
    vbNewLine & "'---------------------------------------------------------------------------------------"
End Property
 
Public Property Get This() As cTOOLS
    Set This = Me
End Property

 
Private Property Get PARAM_ERROR_HANDLER_DEFAULT() As String
     PARAM_ERROR_HANDLER_DEFAULT = "Fim:" & vbNewLine _
                            & "     Exit " & TypeProcedure & vbNewLine _
                            & PARAM_ERROR_HANDLER_NAME & ":" & vbNewLine _
                            & "     Call MOSTRAR_ERRO(Err.Number, Err.Description, """ & ActiveProcedure & "()"")" & vbNewLine _
                            & "     Goto Fim"
   ' PARAM_ERROR_HANDLER_DEFAULT = PARAM_ERROR_HANDLER_NAME & ":" & vbNewLine _
        & VBA.String(2, PARAM_CHAR_IDENTAa‡aƒO) & "select case err.number " & vbNewLine _
            & VBA.String(4, PARAM_CHAR_IDENTAa‡aƒO) & "case 0 " & vbNewLine _
            & VBA.String(4, PARAM_CHAR_IDENTAa‡aƒO) & "case else " & vbNewLine _
                & VBA.String(6, PARAM_CHAR_IDENTAa‡aƒO) & "msgbox err.description  & "" "" & Err.number , vbCritical " & vbNewLine _
        & VBA.String(2, PARAM_CHAR_IDENTAa‡aƒO) & "end Select"
End Property
'/\--------PARAMETROS------/\------------/\------------/\--------------/\-------------/\

Public Sub addFromGuidVBEPRoject()
    On Error Resume Next
    'Muda o registro do windows pra liberar acesso ao VBE ;) WE ARE HACKERS
    Call ChangeRegistry_AccessVBOM
    'Adiciona biblioteca do VBE
    Call Application.VBE.ActiveVBProject.References.AddFromGuid("{0002E157-0000-0000-C000-000000000046}", 2, 0)
End Sub
 
'--------PARTE 1 DAS PROPRIEDADES -------------------
Private Property Get ActiveProcedure() As String
    ActiveProcedure = Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(ActiveStartCodeLine, pInfo.procKind)
End Property
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
 
'-------------parte 2 das propriedades ------------------------------------
'--- Properties para pegar as linhas e colunas ja selecionadas no codemodule
Private Property Get ActiveStartCodeLine() As Long
    Application.VBE.ActiveCodePane.GetSelection ActiveStartCodeLine, 0, 0, 0
End Property
Private Property Get ActiveStartCodeColumn() As Long
    Application.VBE.ActiveCodePane.GetSelection 0, ActiveStartCodeColumn, 0, 0
End Property
Private Property Get ActiveEndCodeLine() As Long
    Application.VBE.ActiveCodePane.GetSelection 0, 0, ActiveEndCodeLine, 0
End Property
Private Property Get ActiveEndCodeColumn() As Long
    Application.VBE.ActiveCodePane.GetSelection 0, 0, 0, ActiveEndCodeColumn
End Property
'/\------------/\------------/\--------------/\-------------/\---------------------
 
 
'-------------Parte 3 Funa§aµes que ira£o mudar o mundo do VBA ------------------------------------
Public Sub inserirTratamentoDeErro()
    Dim nLinha          As Long
    Dim slinha          As String
    Dim sSplit
    Dim linQuebrada     As Long
    On Error GoTo t
    'Verifica se esta¡ numa procedure
    If ActiveProcedure = "" Then Exit Sub
 
    pInfo = ProcedureInfo(ActiveProcedure, Application.VBE.ActiveCodePane.CodeModule, pInfo.procKind)
    
    'Verifica se a declaraa§a£o esta quebrada
    For nLinha = pInfo.ProcBodyLine To (pInfo.ProcCountLines + pInfo.ProcBodyLine)
         slinha = Application.VBE.ActiveCodePane.CodeModule.Lines(nLinha, 1)
         If VBA.Right(slinha, 1) = "_" Then
            linQuebrada = linQuebrada + 1
         Else
            Exit For
         End If
    Next nLinha
    
    
    'Insere na primeira linha o on error e na ultima o texto padra£o
    For nLinha = pInfo.ProcStartLine + pInfo.ProcCountLines - 1 To pInfo.ProcStartLine + 2 Step -1
        slinha = Application.VBE.ActiveCodePane.CodeModule.Lines(nLinha, 1)
        For Each sSplit In VBA.Split(slinha, ":")
            sSplit = VBA.Split(sSplit, "'")(0)
           If IsLinhaMatch(sSplit, "(End (Function|Sub|Property))") Then
                Call Application.VBE.ActiveCodePane.CodeModule.InsertLines _
                                (nLinha, _
                             PARAM_ERROR_HANDLER_DEFAULT)
                Call Application.VBE.ActiveCodePane.CodeModule.InsertLines _
                     (pInfo.ProcBodyLine + 1 + linQuebrada, _
                    "on error goto " & PARAM_ERROR_HANDLER_NAME)
                Exit Sub
            End If
        Next sSplit
    Next nLinha
   
    
    Exit Sub
    Resume
t:
    End
End Sub
 

 
 
 
'-------------Parte 3 Funa§aµes que ira£o mudar o mundo do VBA ------------------------------------
'---------------------------------------------------------------------------------------
' Modulo....: vbProjcet / Ma³dulo
' Rotina....: Public Sub inserirCabea§alhoNaProcedure()
' Autor.....: RONAN VICO
' Contato...: RONANVICO@hotmail.com
' Data......: 23/04/2019
' Descria§a£o.: Insere cabea§alho na procedure (Utilizei ela mesmo para criar esse cabea§alho aqui ;)
'---------------------------------------------------------------------------------------
Public Sub InserirCabecalhoNaProc()
    Dim nLinha          As Long
    Dim slinha          As String
    Dim sSplit
    On Error GoTo t
    'Verifica se esta¡ numa procedure
    If ActiveProcedure = "" Then Exit Sub
    'ebug.Print ActiveProcedure
    pInfo = ProcedureInfo(ActiveProcedure, Application.VBE.ActiveCodePane.CodeModule, pInfo.procKind)
    

    Call Application.VBE.ActiveCodePane.CodeModule.InsertLines _
                    (pInfo.ProcBodyLine, _
                    PARAM_HEADER_DEFAULT)
 

    Exit Sub
    Resume
t:
    End
End Sub
 
 
Public Sub Identar_Modulo(Optional modulo)
    On Error GoTo TratarErro
    Dim md As VBIDE.CodeModule
    If VBA.IsError(modulo) Then
        Set md = Application.VBE.ActiveCodePane.CodeModule
    ElseIf TypeOf modulo Is VBIDE.CodeModule Then
        Set md = modulo
    ElseIf VBA.VarType(modulo) = vbString Then
   
    End If
    Exit Sub
TratarErro:
End Sub
 

Public Function IsLinhaMatch(ByVal Linha As String, ParamArray Padroes() As Variant) As Boolean
 Dim resultado           As Boolean
 Dim Contador            As Byte
 Dim regExp              As Object
 On Error GoTo TratarErro
                                          'New VBScript_RegExp_55.RegExp
    If regExp Is Nothing Then Set regExp = VBA.CreateObject("VBScript.RegExp")
    With regExp
        'Padroes = TFW_AuxArray.Acertar_Array_Parametros(Padroes)
        For Contador = 0 To UBound(Padroes) Step 1
            If Not Padroes(Contador) = VBA.vbNullString Then
                .Pattern = Padroes(Contador)
                If .test(Linha) Then
                    resultado = True
                    Exit For
                End If
            End If
        Next Contador
    End With
    IsLinhaMatch = resultado
 Exit Function
TratarErro:
    'Call TFW_Excecoes.tratarerro(VBA.Err.Description, VBA.Err.Number, "TFW_AuxRegex.IsLinhaMatch()", Erl)
 End Function
 

 Private Function TypeProcedure() As String
    If VBA.InStr(1, pInfo.ProcDeclaration, "Function ", vbBinaryCompare) <> 0 Then
        TypeProcedure = " Function "
    ElseIf VBA.InStr(1, pInfo.ProcDeclaration, "Sub ", vbBinaryCompare) <> 0 Then
        TypeProcedure = " Sub "
    ElseIf VBA.InStr(1, pInfo.ProcDeclaration, "Property ", vbBinaryCompare) <> 0 Then
        TypeProcedure = " Property "
    End If
End Function

 
 
Public Sub ChangeRegistry_AccessVBOM()
    'Made by Ronan Vico
    'helped by Rabaquim
    'helpde by Fernando
    Dim shl
    Dim Key As String
    Key = "HKEY_CURRENT_USER oftware\Microsoft\Office\" & Application.Version & "\Excel ecurity\AccessVBOM"
    Set shl = CreateObject("WScript.Shell")
 
     'Debug.Print shl.regRead(key)
     Call shl.RegWrite(Key, 1, "REG_DWORD")
End Sub
 
Public Sub IdentaVariaveis(Optional Byval PROC as string)
    On Error GoTo f
    'Cara , na£o vou explicar essa funa§a£o porque nem eu sei oq eu fiz , _
                essa funa§a£o segue as leis de software internacional e de Deus , pois sa³ ele _
                deve saber como isso funciona
 
 'Dim Proc                As ProcInfo
 Dim Linha               As String
 Dim linhaFormatada      As String
 Dim linhas              As Variant
 Dim contLine            As Long
 Dim textoProc           As String
 Dim ArrVars             As Variant
 Dim TAntComment         As String
 Dim TPosComment         As String
 Dim contSubLines        As Long
 Dim subLines            As Variant
 Dim subLine             As String
 Dim DimLines            As Variant
 Dim contDimLines        As Long
 Dim DimLine             As Variant
 Dim ProcBodyStart       As Long
 Dim NovoTexto           As String
 Dim contVar             As Long
 Dim textJoin
 
   if PROC = vbnullstring then 
      If ActiveProcedure = "" Then
          Exit Sub
      End If
      PROC = activeprocedure
    end if 
                          
                          
    textoProc = PegarProcedureSemQuebraDeLinha(PROC)
 
    linhas = VBA.Split(textoProc, vbNewLine)
 
    For contLine = 0 To UBound(linhas)
        Linha = linhas(contLine)
        linhaFormatada = formataTexto(Linha, ArrVars)
        If PosComentario(Linha) <> 0 Then
            TAntComment = VBA.Left(linhaFormatada, VBA.InStr(linhaFormatada, "'") - 1)
            TPosComment = VBA.Mid(linhaFormatada, VBA.InStr(linhaFormatada, "'"))
        Else
            TAntComment = linhaFormatada
            TPosComment = ""
        End If
 
        subLines = VBA.Split(TAntComment, ":")
 
        For contSubLines = 0 To UBound(subLines)
            subLine = VBA.Trim(subLines(contSubLines))
 
            If VBA.Left((subLine), 4) = "Dim " Then
                subLine = VBA.Replace(subLine, ",", vbNewLine & "Dim ")
                subLine = SingleSpace(subLine)
                DimLines = VBA.Split(subLine, vbNewLine)
                For contDimLines = 0 To UBound(DimLines)
                    DimLine = DimLines(contDimLines)
                    'Colocando identaa§a£o das variaveis
                    If VBA.InStr(DimLine, " As ") = 0 Then
                        DimLine = DimLine & VBA.Strings.Space$(PARAM_TABULACAO_VARIAVEIS - (VBA.Len(DimLine) - 3)) & " As Variant"
                    ElseIf contDimLines = 0 Then
                        'Ajuste pois dava pal se tivesse maior que a tabulaa§a£o
                        If PARAM_TABULACAO_VARIAVEIS > (VBA.InStr(DimLine, " As ") - 5) Then
                            DimLine = VBA.Left(DimLine, VBA.InStr(DimLine, " As ") - 1) & VBA.Strings.Space$(PARAM_TABULACAO_VARIAVEIS - (VBA.InStr(DimLine, " As ") - 5)) & VBA.Mid(DimLine, VBA.InStr(DimLine, " As ") + 1)
                        End If
                    Else
                        DimLine = VBA.Left(DimLine, VBA.InStr(DimLine, " As ") - 1) & VBA.Strings.Space$(PARAM_TABULACAO_VARIAVEIS + 1 - (VBA.InStr(DimLine, " As ") - 5)) & VBA.Mid(DimLine, VBA.InStr(DimLine, " As ") + 1)
                    End If
                    DimLines(contDimLines) = DimLine
                Next
                'Stop
                'If(UBound(DimLines) > 0, vbNewLine & " ", "") &
                subLines(contSubLines) = VBA.Join(DimLines, vbNewLine)
            End If
            'Stop
        Next contSubLines
        'Stop
 
 
        NovoTexto = VBA.Join(subLines, ":") & TPosComment & vbNewLine
 
        If VBA.Join(ArrVars) <> "" Then
            For contVar = 0 To UBound(ArrVars)
                NovoTexto = VBA.Replace(NovoTexto, tagVarInit & contVar & tagVarEnd, ArrVars(contVar))
            Next
        End If
 
        'a‰ necessario retirar um espaa§o em branco do canto esquerdo pois ele a© inserido sem qerer
        NovoTexto = VBA.Replace(NovoTexto, QUEBRA_DE_LINHA, " _" & VBA.Chr(13))
        If VBA.Left(NovoTexto, 2) = vbNewLine And VBA.Len(NovoTexto) <> 2 Then
            NovoTexto = VBA.Mid(NovoTexto, 3)
        End If
        
        'If VBA.Left(NovoTexto, 1) = " " Then
         '   NovoTexto = VBA.Mid(NovoTexto, 2)
        'End If
        linhas(contLine) = NovoTexto
    Next contLine
 
    If linhas(UBound(linhas)) = "" Or linhas(UBound(linhas)) = vbNewLine Then
         ReDim Preserve linhas(LBound(linhas) To UBound(linhas) - 1)
    End If
 
    
    With pInfo
        ProcBodyStart = .ProcBodyLine
        'Retira espaa§os em brancos , gambiarra mesmo to nem ai!
        textJoin = VBA.Join(linhas)
        textJoin = VBA.Replace(VBA.Replace(textJoin, vbNewLine & " ", vbNewLine), QUEBRA_DE_LINHA, " _" & VBA.Chr(13))
        Call Application.VBE.ActiveCodePane.CodeModule.DeleteLines(ProcBodyStart, .ProcCountLines - (ProcBodyStart - .ProcStartLine))
        Call Application.VBE.ActiveCodePane.CodeModule.InsertLines(ProcBodyStart, textJoin)
    End With
    Exit Sub
    
f:

    Stop
    Resume
 End Sub

Public Sub inserirNumeracaoDeLinha()
    Call ProcedureInfo(ActiveProcedure, Application.VBE.ActiveCodePane.CodeModule, pInfo.procKind)
    Call inserirNumeracaoDeLinhaNaProc(ActiveProcedure)
End Sub
Public Sub inserirNumeracaoDeLinhaNaProc(actProc As String)
On Error GoTo TError
    
    Dim textoProc As String
    Dim linhas
    Dim Linha As String
    Dim c As Long
    'Dim numCarac As Long
    Dim bESCREVER As Boolean
    
    Dim ProcBodyStart As Long
    Dim textJoin
    'Dim actProc       As String
    
    
    'actProc = ActiveProcedure
    
    
    
    Call RetirarNumeraCaoDeLinhasProc(actProc)
    
    textoProc = PegarProcedureSemQuebraDeLinha(actProc)
    linhas = VBA.Split(textoProc, vbNewLine)
    
    bESCREVER = True
    For c = LBound(linhas) + 1 To UBound(linhas)
        If linhas(c) Like "*Select Case*" Then bESCREVER = False
        If linhas(c) Like "*End Select*" Then bESCREVER = True
        If VBA.CStr(linhas(c)) Like "End Function" Or VBA.CStr(linhas(c)) Like "End Sub" Or VBA.CStr(linhas(c)) Like "End Property" Then Exit For
        If bESCREVER Then
            Linha = (PARAM_NUM_LINHA_INICIAL + (PARAM_NUM_LINHA_SOMAR * c - 1)) & " " & VBA.CStr(linhas(c))
            linhas(c) = Linha
        End If
    Next c
    
    With pInfo
        ProcBodyStart = .ProcBodyLine
        'Retira espaa§os em brancos , gambiarra mesmo to nem ai!
        textJoin = VBA.Join(linhas, vbNewLine)
        textJoin = VBA.Replace(VBA.Replace(textJoin, vbNewLine, vbNewLine), QUEBRA_DE_LINHA, " _ " & vbNewLine)
        Call Application.VBE.ActiveCodePane.CodeModule.DeleteLines(ProcBodyStart, .ProcCountLines - (ProcBodyStart - .ProcStartLine))
        Call Application.VBE.ActiveCodePane.CodeModule.InsertLines(ProcBodyStart, textJoin)
    End With
    
Fim:
     Exit Sub
     
TError:
     'Call MOSTRAR_ERRO(Err.Number, Err.Description, "inserirNumeracaoDeLinha()")
     GoTo Fim
     Resume
End Sub
 


Public Sub RetirarNumeraCaoDeLinhas()
    Call ProcedureInfo(ActiveProcedure, Application.VBE.ActiveCodePane.CodeModule, pInfo.procKind)
    Call RetirarNumeraCaoDeLinhasProc(ActiveProcedure)
End Sub

Public Sub RetirarNumeraCaoDeLinhasProc(actProc As String)
    Dim textoProc As String
    Dim linhas
    Dim Linha As String
    Dim c As Long
    Dim numCarac As Long
    Dim ProcBodyStart As Long
    Dim textJoin
    Dim paa§oca
    
    textoProc = PegarProcedureSemQuebraDeLinha(actProc)
    linhas = VBA.Split(textoProc, vbNewLine)
    For c = LBound(linhas) + 1 To UBound(linhas)
        Linha = VBA.CStr(linhas(c))
        If VBA.Len(Linha) > 0 Then
            For numCarac = 1 To VBA.Len(Linha)
                If Not VBA.IsNumeric(VBA.Mid$(Linha, numCarac, 1)) Then
                    Exit For
                Else
                    Mid(Linha, numCarac) = " "
                End If
            Next numCarac
        linhas(c) = Linha
        End If
    Next c
    
 With pInfo
        ProcBodyStart = .ProcBodyLine
        'Retira espaa§os em brancos , gambiarra mesmo to nem ai!
        textJoin = VBA.Join(linhas, vbNewLine)
        textJoin = VBA.Replace(VBA.Replace(textJoin, vbNewLine, vbNewLine), QUEBRA_DE_LINHA, " _ " & vbNewLine)
        Call Application.VBE.ActiveCodePane.CodeModule.DeleteLines(ProcBodyStart, .ProcCountLines - (ProcBodyStart - .ProcStartLine))
        Call Application.VBE.ActiveCodePane.CodeModule.InsertLines(ProcBodyStart, textJoin)
    End With
End Sub
 
Public Sub asdeasd()
    Dim a, b            As Long: a = "My:String": Dim c As Integer, d, e, f, G _
            , S 'COMMENTARIO BACANNA!""'""
'asd
   
End Sub


 
Public Function PegarProcedureSemQuebraDeLinha(ProcedureName As String) As String
    'Se tiver _ no final da linha de ca³digo , significa que a proxima linha pertence ao mesmo comando do VBA _
               Exemplo este comentario , a linha de baixo esta comentada de vido a esse Underline localizado /\ aqui _
               Logo , Iremos transformar tudo em "1 linha" para conseguirmos rodar funa§aµes de formataa§a£o de texto _
               e posteriormente plotar de volta as qebras de llinhas no seus devidos lugares !
 
 
 Dim proc                As ProcInfo
 Dim textAntComment      As String
 Dim textComment         As String
 Dim line                As String
 Dim texto               As String
 Dim i                   As Long
 
 
    proc = ProcedureInfo(ProcedureName, Application.VBE.ActiveCodePane.CodeModule, pInfo.procKind)
    With proc
        i = .ProcBodyLine
        While i <= (.ProcCountLines + .ProcStartLine) - 1
            line = VBA.RTrim(Application.VBE.ActiveCodePane.CodeModule.Lines(i, 1))
            While VBA.Right(line, 1) = "_"
                i = i + 1
                line = VBA.Left(line, VBA.Len(line) - 1) & QUEBRA_DE_LINHA & VBA.RTrim(Application.VBE.ActiveCodePane.CodeModule.Lines(i, 1))
            Wend
            'textAntComment = VBA.Mid(line, 1, PosComentario(line) - 1)
            'textComment = VBA.Mid(line, PosComentario(line))
            texto = texto & line & vbNewLine
            i = i + 1
        Wend
        'Debug.Print Application.VBE.ActiveCodePane.CodeModule.Lines(.ProcBodyLine + 1, .ProcCountLines - 1)
    End With
    PegarProcedureSemQuebraDeLinha = texto
 End Function
 

 
 
Public Function formataTexto(ByVal TextoOriginal As String, Optional ByRef MyArrVar) As String
    'Transforma todas as strings em Variaveis dentro de um array , assim podemos manipular _
            todo o texto sem medo de estar mechendo com dados dentro de string , por exemplo _
            uma string pode ser MyString = "OLA : ' "  , Logo os caracteres ":" e "'" , sa£o _
            importantes na nossa formataa§a£o de texto ,e na£o podemos consideralo na formataa§a£o, _
            por isso monto um array para posteriormente apa³s a formataa§a£o pegar as strings e jogar de volta.
            
On Error GoTo TratarErro

    Dim ArrVars()
    Dim contVar                 As Long
    Dim i                       As Long
    Dim y                       As Long
    Dim c                       As String
    Dim LenMax                  As Long
    Dim texto                   As String
    Dim tag                     As String
    Dim ValorVariavel           As String
   
   
    texto = TextoOriginal
   
    LenMax = VBA.Len(texto)
    'Debug.Print Texto
    While (i <= LenMax)
        i = i + 1
        c = VBA.Mid$(texto, i, 1)
        'Verifica se a© um caracter Aspas Dupla "
        If c = VBA.Chr(34) Then
            'Percorre as proximas letras ata© que a string seja fechada MyString= "abc" <Fim da string
            For y = i + 1 To VBA.Len(texto)
                If VBA.Mid$(texto, y, 1) = VBA.Chr(34) Then
                    If VBA.Mid$(texto, y + 1, 1) = VBA.Chr(34) Then
                        y = y + 1
                    Else
                        Exit For
                    End If
                End If
            Next y
           'Adcionando a variavel nova ao vetor
            ReDim Preserve ArrVars(0 To contVar)
            ValorVariavel = VBA.Mid(texto, i + 1, (y) - (i + 1))
            'Aspas duplas de comentario estavam sendo apagadass
            ArrVars(contVar) = VBA.Chr(34) & ValorVariavel & VBA.Chr(34)
            tag = tagVarInit & contVar & tagVarEnd
            texto = VBA.Mid(texto, 1, i - 1) & tag & VBA.Mid(texto, y + 1, VBA.Len(texto))
            contVar = contVar + 1
            i = i + VBA.Len(tag) - 1
            LenMax = VBA.Len(texto)
            'contVar = contVar + 1
            'VBA.Mid(TEXTO,I+1,Y-1) <- string
            'VBA.Mid(TEXTO,1,I) & VBA.Mid(TEXTO,Y,VBA.Len(TEXTO))
        End If
       
    Wend
   
    MyArrVar = ArrVars
    formataTexto = texto
    Exit Function
    Stop
   
TratarErro:
        Stop
        Resume
        Select Case Err.Number
                Case 0
                Case Else
                        MsgBox Err.Description & " " & Err.Number, vbCritical
        End Select
End Function
 
Public Function PosComentario(ByVal text As String) As Long
    Dim auxtexto        As String
    auxtexto = formataTexto(text)
    PosComentario = (VBA.InStr(1, auxtexto, "'", vbTextCompare))
End Function
  
'---------------------------------------------\/  --------------------------------------------\/'- --------------------------------------------\/  --------------------------------------------\/
'PARTE 4 --------------------------------------------\/ FUNCTIONS UTILIZADAS PARA PEGAR OS DADOS DO VBE
'---------------------------------------------\/  --------------------------------------------\/'- --------------------------------------------\/  --------------------------------------------\/
Private Function ProcedureInfo(ProcName As String, CodeMod As VBIDE.CodeModule, procKind As VBIDE.vbext_ProcKind) As ProcInfo
 Dim BodyLine            As Long
 Dim Declaration         As String
 Dim FirstLine           As String
 
    BodyLine = CodeMod.ProcStartLine(ProcName, pInfo.procKind)
    If BodyLine > 0 Then
        With CodeMod
            pInfo.ProcName = ProcName
            pInfo.procKind = procKind
            pInfo.ProcBodyLine = .ProcBodyLine(ProcName, procKind)
            pInfo.ProcCountLines = .ProcCountLines(ProcName, procKind)
            pInfo.ProcStartLine = .ProcStartLine(ProcName, procKind)
 
            FirstLine = .Lines(pInfo.ProcBodyLine, 1)
            If VBA.Strings.StrComp(VBA.Left(FirstLine, VBA.Len("Public")), "Public", vbBinaryCompare) = 0 Then
                pInfo.ProcScope = ScopePublic
            ElseIf VBA.Strings.StrComp(VBA.Left(FirstLine, VBA.Len("Private")), "Private", vbBinaryCompare) = 0 Then
                pInfo.ProcScope = ScopePrivate
            ElseIf VBA.Strings.StrComp(VBA.Left(FirstLine, VBA.Len("Friend")), "Friend", vbBinaryCompare) = 0 Then
                pInfo.ProcScope = ScopeFriend
            Else
                pInfo.ProcScope = ScopeDefault
            End If
            pInfo.ProcDeclaration = GetProcedureDeclaration(CodeMod, ProcName, LineSplitKeep)
        End With
    End If
 
    ProcedureInfo = pInfo
 
 End Function
 

 
 
Private Function GetProcedureDeclaration(CodeMod As VBIDE.CodeModule, _
    ProcName As String, _
    Optional LineSplitBehavior As LineSplits = LineSplitRemove)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetProcedureDeclaration
' This return the procedure declaration of ProcName in CodeMod. The LineSplitBehavior
' determines what to do with procedure declaration that span more than one line using
' the "_" line continuation character. If LineSplitBehavior is LineSplitRemove, the
' entire procedure declaration is converted to a single line of text. If
' LineSplitBehavior is LineSplitKeep the "_" characters are retained and the
' declaration is split with vbNewLine into multiple lines. If LineSplitBehavior is
' LineSplitConvert, the "_" characters are removed and replaced with vbNewLine.
' The function returns vbNullString if the procedure could not be found.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim LineNum As Long
    Dim S As String
    Dim Declaration As String
   
    On Error Resume Next
    LineNum = CodeMod.ProcBodyLine(ProcName, pInfo.procKind)
    If Err.Number <> 0 Then
        Exit Function
    End If
    S = CodeMod.Lines(LineNum, 1)
    Do While VBA.Right(S, 1) = "_"
        Select Case True
            Case LineSplitBehavior = LineSplitConvert
                S = VBA.Left(S, VBA.Len(S) - 1) & vbNewLine
            Case LineSplitBehavior = LineSplitKeep
                S = S & vbNewLine
            Case LineSplitBehavior = LineSplitRemove
                S = VBA.Left(S, VBA.Len(S) - 1) & " "
        End Select
        Declaration = Declaration & S
        LineNum = LineNum + 1
        S = CodeMod.Lines(LineNum, 1)
    Loop
    Declaration = SingleSpace(Declaration & S)
    GetProcedureDeclaration = Declaration
   
 
End Function
 
Private Function SingleSpace(ByVal text As String) As String
    Dim Pos As String
    Pos = InStr(1, text, VBA.Space(2), vbBinaryCompare)
    Do Until Pos = 0
        text = VBA.Replace(text, VBA.Space(2), VBA.Space(1))
        Pos = InStr(1, text, VBA.Space(2), vbBinaryCompare)
    Loop
    SingleSpace = VBA.Trim$(text)
End Function
 
Sub ShowProcedureInfo()
    Dim vbproj As VBIDE.VBProject
    Dim vbCOMP As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim CompName As String
    Dim ProcName As String
    Dim procKind As VBIDE.vbext_ProcKind
    Dim pInfo As ProcInfo
   
    CompName = "md_VBACrack"
    ProcName = "Hook"
    procKind = vbext_pk_Proc
   
    Set vbproj = ActiveWorkbook.VBProject
    Set vbCOMP = vbproj.VBComponents(CompName)
    Set CodeMod = vbCOMP.CodeModule
   
    pInfo = ProcedureInfo(ProcName, CodeMod, pInfo.procKind)
   
    Debug.Print "ProcName: " & pInfo.ProcName
    Debug.Print "ProcKind: " & CStr(pInfo.procKind)
    Debug.Print "ProcStartLine: " & CStr(pInfo.ProcStartLine)
    Debug.Print "ProcBodyLine: " & CStr(pInfo.ProcBodyLine)
    Debug.Print "ProcCountLines: " & CStr(pInfo.ProcCountLines)
    Debug.Print "ProcScope: " & CStr(pInfo.ProcScope)
    Debug.Print "ProcDeclaration: " & pInfo.ProcDeclaration
End Sub
 
 
'---------------------------------------------------------------------------------------
' Modulo....: cTOOLS \ CodeModule
' Rotina....: Public Function toUpperCase()
' Autor.....: RONAN VICO
' Contato...: RONANVICO@hotmail.com
' Data......: 30/04/2019
' Descria§a£o.: Deixa a selea§a£o do texto para maiaºscula
'---------------------------------------------------------------------------------------
Public Function toUpperCase()
 On Error GoTo TratarErro
 Dim i                   As Long
 Dim newText             As String
 Dim lineText            As String
 Dim sL                  As Long
 Dim eL                  As Long
 Dim sC                  As Long
 Dim eC                  As Long
 
 
    pInfo = ProcedureInfo(ActiveProcedure, Application.VBE.ActiveCodePane.CodeModule, pInfo.procKind)

    Call Application.VBE.ActiveCodePane.GetSelection(sL, sC, eL, eC)
    
  
        
         If sL = eL Then
            lineText = Application.VBE.ActiveCodePane.CodeModule.Lines(sL, 1)
            newText = VBA.Mid(lineText, 1, sC - 1) & VBA.UCase$(VBA.Mid(lineText, sC, eC - sC)) & VBA.Mid(lineText, eC)
            Call Application.VBE.ActiveCodePane.CodeModule.ReplaceLine(sL, newText)
         Else
           For i = sL To eL
                newText = ""
                lineText = Application.VBE.ActiveCodePane.CodeModule.Lines(i, 1)
                If i = sL Then
                     newText = VBA.Mid(lineText, 1, sC - 1) & VBA.UCase$(VBA.Mid(lineText, sC))
                ElseIf i = eL Then
                      newText = VBA.UCase$(VBA.Mid(lineText, 1, eC - 1)) & VBA.Mid(lineText, eC)
                Else
                    newText = VBA.UCase$(lineText)
                End If
                Call Application.VBE.ActiveCodePane.CodeModule.ReplaceLine(i, newText)
            Next i
         End If
        Call Application.VBE.ActiveCodePane.SetSelection(sL, sC, eL, eC)
TratarErro:
        Select Case Err.Number
                Case 0
                Case Else
                        MsgBox Err.Description & " " & Err.Number, vbCritical
        End Select
        
 End Function
 
 
Public Function toLowerCase()
 On Error GoTo TratarErro
 Dim i                   As Long
 Dim newText             As String
 Dim lineText            As String
 Dim sL                  As Long
 Dim eL                  As Long
 Dim sC                  As Long
 Dim eC                  As Long

 
    pInfo = ProcedureInfo(ActiveProcedure, Application.VBE.ActiveCodePane.CodeModule, pInfo.procKind)

    Call Application.VBE.ActiveCodePane.GetSelection(sL, sC, eL, eC)
    
     If sL = eL Then
        lineText = Application.VBE.ActiveCodePane.CodeModule.Lines(sL, 1)
        newText = VBA.Mid(lineText, 1, sC - 1) & VBA.LCase(VBA.Mid(lineText, sC, eC - sC)) & VBA.Mid(lineText, eC)
        Call Application.VBE.ActiveCodePane.CodeModule.ReplaceLine(sL, newText)
     Else
       For i = sL To eL
            newText = ""
            lineText = Application.VBE.ActiveCodePane.CodeModule.Lines(i, 1)
            If i = sL Then
                newText = VBA.Mid(lineText, 1, sC - 1) & VBA.LCase(VBA.Mid(lineText, sC))
            ElseIf i = eL Then
                newText = VBA.LCase(VBA.Mid(lineText, 1, eC - 1)) & VBA.Mid(lineText, eC)
            Else
                newText = VBA.LCase(lineText)
            End If
            Call Application.VBE.ActiveCodePane.CodeModule.ReplaceLine(i, newText)
        Next i
     End If
    
    
    Call Application.VBE.ActiveCodePane.SetSelection(sL, sC, eL, eC)
    
TratarErro:
        Select Case Err.Number
                Case 0
                Case Else
                        MsgBox Err.Description & " " & Err.Number, vbCritical
        End Select
        
 End Function
 

Public Function AboutMe()
    MsgBox "Feito por Ronan Vico", vbInformation, "RVTool"
End Function


Public Function Hook()
    Call modVBACrack.Hook
End Function

'---------------------------------------------------------------------------------------
' Modulo....: cTOOLS \ CodeModule
' Rotina....: Public Function PegarSelecao() As String
' Autor.....: RONAN VICO
' Contato...: RONANVICO@hotmail.com
' Data......: 21/05/2019
' Descria§a£o.: Pega o texto selecionado no codemodule e joga numa string
'---------------------------------------------------------------------------------------
Public Function PegarSelecao() As String
On Error GoTo TError
Dim sL&, sC&, eL&, eC&, i&
Dim S, lineText$, newText$

'pInfo = ProcedureInfo(ActiveProcedure, Application.VBE.ActiveCodePane.CodeModule, pInfo.procKind)
Call Application.VBE.ActiveCodePane.GetSelection(sL, sC, eL, eC)

 If sL = eL Then
    lineText = Application.VBE.ActiveCodePane.CodeModule.Lines(sL, 1)
    newText = VBA.Mid(lineText, sC, eC - sC)
 Else
   For i = sL To eL
        lineText = Application.VBE.ActiveCodePane.CodeModule.Lines(i, 1)
        If i = sL Then
            newText = (VBA.Mid(lineText, sC))
        ElseIf i = eL Then
            newText = newText & (VBA.Mid(lineText, 1, eC - 1)) '& VBA.Mid(lineText, eC)
        Else
            newText = newText & (lineText)
        End If
        newText = newText & VBA.Chr(13)
    Next i
    newText = VBA.Left$(newText, VBA.Len(newText) - 1)
 End If
 PegarSelecao = newText
Fim:
     Exit Function
TError:
    ' Call MOSTRAR_ERRO(Err.Number, Err.Description, "PegarSelecao()")
     GoTo Fim
End Function

'---------------------------------------------------------------------------------------
' Modulo....: cTOOLS \ CodeModule
' Rotina....: Public Sub Copiar(ByVal NumeroDaAreaDeTransferencia As Long)
' Autor.....: RONAN VICO
' Contato...: RONANVICO@hotmail.com
' Data......: 21/05/2019
' Descria§a£o.: Joga o texto selecionado no registro dos settings do vba para depois colar
'---------------------------------------------------------------------------------------
Public Sub Copiar(ByVal NumeroDaAreaDeTransferencia As Long)
On Error GoTo TError
    Call VBA.SaveSetting(C_APPNAME, C_SECTION_COPIAR, NumeroDaAreaDeTransferencia, PegarSelecao())
Fim:
     Exit Sub
     Resume
TError:
     GoTo Fim
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: cTOOLS \ CodeModule
' Rotina....: Public Sub Colar(ByVal NumeroDaAreaDeTransferencia As Long)
' Autor.....: RONAN VICO
' Contato...: RONANVICO@hotmail.com
' Data......: 21/05/2019
' Descria§a£o.: Joga o texto na onde o cursor estiver selecionado
'---------------------------------------------------------------------------------------
Public Sub Colar(ByVal NumeroDaAreaDeTransferencia As Long)
On Error GoTo TError
    Dim sL&, sC&, eL&, eC&
    Dim StringParaColar$
    Dim eLString$
    Dim lineText$, newText$
    Dim i&
    
    Call Application.VBE.ActiveCodePane.GetSelection(sL, sC, eL, eC)
    'Pega do registro o texto a colar
    StringParaColar = VBA.GetSetting(C_APPNAME, C_SECTION_COPIAR, NumeroDaAreaDeTransferencia)


    'pInfo = ProcedureInfo(ActiveProcedure, Application.VBE.ActiveCodePane.CodeModule, pInfo.procKind)
    
     If sL = eL Then
        lineText = Application.VBE.ActiveCodePane.CodeModule.Lines(sL, 1)
        newText = VBA.Mid(lineText, 1, sC - 1) & StringParaColar & VBA.Mid(lineText, eC)
        If sL > Application.VBE.ActiveCodePane.CodeModule.CountOfLines Then
            Call Application.VBE.ActiveCodePane.CodeModule.InsertLines(sL - 1, newText)
        Else
            Call Application.VBE.ActiveCodePane.CodeModule.ReplaceLine(sL, newText)
        End If
        
     Else
       For i = sL To eL
            newText = ""
            lineText = Application.VBE.ActiveCodePane.CodeModule.Lines(i, 1)
            If i = sL Then
                 newText = VBA.Mid(lineText, 1, sC - 1)
            ElseIf i = eL Then
                  newText = VBA.Mid(lineText, eC)
            Else
                newText = ""
            End If
            Call Application.VBE.ActiveCodePane.CodeModule.ReplaceLine(i, newText)
        Next i
        StringParaColar = Application.VBE.ActiveCodePane.CodeModule.Lines(sL, 1) & StringParaColar
        Call Application.VBE.ActiveCodePane.CodeModule.ReplaceLine(sL, StringParaColar)
     End If
    Call Application.VBE.ActiveCodePane.SetSelection(sL, sC, eL, eC)
    
Fim:
     Exit Sub
     Resume
TError:
     GoTo Fim
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: cTOOLS \ CodeModule
' Rotina....: Public Sub LimparColar()
' Autor.....: RONAN VICO
' Contato...: RONANVICO@hotmail.com
' Data......: 21/05/2019
' Descria§a£o.: 'Retira todos os colars da memoria
'---------------------------------------------------------------------------------------
Public Sub LimparColar()
On Error GoTo TError
    If MsgBox("Deseja mesmo Limpar todos os ""Colar""?", vbYesNo + vbExclamation, "Atenção") = vbNo Then Exit Sub
    Dim i&
    
    For i = 1 To 10
        Call VBA.SaveSetting(C_APPNAME, C_SECTION_COPIAR, i, "")
    Next i
    
Fim:
     Exit Sub
TError:
     Debug.Print "erro ao rodar limparColar"
     GoTo Fim
End Sub


Public Function teste2() As cTOOLS
    Set teste2 = Me
End Function

Public Function teste3() As cTOOLS
    Set teste3 = Me
End Function



Public Sub GetFunctionAndSubNames()
On Error GoTo TError
    Dim item            As Variant

    For Each item In Application.VBE.ActiveVBProject.VBComponents
        'If ComponentTypeToString(vbext_ct_StdModule) = "Code Module" Then
            ListProcedures item.Name, True
            'Debug.Print item.CodeModule.Lines(1, item.CodeModule.CountOfLines)
        'End If
    Next item
Fim:
     Exit Sub
     
TError:
     'Call MOSTRAR_ERRO(Err.Number, Err.Description, "GetFunctionAndSubNames()")
     GoTo Fim
End Sub
 

Public Sub GetFunctionAndSubNameATual()
    ListProcedures Application.VBE.ActiveCodePane.CodeModule, True
End Sub


Private Sub ListProcedures(strName As String, Optional blnWithParentInfo = False)
On Error GoTo TError
 
    'Microsoft Visual Basic for Applications Extensibility 5.3 library
 
    Dim vbproj          As VBIDE.VBProject
    Dim vbCOMP          As VBIDE.VBComponent
    Dim CodeMod         As VBIDE.CodeModule
    Dim LineNum         As Long
    Dim ProcName        As String
    Dim procKind        As VBIDE.vbext_ProcKind
    Dim strSubsInfo     As String
    
    Set vbproj = Application.VBE.ActiveVBProject
    Set vbCOMP = vbproj.VBComponents(strName)
    Set CodeMod = vbCOMP.CodeModule
 
    With CodeMod
        LineNum = .CountOfDeclarationLines + 1
        
        Do Until LineNum >= .CountOfLines
            ProcName = .ProcOfLine(LineNum, procKind)
 
            If blnWithParentInfo Then
                strSubsInfo = strSubsInfo & IIf(strSubsInfo = vbNullString, vbNullString, vbCrLf) & strName & "." & ProcName
            Else
                strSubsInfo = strSubsInfo & IIf(strSubsInfo = vbNullString, vbNullString, vbCrLf) & ProcName
            End If
 
            LineNum = .ProcStartLine(ProcName, procKind) + .ProcCountLines(ProcName, procKind) + 1
        Loop
    End With
    
    If strSubsInfo <> vbNullString Then Debug.Print strSubsInfo
Fim:
     Exit Sub
TError:
     'Call MOSTRAR_ERRO(Err.Number, Err.Description, "ListProcedures()")
     Stop
     GoTo Fim
End Sub




Public Sub FecharTodasJanelas()
    Dim i
    Dim w
    
    For i = 1 To Application.VBE.CodePanes.Count
        Application.VBE.CodePanes(1).Window.Close
    Next i

    On Error Resume Next
    For Each w In Application.VBE.Windows
        If w.Type = vbext_wt_CodeWindow Then w.Close
        If w.Type = vbext_wt_Designer Then w.Close
    Next

End Sub

Public Sub FecharProjectExplorer()
    Dim w As VBIDE.Window
    Dim i As Long
    
    On Error Resume Next
    
    For Each w In Application.VBE.Windows
        If w.Type = vbext_wt_ProjectWindow Or w.Type = vbext_wt_PropertyWindow Then w.Close
    Next

End Sub




Public Sub Verifica_Variaveis()
On Error GoTo Fim
If ActiveProcedure = "" Then
    Exit Sub
End If
    
On Error GoTo TError
    Dim textoProc   As String
    Dim linhas
    Dim Linha       As String
    Dim LinhaVerificacaoSeFoiUsada       As String
    Dim contLine    As Long
    Dim contLine2    As Long
    Dim ArrAux
    Dim AntComment
    Dim LinhaSemComentario As String
    Dim ArrVars
    Dim ArrVarsPlus
    Dim LinhaDeDeclaraa§aµes
    Dim VarDeclaration
    Dim strRegexPatter      As String
    Dim DicVariaveis As New Scripting.Dictionary
    Dim VariaveisNaoUsadas
    Dim contVarsNaoUsadas As Long
    
    
    textoProc = PegarProcedureSemQuebraDeLinha(ActiveProcedure)
    'Stop
    
    linhas = VBA.Split(textoProc, vbNewLine)
    For contLine = 0 To UBound(linhas)
        Linha = linhas(contLine)
        Linha = formataTexto(Linha, ArrAux)
        
        If Linha = vbNullString Then Linha = " "
        LinhaSemComentario = VBA.Split(Linha, "'")(0)

        For Each AntComment In VBA.Split(LinhaSemComentario, ":")
            If IsLinhaMatch(AntComment, " Dim ") Or VBA.Left(AntComment, 4) = "Dim " Then
                If VBA.Left(AntComment, 4) = "Dim " Then AntComment = " " & AntComment
                'Stop
                ArrVars = VBA.Split(AntComment, " Dim ")
                LinhaDeDeclaraa§aµes = ArrVars(1)
                ArrVars = VBA.Split(LinhaDeDeclaraa§aµes, ",")
                For Each VarDeclaration In ArrVars
                    VarDeclaration = VBA.Split(VBA.Split(VarDeclaration, " As ")(0), "(")(0)
                    VarDeclaration = VBA.Replace(VBA.Replace(VarDeclaration, " ", ""), vbTab, "")
                    DicVariaveis.Add VarDeclaration, False
                    
                    'strRegexPatter = "(\\(v\\)|\\(v,|,v\\)|, v\\)|,v,|, v,| v = |^v = | v )"
                    'strRegexPatter = "(\\(v|\\(v\\)|\\(v,|, v\\)|,v\\)|^v = | v = |,v,|, v,| v = |^v | v )"
                    strRegexPatter = "(\(v\)|\(v,|, v\)|,v\)|^v = | v = |,v,|, v,| v = |^v | v |v\()"
                    strRegexPatter = VBA.Replace(strRegexPatter, "v", VarDeclaration)
                    'If VarDeclaration = "NovoTexto" Then Stop
                    
                    For contLine2 = 0 To UBound(linhas)
                        LinhaVerificacaoSeFoiUsada = linhas(contLine2)
                        If contLine2 <> contLine Then
                            'If VBA.InStr(LinhaVerificacaoSeFoiUsada, VarDeclaration) <> 0 Then Stop
                            If IsLinhaMatch(LinhaVerificacaoSeFoiUsada, strRegexPatter) Then
                                DicVariaveis(VarDeclaration) = True
                                Exit For
                            End If
                        End If
                    Next contLine2
                Next
            End If
        Next
    Next contLine
    

    
    For Each VariaveisNaoUsadas In DicVariaveis
        If DicVariaveis(VariaveisNaoUsadas) = False And VariaveisNaoUsadas <> "" Then
            Debug.Print "Variavel Na£o Usada: " & VariaveisNaoUsadas
            contVarsNaoUsadas = contVarsNaoUsadas + 1
        Else
            'Debug.Print "Variavel Usada: " & VariaveisNaoUsadas
        End If
    Next
    
    If contVarsNaoUsadas = 0 Then
        Debug.Print "Todas as variaveis esta£o sendo utilizadas , Paraba©ns mano!"
    End If
    'Stop
    
Fim:
     Exit Sub
     Resume
TError:
     'Call MOSTRAR_ERRO(Err.Number, Err.Description, "Verifica_Variaveis()")
     Stop
     GoTo Fim
End Sub



Public Function formataTexto2(ByVal TextoOriginal As String, Optional ByRef MyArrVar) As String
'Transforma todas as strings em Variaveis dentro de um array , assim podemos manipular _
            todo o texto sem medo de estar mechendo com dados dentro de string , por exemplo _
            uma string pode ser MyString = "OLA : ' "  , Logo os caracteres ":" e "'" , sa£o _
            importantes na nossa formataa§a£o de texto ,e na£o podemos consideralo na formataa§a£o, _
            por isso monto um array para posteriormente apa³s a formataa§a£o pegar as strings e jogar de volta.

 On Error GoTo TratarErro

 Dim ArrVars()           As Variant: Dim contVar            As Long: Dim i                  As Long:
Dim y                   As Long
Dim isString            As Boolean
Dim c                   As String
Dim Var                 As String
Dim LenMax              As Long
Dim PosEndQuote         As Long
Dim texto               As String
Dim tag                 As String
Dim ValorVariavel       As String

     texto = TextoOriginal

     LenMax = VBA.Len(texto)
     'Debug.Print Texto
     While (i <= LenMax)
         i = i + 1
         c = VBA.Mid$(texto, i, 1)
         'Verifica se a© um caracter Aspas Dupla ""
         If c = VBA.Chr(34) Then
             'Percorre as proximas letras ata© que a string seja fechada MyString= "abc" <Fim da string
             For y = i + 1 To VBA.Len(texto)
                 If VBA.Mid$(texto, y, 1) = VBA.Chr(34) Then
                     If VBA.Mid$(texto, y + 1, 1) = VBA.Chr(34) Then
                         y = y + 1
                     Else
                         Exit For
                     End If
                 End If
             Next y
            'Adcionando a variavel nova ao vetor
             ReDim Preserve ArrVars(0 To contVar)
             ValorVariavel = VBA.Mid(texto, i + 1, (y) - (i + 1))
             'Aspas duplas de comentario estavam sendo apagadass
             ArrVars(contVar) = VBA.Chr(34) & ValorVariavel & VBA.Chr(34)
             tag = tagVarInit & contVar & tagVarEnd
             texto = VBA.Mid(texto, 1, i - 1) & tag & VBA.Mid(texto, y + 1, VBA.Len(texto))
             contVar = contVar + 1
             i = i + VBA.Len(tag) - 1
             LenMax = VBA.Len(texto)
             'contVar = contVar + 1
             'VBA.Mid(TEXTO,I+1,Y-1) <- string
             'VBA.Mid(TEXTO,1,I) & VBA.Mid(TEXTO,Y,VBA.Len(TEXTO))
         End If

     Wend

     MyArrVar = ArrVars
     formataTexto2 = texto
     Exit Function
     Stop

TratarErro:
         Stop
         Resume
         Select Case Err.Number
                 Case 0
                 Case Else
                         MsgBox Err.Description & " " & Err.Number, vbCritical
         End Select
End Function


Public Function Exist_Procedure(ProcName As String, CodeMod As VBIDE.CodeModule, procKind As VBIDE.vbext_ProcKind) As Boolean
On Error GoTo f
    Dim p As ProcInfo
    p = ProcedureInfo(ProcName, CodeMod, procKind)
    Exist_Procedure = True
    Exit Function
f:
End Function

Public Sub Delete_Tudo_que_Possui()
    Call Deletar_Linhas_Inseridas_Que_Possuem
End Sub

Public Sub Deletar_Linhas_Inseridas_Que_Possuem(Optional ByVal texto As String = "'RV_INSERTED_BY_CODE")
On Error GoTo TError
    'ListProcedures2 Application.VBE.ActiveCodePane.CodeModule, True
    Dim vbproj          As VBIDE.VBProject
    Dim vbCOMP          As VBIDE.VBComponent
    Dim CodeMod         As VBIDE.CodeModule
    Dim procKind        As VBIDE.vbext_ProcKind
    Dim LineNum         As Long
    Dim ProcName    As String
    Dim nLinha      As Long
    Dim slinha As String
    Dim LineNumProc As Long

    Set vbproj = Application.VBE.ActiveVBProject

    If vbproj.Filename Like "*RVTool*" Then Exit Sub
    
    For Each vbCOMP In vbproj.VBComponents
        Set CodeMod = vbCOMP.CodeModule
        With CodeMod
            LineNum = .CountOfDeclarationLines + 1
            While LineNum <= .CountOfLines
                If VBA.UCase(.Lines(LineNum, 1)) Like "*" & VBA.UCase(texto) & "*" Then
                    Call .DeleteLines(LineNum, 1)
                    LineNum = LineNum - 1
                End If
                LineNum = LineNum + 1
            Wend
        End With
    Next
    
Fim:
     Exit Sub
     Resume
TError:
     Stop
     GoTo Fim
End Sub

Public Sub Change_color_White_Theme()
    CreateObject("WScript.Shell").RegWrite REG_BACK_COLOR, BACK_COLOR_WHITE_THEME, "REG_SZ"
    CreateObject("WScript.Shell").RegWrite Reg_FORE_COLOR, FORE_COLOR_WHITE_THEME, "REG_SZ"
    MsgBox "Pronto!, Reabra o Excel", vbInformation, "Reabra o Excel!"
End Sub

Public Sub Change_color_Dark_Theme()
    CreateObject("WScript.Shell").RegWrite REG_BACK_COLOR, BACK_COLOR_BLACK_THEME, "REG_SZ"
    CreateObject("WScript.Shell").RegWrite Reg_FORE_COLOR, FORE_COLOR_BLACK_THEME, "REG_SZ"
    MsgBox "Pronto!, Reabra o Excel", vbInformation, "Reabra o Excel!"
End Sub


Public Sub Atualizar_RVTool()

    If ThisWorkbook.VBProject.Protection = vbext_pp_locked Then
        MsgBox "Desproteja o VBE do RVTool!", vbExclamation, "ATENa‡aƒO!"
        Exit Sub
    End If
    
On Error GoTo TError
    Dim FSO As New Scripting.FileSystemObject
    Dim fl   As Scripting.File
    Dim fold As Scripting.Folder
    Dim vbproj As VBIDE.VBProject
    Dim vbCOMP As VBIDE.VBComponent
    Dim oAPP   As Object
    Dim objZipItems As Object
    Dim objFolderDestination As Object
    Dim PASTA                As String
    Dim i                    As Long
    Dim File   As String
    Dim Folder As String
    Const URL_PROJETO = "https://github.com/RonanVico/RVTool/archive/master.zip"
    
    
    Folder = ThisWorkbook.path & "\RVTool" & VBA.Format(VBA.Now(), "ddmmyyyyhhmmss") & "new"
    Call VBA.FileSystem.MkDir(Folder)
    Call URLDownloadToFile(0, URL_PROJETO, Folder & ".zip", 0, 0)
    File = Folder & ".zip\RVTool-master"
    With CreateObject("Shell.Application")
        .Namespace(Folder & "\").copyHere .Namespace(Folder & ".zip\").Items
    End With
    
    For i = 1 To 1000: VBA.DoEvents: Next i
    
    PASTA = Folder & "\RVTool-master"
    Set fold = FSO.GetFolder(PASTA)
    Set vbproj = ThisWorkbook.VBProject
    For Each fl In fold.Files
        If IsLinhaMatch(VBA.LCase(VBA.Split(fl.Name, ".")(UBound(VBA.Split(fl.Name, ".")))), "(bas|cls|md|frm)$") Then
            For Each vbCOMP In vbproj.VBComponents
                If VBA.UCase(vbCOMP.Name) = VBA.UCase(VBA.Split(fl.Name, ".")(0)) Then
                    vbCOMP.Name = vbCOMP.Name & "DELETAR__"
                    Exit For
                End If
            Next vbCOMP
            Call vbproj.VBComponents.Import(fl.path)
        End If
    Next fl

    For Each vbCOMP In vbproj.VBComponents
        If VBA.UCase(vbCOMP.Name) Like "*DELETAR__*" Then
            Call vbproj.VBComponents.Remove(vbCOMP)
        End If
    Next vbCOMP
    
    
    
Fim:
     On Error Resume Next
     Call TryTodeleteFolder
     Call Kill(Folder & ".zip")
     Exit Sub
     Resume
TError:
     'Call MOSTRAR_ERRO(Err.Number, Err.Description, "Atualizar()")
     Stop
     GoTo Fim
End Sub

Sub TryTodeleteFolder()
    On Error Resume Next
    Dim FSO As New Scripting.FileSystemObject
    Dim fl   As Scripting.File
    Dim fold As Scripting.Folder
    Dim fold2 As Scripting.Folder
    
    Set fold = FSO.GetFolder(ThisWorkbook.path)
    
    For Each fold2 In fold.SubFolders
        If VBA.LCase(fold2.Name) Like "*rvtool*new*" Then
            fold2.Delete
        End If
    Next fold2
End Sub



Public Sub Dev()
On Error GoTo TError
    'ListProcedures2 Application.VBE.ActiveCodePane.CodeModule, True
    Dim vbproj          As VBIDE.VBProject
    Dim vbCOMP          As VBIDE.VBComponent
    Dim CodeMod         As VBIDE.CodeModule
    Dim actWB           As Workbook
    Dim LineNum         As Long
    Dim TExtoCodeModule As String
    Dim texto           As String
    Dim ContadorDeLIn   As Long
    Dim bESCREVER       As Boolean
    Dim bSelectC        As Boolean
    Dim AuxNum          As String
    Dim i               As Long
    
    
    Const Texto_Do_Inicio As String = "Dim ID_PROC as double : ID_PROC = INSERIR_LOG(""@PROCEDURE"") "
    Const Texto_Do_Final  As String = "call UPDATE_HORARIO_SAIDA_LOG(ID_PROC)"
    
    Set actWB = ActiveWorkbook
    Set vbproj = Application.VBE.ActiveVBProject

    If vbproj.Filename Like "*RVTool*" Then Exit Sub
    

    VBA.DoEvents
    VBA.DoEvents
    VBA.DoEvents
    

    
    For Each vbCOMP In vbproj.VBComponents
        If vbCOMP.Name = "md_PainelResultados" Then Stop
        If vbCOMP.Name <> "zLOG_" And vbCOMP.Name <> "md_CONEXAO" Then _
            Set CodeMod = vbCOMP.CodeModule
            ContadorDeLIn = modParametros_USUARIO.PARAM_NUM_LINHA_INICIAL
           With CodeMod
                TExtoCodeModule = vbNullString
                LineNum = .CountOfDeclarationLines + 1
                If .CountOfDeclarationLines > 0 Then
                    TExtoCodeModule = .Lines(1, .CountOfDeclarationLines)
                    bSelectC = False
                    
                    Do Until LineNum > .CountOfLines
                        texto = VBA.Trim(formataTexto(.Lines(LineNum, 1)))
                        If texto = vbNullString Or VBA.Left(VBA.Trim(texto), 1) = "'" Then
                            TExtoCodeModule = TExtoCodeModule & vbCrLf & .Lines(LineNum, 1)
                        Else
                            
                            If ( _
                                VBA.Left(texto, 4) = "Sub " Or VBA.Left(texto, 9) = "Function " Or VBA.Left(texto, 9) = "Property " _
                                Or texto Like "End Sub*" Or texto Like "End Function*" Or texto Like "End Property*" _
                                Or IsLinhaMatch(texto, "(Private|Public|Friend) (Sub|Function|Property)") _
                                ) _
                            Then
                                bESCREVER = False
                            Else
                                If VBA.Right(.Lines(LineNum - 1, 1), 1) = "_" Then
                                    bESCREVER = False
                                Else
                                   bESCREVER = True
                                End If
                            End If
                            If texto Like "*Select Case*" Then bSelectC = True
                            If texto Like "*End Select*" Then bSelectC = False
                            If Not bSelectC And bESCREVER Then
                                AuxNum = .Lines(LineNum, 1)
                                If AuxNum <> vbNullString Then AuxNum = VBA.Split(AuxNum, " ")(0)
                                If AuxNum <> vbNullString Then AuxNum = VBA.Split(AuxNum, ":")(0)

                                If VBA.IsNumeric(AuxNum) Then
                                    TExtoCodeModule = TExtoCodeModule & vbCrLf & ContadorDeLIn & Mid(.Lines(LineNum, 1), VBA.Len(VBA.Split(.Lines(LineNum, 1), " ")(0)) + 1)
                                Else
                                    TExtoCodeModule = TExtoCodeModule & vbCrLf & ContadorDeLIn & " " & .Lines(LineNum, 1)
                                End If
                            Else
                                TExtoCodeModule = TExtoCodeModule & vbCrLf & .Lines(LineNum, 1)
                            End If
                            ContadorDeLIn = ContadorDeLIn + modParametros_USUARIO.PARAM_NUM_LINHA_SOMAR
                        End If
                        LineNum = LineNum + 1
                            For i = 1 To 5
                                VBA.DoEvents
                            Next
                    Loop
                    Debug.Print vbCOMP.Name
                    Call .DeleteLines(1, .CountOfLines)
                    For i = 1 To 100
                        VBA.DoEvents
                    Next
                    Call .InsertLines(1, TExtoCodeModule)
                    For i = 1 To 100
                        VBA.DoEvents
                    Next
                End If
            End With
    Next
    
Fim:
     Exit Sub
     Resume
TError:
      Stop
     GoTo Fim
End Sub

Public Sub Snippets()
On Error GoTo TError
Dim sL&, sC&, eL&, eC&, i&
Dim NEWsL&, NEWsC&, NEWeL&, NEWeC&
Dim SomarSC&, SomarEC&

Dim ProcInfo    As ProcInfo
Dim procKind    As VBIDE.vbext_ProcKind
Dim Linha       As String
Dim TextoVerificar As String
Dim CodigoParaInserir As String
Dim novaLInha        As String
Dim tamanhoSnippet As Long
Dim auxParam1

    'Pega a localizaa§a£o do cursor no ca³digo
    Call Application.VBE.ActiveCodePane.GetSelection(sL, sC, eL, eC)
    'PEga a linha inteira ata© a primeira posia§a£o do cursor do teclado
    Linha = VBA.Replace(VBA.Trim(VBA.Mid(Application.VBE.ActiveCodePane.CodeModule.Lines(sL, 1), 1, eC - 1)), " (", "(")
    'Retira os tabs por precaua§a£o
    Linha = VBA.Replace(Linha, vbTab, " ")
    If VBA.InStr(Linha, " ") <> 0 Then
        TextoVerificar = VBA.Mid(Linha, VBA.InStrRev(Linha, " ") + 1)
    Else
        TextoVerificar = Linha
    End If
    
    'retorna variavel da linha
    Linha = VBA.Replace(Application.VBE.ActiveCodePane.CodeModule.Lines(sL, 1), " (", "(")
    
    Select Case True
        Case VBA.LCase(TextoVerificar) Like "*fn"
            tamanhoSnippet = 2: SomarSC = 16: SomarEC = 8
            CodigoParaInserir = "Public Function MyFunc_" & VBA.Chr(CInt(VBA.Rnd(2) * 100)) & "() as Variant" & _
            vbNewLine & vbNewLine & "End Function"
        Case VBA.LCase(TextoVerificar) Like "*sub"
            tamanhoSnippet = 3: SomarSC = 11: SomarEC = 7
            CodigoParaInserir = "Public Sub MySUB_" & VBA.Chr(CInt(VBA.Rnd(2) * 100)) & "()" & _
            vbNewLine & vbNewLine & "End Sub"
        Case VBA.LCase(TextoVerificar) Like "*sub(*" Or VBA.LCase(TextoVerificar) Like "*sub (*"
            auxParam1 = VBA.Replace(VBA.Split(TextoVerificar, "sub(")(1), ")", "")
            tamanhoSnippet = 5 + VBA.Len(auxParam1) + IIf(VBA.LCase(TextoVerificar) Like "*sub (*", 1, 0)
            CodigoParaInserir = "Public Sub " & auxParam1 & "()" & _
            vbNewLine & vbNewLine & "End Sub"
        Case VBA.LCase(TextoVerificar) Like "*fn (*" Or VBA.LCase(TextoVerificar) Like "*fn(*"
            auxParam1 = VBA.Replace(VBA.Split(TextoVerificar, "fn(")(1), ")", "")
            tamanhoSnippet = VBA.Len(auxParam1) + 4 + IIf(TextoVerificar Like "*fn (*", 0, 1): SomarSC = 16 + VBA.Len(auxParam1) + 6: SomarEC = 7
            CodigoParaInserir = "Public Function " & auxParam1 & "() as Variant" & _
            vbNewLine & vbNewLine & "End Function"
        Case VBA.LCase(TextoVerificar) Like "*prop (*" Or VBA.LCase(TextoVerificar) Like "*prop(*"
            auxParam1 = VBA.Replace(VBA.Split(TextoVerificar, "prop(")(1), ")", "")
            tamanhoSnippet = VBA.Len(auxParam1) + 6 + IIf(TextoVerificar Like "*prop (*", 0, 2): SomarSC = 23 + VBA.Len(auxParam1): SomarEC = 8
            CodigoParaInserir = "Public Property Let " & auxParam1 & "(newValue)" & _
            vbNewLine & "m_" & auxParam1 & " = newValue " & _
            vbNewLine & vbNewLine & "End Property" & _
            vbNewLine & "Public Property Get " & auxParam1 & "() as Variant" & _
            vbNewLine & auxParam1 & " = m_" & auxParam1 & _
            vbNewLine & "End Property"
    End Select
    
    If CodigoParaInserir <> vbNullString Then
        CodigoParaInserir = CodigoParaInserir & vbNewLine
        'caiu em alguma regra , deve realizar procedimento de snippet
        If sC - 1 - tamanhoSnippet < 0 Then
            novaLInha = CodigoParaInserir & VBA.Mid(Linha, sC)
        Else
            novaLInha = VBA.Mid(Linha, 1, sC - 1 - tamanhoSnippet) & CodigoParaInserir & VBA.Mid(Linha, sC)
        End If
        
        Call Application.VBE.ActiveCodePane.CodeModule.ReplaceLine(sL, novaLInha)
        On Error Resume Next
        Call Application.VBE.ActiveCodePane.SetSelection(sL, sC + SomarSC - tamanhoSnippet, sL, sC + SomarEC + SomarSC - tamanhoSnippet)
        Call inserirTratamentoDeErro
        Call Application.VBE.ActiveCodePane.SetSelection(sL, sC + SomarSC - tamanhoSnippet, sL, sC + SomarEC + SomarSC - tamanhoSnippet)
    End If
    
Fim:
     Exit Sub
     Resume
TError:
     'Call MOSTRAR_ERRO(Err.Number, Err.Description, "Snippets()")
     Debug.Print Err.Description
     Stop
     GoTo Fim
End Sub



Sub Unzip1()
    Dim FSO As Object
    Dim oAPP As Object
    Dim Fname As Variant
    Dim FileNameFolder As Variant
    Dim DefPath As String
    Dim strDate As String

    Fname = Application.GetOpenFilename(filefilter:="Zip Files (*.zip), *.zip", _
                                        MultiSelect:=False)
    If Fname = False Then
        'Do nothing
    Else
        'Root folder for the new folder.
        'You can also use DefPath = "C:\Users\Ron\test\"
        DefPath = Application.DefaultFilePath
        If Right(DefPath, 1) <> "\" Then
            DefPath = DefPath & "\"
        End If

        'Create the folder name
        strDate = Format(Now, " dd-mm-yy h-mm-ss")
        FileNameFolder = DefPath & "MyUnzipFolder " & strDate & "\"

        'Make the normal folder in DefPath
        MkDir FileNameFolder

        'Extract the files into the newly created folder
      

        'If you want to extract only one file you can use this:
        'oApp.Namespace(FileNameFolder).CopyHere _
         'oApp.Namespace(Fname).items.Item("test.txt")

        MsgBox "You find the files here: " & FileNameFolder

        On Error Resume Next
        Set FSO = CreateObject("scripting.filesystemobject")
        FSO.DeleteFolder Environ("Temp") & "\Temporary Directory*", True
    End If
End Sub




--FEITO POR RONAN VICO 04/02/2020
-- NA LIVE DO GUIA DO EXCEL
Public Sub IdentarProcedure()
On Error GoTo TError
If ActiveProcedure = "" Then Exit Sub
Dim DicOpenIdentWords  As New Scripting.Dictionary
Dim arrWords           As Variant
Dim contWord           As Long
Dim lin                As Long
Dim arrLinhas          As Variant
Dim strLinha           As String
Dim key
Dim actProc            As String
Dim bSomarTAB          As Boolean
Dim bSubtrairTAB       As Boolean
Dim contQuantidadeIdentar   As Long
'
Dim vbProj             As VBIDE.VBProject
Dim vbCodeModule       As VBIDE.CodeModule
'
Dim NEW_CODE_TEXT      As String

Set vbCodeModule = Application.VBE.ActiveCodePane.CodeModule
actProc = ActiveProcedure

Const caracterDeIdentacao  As String = PARAM_CHAR_IDENTAÇÃO

    contQuantidadeIdentar = 0
    With DicOpenIdentWords
        .Add "Do |Do$", "(Loop)"
        .Add "For ", "(Next)"
        .Add "If .*Then$", "(Else$|ElseIf .*&Then$|End If$)"
        .Add "ElseIf ", "(Else$|ElseIf .*&Then$|End If$)"
        .Add "Else$", "(End If)"
        .Add "While ", "(Wend)"
        .Add "Select Case", "(End Select)"
        .Add "Case ", "(Case |End Select)"
        .Add "With ", "(End With$)"
        .Add "Public |Private |Friend ", "(End Sub|End Function|End Property)"
        .Add "Sub |Function |Property ", "(End Sub|End Function|End Property)"
    End With
    
    Dim TEXTO_DA_PROC As String
    Dim arrVars
    TEXTO_DA_PROC = formataTexto2(PegarProcedureSemQuebraDeLinha(ActiveProcedure()), arrVars)
    TEXTO_DA_PROC = VBA.Replace(PegarProcedureSemQuebraDeLinha(ActiveProcedure()), QUEBRA_DE_LINHA, "_" & vbCrLf)
    arrLinhas = VBA.Split(TEXTO_DA_PROC, vbCrLf)
    
    
    'Joga tudo para a esquerda para começar identar
    For lin = LBound(arrLinhas, 1) To UBound(arrLinhas, 1)
        'Problema com numeração de linhas
        strLinha = VBA.LTrim(arrLinhas(lin))
        If VBA.IsNumeric(VBA.Left(strLinha, 1)) Then
            strLinha = VBA.Mid(strLinha, VBA.InStr(strLinha, " "))
        End If
        arrLinhas(lin) = strLinha
    Next lin
    

    
    For lin = LBound(arrLinhas, 1) To UBound(arrLinhas, 1)
        'devido a numeração de linhas
        strLinha = VBA.RTrim(VBA.Split(IIf(arrLinhas(lin) = vbNullString, " ", arrLinhas(lin)), "'")(0))
        If VBA.IsNumeric(VBA.Left(strLinha, 1)) Then
            strLinha = VBA.Mid(strLinha, VBA.InStr(strLinha, " "))
        End If
        strLinha = IIf(strLinha = vbNullString, " ", strLinha)
        strLinha = IIf(VBA.Right$(strLinha, 1) = "_", VBA.Mid(strLinha, 1, VBA.Len(strLinha) - 1), strLinha)

        
        bSubtrairTAB = False
        bSomarTAB = False
        
        If VBA.IsArray(arrWords) Then
            'Verifica a Linha para Subtrair
            If IsLinhaMatch(strLinha, DicOpenIdentWords(arrWords(UBound(arrWords)))) Then
                bSubtrairTAB = True
                'Clausula especial para Select  Case , pois identação é diferente
                If arrWords(UBound(arrWords)) = "Case " And IsLinhaMatch(strLinha, "^(End Select)") Then
                     contQuantidadeIdentar = contQuantidadeIdentar - 1
                     'Pop the Stack
                    If UBound(arrWords) = 0 Then Set arrWords = Nothing Else ReDim Preserve arrWords(0 To UBound(arrWords) - 1)
                End If
                'Pop the Stack
                If UBound(arrWords) = 0 Then Set arrWords = Nothing Else ReDim Preserve arrWords(0 To UBound(arrWords) - 1)
            End If
        End If
        For Each key In DicOpenIdentWords
            'verifica linha para somar
            If IsLinhaMatch(strLinha, "^(" & key & ")") And Not IsLinhaMatch(strLinha, DicOpenIdentWords(key)) Then
                bSomarTAB = True
                'Push the Stack
                If VBA.IsArray(arrWords) Then
                    ReDim Preserve arrWords(0 To UBound(arrWords) + 1)
                Else
                    arrWords = VBA.Split("", "")
                End If
                
                arrWords(UBound(arrWords)) = key
                Exit For
            End If
            
        Next key
        
        If bSubtrairTAB Then contQuantidadeIdentar = contQuantidadeIdentar - 1
        
        strLinha = arrLinhas(lin)
        If VBA.IsNumeric(VBA.Left(strLinha, 1)) Then
            strLinha = VBA.Mid(strLinha, 1, VBA.InStr(strLinha, " ") - 1) & VBA.Strings.String(contQuantidadeIdentar, caracterDeIdentacao) & VBA.Mid(strLinha, VBA.InStr(strLinha, " "))
        Else
            arrLinhas(lin) = VBA.Strings.String(contQuantidadeIdentar, caracterDeIdentacao) & strLinha
        End If
        If bSomarTAB Then contQuantidadeIdentar = contQuantidadeIdentar + 1
    Next lin
    
    
    While arrLinhas(UBound(arrLinhas)) = vbNullString
        ReDim Preserve arrLinhas(0 To UBound(arrLinhas) - 1)
    Wend
    
    NEW_CODE_TEXT = VBA.Join(arrLinhas, vbCrLf)
    
    With pInfo
        Call vbCodeModule.DeleteLines(.ProcBodyLine, .ProcCountLines - (.ProcBodyLine - .ProcStartLine))
        Call vbCodeModule.InsertLines(.ProcBodyLine, NEW_CODE_TEXT)
    End With
    
    Call IdentaVariaveis(actProc)
    
Fim:
     Exit Sub
TError:
     'Call MOSTRAR_ERRO(Err.Number, Err.Description, "dev()")
     Stop
     Resume
     GoTo Fim
End Sub
                                
                                
                                
