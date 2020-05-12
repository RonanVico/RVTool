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
'VERSAO  1.02.1

'*MANUTENa√¢‚Ç¨¬°a√¢‚Ç¨¬¢ES / MELHORIAS*
'Criado por: Ronan Raphael Vico // ronanvico@hotmail.com // https://br.linkedin.com/in/ronan-vico
'   Descricao:Modulo utilizado para programar o IDE VBE, facilitando criacao de rotinas e manutencao de codigos.
'       as rotinas serao utilizadas em botoes programaveis na barra de comandos do VBE dentro do EXCEL
'
'
'
' Codigos utilizados como ajuda e fornecedores
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
'-------- ABAIXO ESTAO OS PARAMETROS UTILIZADOS PARA PROGRAMACAO
Private Const QUEBRA_DE_LINHA                As String = "_VBNEWLINE!"
Private Const tagVarInit                     As String = "[V@_"
Private Const tagVarEnd                      As String = "@]"


'---------------------------------------------------------------------------------------
' Modulo....: cTOOLS \ CodeModule
' Rotina....: Private Property Get PARAM_HEADER_DEFAULT() As String
' Autor.....: RONAN VICO
' Contato...: RONANVICO@hotmail.com
' Data......: 14/05/2019
' Descricao.: Retorna o cabecalho - Returns the proc header
'---------------------------------------------------------------------------------------
Private Property Get PARAM_HEADER_DEFAULT() As String
pInfo = ProcedureInfo(ActiveProcedure, Application.VBE.ActiveCodePane.CodeModule, pInfo.procKind)
PARAM_HEADER_DEFAULT = "'---------------------------------------------------------------------------------------" & _
    vbNewLine & "' Autor.....: " & PARAM_PROGRAMADOR & _
    vbNewLine & "' Contato...: " & PARAM_PROGRAMADOR_MAIL & " - Empresa: " & PARAM_EMPRESA & " - Rotina: " & VBA.Replace(pInfo.ProcDeclaration, "_" & vbNewLine, "") & _
    vbNewLine & "' Data......: " & VBA.CStr(VBA.Date) & _
    vbNewLine & "' Descricao.: " & _
    vbNewLine & "'---------------------------------------------------------------------------------------"
End Property
 
Public Property Get This() As cTOOLS
    Set This = Me
End Property

 
Private Property Get PARAM_ERROR_HANDLER_DEFAULT() As String
     PARAM_ERROR_HANDLER_DEFAULT = "Fim:" & vbNewLine _
                            & "     Exit " & TypeProcedure & vbNewLine _
                            & PARAM_ERROR_HANDLER_NAME & ":" & vbNewLine _
                            & "     call msgbox(err.Number & err.Description)" & vbNewLine _
                            & "     Goto Fim"
                            '& "     Call MOSTRAR_ERRO(Err.Number, Err.Description, """ & ActiveProcedure & "()"")" & vbNewLine _
                            & "     Goto Fim"
End Property
'/\--------PARAMETROS------/\------------/\------------/\--------------/\-------------/\

Public Sub addFromGuidVBEPRoject()
    On Error Resume Next
    '#PT Muda o registro do windows pra liberar acesso ao VBE ;) WE ARE HACKERS
    '#EN Change the registry of windows to give us access to VBE project
    Call ChangeRegistry_AccessVBOM
    '#PT Adiciona biblioteca do VBE
    '#EN Add VBProject Library
    Call Application.VBE.ActiveVBProject.References.AddFromGuid("{0002E157-0000-0000-C000-000000000046}", 2, 0)
End Sub
 
'--------PARTE 1 DAS PROPRIEDADES -------------------
Private Property Get ActiveProcedure() As String
    ActiveProcedure = Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(ActiveStartCodeLine, pInfo.procKind)
End Property
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
'-------------parte 2 das propriedades ------------------------------------
'--- #PT Properties para pegar as linhas e colunas ja selecionadas no codemodule
'--- #EN Properties to get Start and End Columns/Lines
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
 
 
'-------------Parte 3 Funcoes que irao mudar o mundo do VBA ------------------------------------

Public Sub InsertErrorTreatment() 'Inserir Tratamento de Erro
    Dim nLinha          As Long
    Dim slinha          As String
    Dim sSplit
    Dim linQuebrada     As Long
    On Error GoTo t
    'Verifica se ESTAO numa procedure
    'If not is in a procedure then exit sub
    If ActiveProcedure = "" Then Exit Sub
 
    pInfo = ProcedureInfo(ActiveProcedure, Application.VBE.ActiveCodePane.CodeModule, pInfo.procKind)
    
    '#pt Verifica se a declaracao esta quebrada
    '#en Check if Proc have Break lines with _ at end
    For nLinha = pInfo.ProcBodyLine To (pInfo.ProcCountLines + pInfo.ProcBodyLine)
         slinha = Application.VBE.ActiveCodePane.CodeModule.Lines(nLinha, 1)
         If VBA.Right(slinha, 1) = "_" Then
            linQuebrada = linQuebrada + 1
         Else
            Exit For
         End If
    Next nLinha
     
    '#pt Insere na primeira linha o on error e na ultima o texto padrao
    '#en Insert in the first line ON ERROR  GOTO , and the last line the Default Text
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
 
 
'-------------Parte 3 Funcoes que irao mudar o mundo do VBA ------------------------------------
'---------------------------------------------------------------------------------------
' Modulo....: vbProjcet / Modulo
' Rotina....: Public Sub InsertProcedureHeaderedure()
' Autor.....: RONAN VICO
' Contato...: RONANVICO@hotmail.com
' Data......: 23/04/2019
' Descricao.: Insere cabecalho na procedure (Utilizei ela mesmo para criar esse cabecalho aqui ;)'
' Description: Insert Procedure's Header
'---------------------------------------------------------------------------------------
'#PT InsertProcedureHeader
Public Sub InsertProcedureHeader()
    Dim nLinha          As Long
    Dim slinha          As String
    Dim sSplit
    On Error GoTo t
    'Verifica se ESTAO numa procedure
    'Check if is in a procedure
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
 
 
'---------------------------------------------------------------------------------------
' Autor.....: RONAN VICO
' Contato...: ronanvico@hotmail.com.br - Empresa: Ronan Vico - Rotina: Public Function IsLinhaMatch(ByVal Linha As String, ParamArray Padroes() As Variant) As Boolean
' Data......: 08/02/2020
' Descricao.: Check if a line Matchs a REGEX Pattern
'---------------------------------------------------------------------------------------
Public Function IsLinhaMatch(ByVal Linha As String, ParamArray Padroes() As Variant) As Boolean
 Dim resultado           As Boolean
 Dim Contador            As Byte
 Dim regExp              As Object
 On Error GoTo TratarErro
                                          'New VBScript_RegExp_55.RegExp
    If regExp Is Nothing Then Set regExp = VBA.CreateObject("VBScript.RegExp")
    With regExp
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
 End Function
 

 Private Function TypeProcedure() As String
    'pInfo.ProcDeclaration is Global
    If VBA.InStr(1, pInfo.ProcDeclaration, "Function ", vbBinaryCompare) <> 0 Then
        TypeProcedure = " Function "
    ElseIf VBA.InStr(1, pInfo.ProcDeclaration, "Sub ", vbBinaryCompare) <> 0 Then
        TypeProcedure = " Sub "
    ElseIf VBA.InStr(1, pInfo.ProcDeclaration, "Property ", vbBinaryCompare) <> 0 Then
        TypeProcedure = " Property "
    End If
End Function

'Indenta Variaveis
'---------------------------------------------------------------------------------------
' Autor.....: RONAN VICO
' Contato...: ronanvico@hotmail.com.br - Empresa: Ronan Vico - Rotina: Public Sub IndentVariables(Optional ByVal proc As String)
' Data......: 08/02/2020
' Descricao.:
' Transform This : Dim a,b,c as integer
'into this:
'dim a      as variant
'dim b      as variant
'dim c      as integer
'---------------------------------------------------------------------------------------
Public Sub IndentVariables(Optional ByVal proc As String)
    On Error GoTo f
    '#PT Cara , nao vou explicar essa funcao porque nem eu sei oq eu fiz , _
                essa funcao segue as leis de software internacional e de Deus , pois so ele _
                deve saber como isso funciona
                
    '#EN Dude , i wont explain this procedure , neither i or God knows what was done here
    
    
 'Dim Proc                As ProcInfo
 Dim Linha               As String
 Dim linhaFormatada      As String
 Dim linhas              As Variant
 Dim contLine            As Long
 Dim textoProc           As String
 Dim arrVars             As Variant
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
 
   If proc = vbNullString Then
      If ActiveProcedure = "" Then
          Exit Sub
      End If
      proc = ActiveProcedure
    End If
                          
                          
    textoProc = GetProcedureTextWithoutBreakLines(proc)
 
    linhas = VBA.Split(textoProc, vbNewLine)
 
    For contLine = 0 To UBound(linhas)
        Linha = linhas(contLine)
        linhaFormatada = formataTexto(Linha, arrVars)
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
                    '#pt Colocando Indentacao das variaveis
                    '#en Creating the "As Variant" when it do not have
                    If VBA.InStr(DimLine, " As ") = 0 Then
                        DimLine = DimLine & VBA.Strings.Space$(PARAM_TABULACAO_VARIAVEIS - (VBA.Len(DimLine) - 3)) & " As Variant"
                    ElseIf contDimLines = 0 Then
                        'Ajuste pois dava pal se tivesse maior que a tabulacao
                        If PARAM_TABULACAO_VARIAVEIS > (VBA.InStr(DimLine, " As ") - 5) Then
                            DimLine = VBA.Left(DimLine, VBA.InStr(DimLine, " As ") - 1) & VBA.Strings.Space$(PARAM_TABULACAO_VARIAVEIS - (VBA.InStr(DimLine, " As ") - 5)) & VBA.Mid(DimLine, VBA.InStr(DimLine, " As ") + 1)
                        End If
                    Else
                        DimLine = VBA.Left(DimLine, VBA.InStr(DimLine, " As ") - 1) & VBA.Strings.Space$(PARAM_TABULACAO_VARIAVEIS + 1 - (VBA.InStr(DimLine, " As ") - 5)) & VBA.Mid(DimLine, VBA.InStr(DimLine, " As ") + 1)
                    End If
                    DimLines(contDimLines) = DimLine
                Next
                subLines(contSubLines) = VBA.Join(DimLines, vbNewLine)
            End If
        Next contSubLines
        NovoTexto = VBA.Join(subLines, ":") & TPosComment & vbNewLine
        If VBA.Join(arrVars) <> "" Then
            For contVar = 0 To UBound(arrVars)
                NovoTexto = VBA.Replace(NovoTexto, tagVarInit & contVar & tagVarEnd, arrVars(contVar))
            Next
        End If
        '#PT E necessario retirar um espaco em branco do canto esquerdo pois ele e inserido sem qerer
        '#EN We need to skape 1 Space character in left side , pois its a missplaced
        
        NovoTexto = VBA.Replace(NovoTexto, QUEBRA_DE_LINHA, " _" & VBA.Chr(13))
        If VBA.Left(NovoTexto, 2) = vbNewLine And VBA.Len(NovoTexto) <> 2 Then
            NovoTexto = VBA.Mid(NovoTexto, 3)
        End If
        linhas(contLine) = NovoTexto
    Next contLine
 
    If linhas(UBound(linhas)) = "" Or linhas(UBound(linhas)) = vbNewLine Then
         ReDim Preserve linhas(LBound(linhas) To UBound(linhas) - 1)
    End If
 
    
    With pInfo
        ProcBodyStart = .ProcBodyLine
        '#PT Retira espacos em brancos , gambiarra mesmo to nem ai!
        '#EN Remove Blank spaces
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

'inserir NumeracaoDeLinha
Public Sub InsertLineNumber()
    Call ProcedureInfo(ActiveProcedure, Application.VBE.ActiveCodePane.CodeModule, pInfo.procKind)
    Call InsertLineNumberInProc(ActiveProcedure)
End Sub
'---------------------------------------------------------------------------------------
' Autor.....: RONAN VICO
' Contato...: ronanvico@hotmail.com.br - Empresa: Ronan Vico - Rotina: Public Sub InsertLineNumberInProc(actProc As String)
' Data......: 08/02/2020
' Descricao.: Insert the Lines on every single line of procedure from one Param to Another
'---------------------------------------------------------------------------------------
Public Sub InsertLineNumberInProc(actProc As String)
On Error GoTo TError
Dim textoProc                          As String
Dim linhas                             As Variant
Dim Linha                              As String
Dim c                                  As Long
Dim bESCREVER                          As Boolean
Dim ProcBodyStart                      As Long
Dim textJoin                           As Variant

    Call RemoveLineNumberProc(actProc)
    textoProc = GetProcedureTextWithoutBreakLines(actProc)
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
        '#pt Retira espacos em brancos , gambiarra mesmo to nem ai!
        '#en Remove Blank Spaces
        textJoin = VBA.Join(linhas, vbNewLine)
        textJoin = VBA.Replace(VBA.Replace(textJoin, vbNewLine, vbNewLine), QUEBRA_DE_LINHA, " _ " & vbNewLine)
        Call Application.VBE.ActiveCodePane.CodeModule.DeleteLines(ProcBodyStart, .ProcCountLines - (ProcBodyStart - .ProcStartLine))
        Call Application.VBE.ActiveCodePane.CodeModule.InsertLines(ProcBodyStart, textJoin)
    End With

Fim:
     Exit Sub
TError:
     GoTo Fim
     Resume
End Sub

 

'Remove numercao de Linha
Public Sub RemoveLineNumber()
    Call ProcedureInfo(ActiveProcedure, Application.VBE.ActiveCodePane.CodeModule, pInfo.procKind)
    Call RemoveLineNumberProc(ActiveProcedure)
End Sub

'---------------------------------------------------------------------------------------
' Autor.....: RONAN VICO
' Contato...: ronanvico@hotmail.com.br - Empresa: Ronan Vico - Rotina: Public Sub RemoveLineNumberProc(actProc As String)
' Data......: 08/02/2020
' Descricao.: REmove all Line Numbers from a Procedure
'---------------------------------------------------------------------------------------
Public Sub RemoveLineNumberProc(actProc As String)
Dim textoProc                          As String
Dim linhas                             As Variant
Dim Linha                              As String
Dim c                                  As Long
Dim numCarac                           As Long
Dim ProcBodyStart                      As Long
Dim textJoin                           As Variant
Dim pacoca                             As Variant

    textoProc = GetProcedureTextWithoutBreakLines(actProc)
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
        'Retira espacos em brancos , gambiarra mesmo to nem ai!
        'Remove blank spaces
        textJoin = VBA.Join(linhas, vbNewLine)
        textJoin = VBA.Replace(VBA.Replace(textJoin, vbNewLine, vbNewLine), QUEBRA_DE_LINHA, " _ " & vbNewLine)
        Call Application.VBE.ActiveCodePane.CodeModule.DeleteLines(ProcBodyStart, .ProcCountLines - (ProcBodyStart - .ProcStartLine))
        Call Application.VBE.ActiveCodePane.CodeModule.InsertLines(ProcBodyStart, textJoin)
    End With
End Sub

 
'Pegar ProcedureSemQuebraDeLinha
'---------------------------------------------------------------------------------------
' Autor.....: RONAN VICO
' Contato...: ronanvico@hotmail.com.br - Empresa: Ronan Vico - Rotina: Public Function GetProcedureTextWithoutBreakLines(ProcedureName As String) As String
' Data......: 08/02/2020
' Descricao.: Inside the proc is writed
'---------------------------------------------------------------------------------------
Public Function GetProcedureTextWithoutBreakLines(ProcedureName As String, Optional cModule As CodeModule) As String
'#PT Se tiver _ no final da linha de codigo , significa que a proxima linha pertence ao mesmo comando do VBA _
               Exemplo este comentario , a linha de baixo esta comentada de vido a esse Underline localizado /\ aqui _
               Logo , Iremos transformar tudo em "1 linha" para conseguirmos rodar funcoes de formatacao de texto _
               e posteriormente plotar de volta as qebras de llinhas no seus devidos lugares !
 
'#EN If there is _ at the end of the code line, it means that the next line belongs to the same VBA command _
            †Example this comment, the bottom line is commented due to that Underline located / \ here _
            Soon, we will transform everything into "1 line" to be able to run text formatting functions _
            and then plot back the lines of lines in their proper places!
 
 Dim proc                As ProcInfo
 Dim textAntComment      As String
 Dim textComment         As String
 Dim line                As String
 Dim texto               As String
 Dim i                   As Long
 Dim procKind            As VBIDE.vbext_ProcKind
 
    If cModule Is Nothing Then Set cModule = Application.VBE.ActiveCodePane.CodeModule
    proc = ProcedureInfo(ProcedureName, cModule, procKind)
    With proc
        i = .ProcBodyLine
        While i <= (.ProcCountLines + .ProcStartLine) - 1
            line = VBA.RTrim(Application.VBE.ActiveCodePane.CodeModule.Lines(i, 1))
            While VBA.Right(line, 1) = "_"
                i = i + 1
                line = VBA.Left(line, VBA.Len(line) - 1) & QUEBRA_DE_LINHA & VBA.RTrim(Application.VBE.ActiveCodePane.CodeModule.Lines(i, 1))
            Wend
            texto = texto & line & vbNewLine
            i = i + 1
        Wend
    End With
    GetProcedureTextWithoutBreakLines = texto
 End Function
'---------------------------------------------------------------------------------------
' Autor.....: RONAN VICO
' Contato...: ronanvico@hotmail.com.br - Empresa: Ronan Vico - Rotina: Public Function formataTexto(ByVal TextoOriginal As String, Optional ByRef MyArrVar) As String
' Data......: 08/02/2020
' Descricao.: Inside the procedure
'---------------------------------------------------------------------------------------
Public Function formataTexto(ByVal TextoOriginal As String, Optional ByRef MyArrVar) As String
    '#PT Transforma todas as strings em Variaveis dentro de um array , assim podemos manipular _
            todo o texto sem medo de estar mechendo com dados dentro de string , por exemplo _
            uma string pode ser MyString = "OLA : ' "  , Logo os caracteres ":" e "'" , sao _
            importantes na nossa formatacao de texto ,e nao podemos consideralo na formatacao, _
            por isso monto um array para posteriormente apos a formatacao pegar as strings e jogar de volta.

    '#EN Transform all strings into Variables within an array, so we can manipulate _
    all the text without fear of being messing with data inside a string, for example _
    †a string can be MyString = "OLA: '", So the characters ":" and "'", are _
    †important in our text formatting, and we cannot consider it in formatting, _
   †so I set up an array for later, after formatting, take the strings and play them back.
On Error GoTo TratarErro

Dim arrVars()                          As Variant
Dim contVar                            As Long
Dim i                                  As Long
Dim Y                                  As Long
Dim c                                  As String
Dim LenMax                             As Long
Dim texto                              As String
Dim tag                                As String
Dim ValorVariavel                      As String

    texto = TextoOriginal
    LenMax = VBA.Len(texto)
    'Debug.Print Texto
    While (i <= LenMax)
        i = i + 1
        c = VBA.Mid$(texto, i, 1)
        '#PT Verifica se e um caracter Aspas Dupla ""
        '#EN Check there is an Chr(34) ""
        If c = VBA.Chr(34) Then
            '#PT Percorre as proximas letras ate que a string seja fechada MyString= "abc" <Fim da string
            '#EN Run the next letters until the string is closed MyString= "abc" <End of string
            For Y = i + 1 To VBA.Len(texto)
                If VBA.Mid$(texto, Y, 1) = VBA.Chr(34) Then
                    If VBA.Mid$(texto, Y + 1, 1) = VBA.Chr(34) Then
                        Y = Y + 1
                    Else
                        Exit For
                    End If
                End If
            Next Y
           '#PT Adcionando a variavel nova ao vetor
           '#EN add new variable to the array
            ReDim Preserve arrVars(0 To contVar)
            ValorVariavel = VBA.Mid(texto, i + 1, (Y) - (i + 1))
            '#PT Aspas duplas de comentario estavam sendo apagadass
            '#EN Do not delete "" from Comments
            arrVars(contVar) = VBA.Chr(34) & ValorVariavel & VBA.Chr(34)
            tag = tagVarInit & contVar & tagVarEnd
            texto = VBA.Mid(texto, 1, i - 1) & tag & VBA.Mid(texto, Y + 1, VBA.Len(texto))
            contVar = contVar + 1
            i = i + VBA.Len(tag) - 1
            LenMax = VBA.Len(texto)
        End If
    Wend
    MyArrVar = arrVars
    formataTexto = texto
    Exit Function
TratarErro:
        Stop
        Resume
        Select Case Err.Number
                Case 0
                Case Else
                        MsgBox Err.Description & " " & Err.Number, vbCritical
        End Select
End Function

 
'---------------------------------------------------------------------------------------
' Autor.....: RONAN VICO
' Contato...: ronanvico@hotmail.com.br - Empresa: Ronan Vico - Rotina: Public Function PosComentario(ByVal text As String) As Long
' Data......: 08/02/2020
' Descricao.: Get the position of Comments
'---------------------------------------------------------------------------------------
Public Function PosComentario(ByVal text As String) As Long
    Dim auxtexto        As String
    auxtexto = formataTexto(text)
    PosComentario = (VBA.InStr(1, auxtexto, "'", vbTextCompare))
End Function
  
'---------------------------------------------\/  --------------------------------------------\/'- --------------------------------------------\/  --------------------------------------------\/
'#PT PARTE 4 --------------------------------------------\/ FUNCTIONS UTILIZADAS PARA PEGAR OS DADOS DO VBE
'#EN PART  4 --------------------------------------------\/ FUNCTIONS TO GET VBE DATAS
'---------------------------------------------\/  --------------------------------------------\/'- --------------------------------------------\/  --------------------------------------------\/
Private Function ProcedureInfo(ProcName As String, CodeMod As VBIDE.CodeModule, procKind As VBIDE.vbext_ProcKind) As ProcInfo
On Error GoTo TError
 Dim bodyLine            As Long
 Dim Declaration         As String
 Dim FirstLine           As String


    procKind = GetProcKind(ProcName, CodeMod)
    bodyLine = CodeMod.ProcStartLine(ProcName, procKind)
    If bodyLine > 0 Then
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
 
Fim:
     Exit Function
     Resume
TError:
     Call MsgBox(Err.Number & Err.Description)
     Stop
     GoTo Fim
 End Function
 
Private Function GetProcKind(ProcName As String, CodeMod As VBIDE.CodeModule) As VBIDE.vbext_ProcKind
    Dim i                   As Long
    Dim bodyLine           As Long
  'Get the Proc Kind
    On Error Resume Next
    Do
        Err.Clear
        bodyLine = CodeMod.ProcStartLine(ProcName, i)
        i = i + 1
    Loop While Err.Number <> 0 And i <= 3 'Vbext_Pk_Get
    GetProcKind = i - 1
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
 
'---------------------------------------------------------------------------------------
' Modulo....: cTOOLS \ CodeModule
' Rotina....: Public Function toUpperCase()
' Autor.....: RONAN VICO
' Contato...: RONANVICO@hotmail.com
' Data......: 30/04/2019
' Descricao.: Deixa a selecao do texto para maiuscula
' Descricao.: put the selection to upper case
'---------------------------------------------------------------------------------------
Public Function toUpperCase()
 On Error GoTo TratarErro
Dim i                                  As Long
Dim newText                            As String
Dim lineText                           As String
Dim sL                                 As Long
Dim eL                                 As Long
Dim sC                                 As Long
Dim eC                                 As Long


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

 
 
'---------------------------------------------------------------------------------------
' Autor.....: RONAN VICO
' Contato...: ronanvico@hotmail.com.br - Empresa: Ronan Vico - Rotina: Public Function toLowerCase()
' Data......: 08/02/2020
' Descricao.: Put the Selection Text in VBE to lower case
'---------------------------------------------------------------------------------------
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
    'MsgBox "Feito por Ronan Vico", vbInformation, "RVTool"
    frmRonanVico.Show
End Function

'Call the Unlock vba project
Public Function Hook()
    Call modVBACrack.Hook
End Function

'---------------------------------------------------------------------------------------
' Modulo....: cTOOLS \ CodeModule
' Rotina....: Public Function GetSelectedText() As String
' Autor.....: RONAN VICO
' Contato...: RONANVICO@hotmail.com
' Data......: 21/05/2019
' Descricao.: Pega o texto selecionado no codemodule e joga numa string
' Description .: get the Selection Text in CodeModule and put in a string Pegar Selecao
'---------------------------------------------------------------------------------------
Public Function GetSelectedText() As String
On Error GoTo TError
Dim sL&, sC&, eL&, eC&, i&
Dim S, lineText$, newText$

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
 GetSelectedText = newText
Fim:
     Exit Function
TError:
    ' Call MOSTRAR_ERRO(Err.Number, Err.Description, "GetSelectedText()")
     GoTo Fim
End Function

'---------------------------------------------------------------------------------------
' Modulo....: cTOOLS \ CodeModule
' Rotina....: Public Sub CopyText(ByVal TransferAreaNumber As Long)
' Autor.....: RONAN VICO
' Contato...: RONANVICO@hotmail.com
' Data......: 21/05/2019
' Descricao.: Joga o texto selecionado no registro dos settings do vba para depois PasteText
' Description.: Put the electionText in Registry so we can read it later
'---------------------------------------------------------------------------------------
Public Sub CopyText(ByVal TransferAreaNumber As Long)
On Error GoTo TError
    Call VBA.SaveSetting(C_APPNAME, C_SECTION_CopyText, TransferAreaNumber, GetSelectedText())
Fim:
     Exit Sub
     Resume
TError:
     GoTo Fim
End Sub

'---------------------------------------------------------------------------------------
' Modulo....: cTOOLS \ CodeModule
' Rotina....: Public Sub PasteText(ByVal TransferAreaNumber As Long)
' Autor.....: RONAN VICO
' Contato...: RONANVICO@hotmail.com
' Data......: 21/05/2019
' Descricao.: Joga o texto na onde o cursor estiver selecionado
' Description.: Put the text where the cursor in Code Module is
'---------------------------------------------------------------------------------------
Public Sub PasteText(ByVal TransferAreaNumber As Long)
On Error GoTo TError
    Dim sL&, sC&, eL&, eC&
    Dim StringParaPasteText$
    Dim eLString$
    Dim lineText$, newText$
    Dim i&
    
    Call Application.VBE.ActiveCodePane.GetSelection(sL, sC, eL, eC)
    'Pega do registro o texto a PasteText
    StringParaPasteText = VBA.GetSetting(C_APPNAME, C_SECTION_CopyText, TransferAreaNumber)

     If sL = eL Then
        lineText = Application.VBE.ActiveCodePane.CodeModule.Lines(sL, 1)
        newText = VBA.Mid(lineText, 1, sC - 1) & StringParaPasteText & VBA.Mid(lineText, eC)
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
        StringParaPasteText = Application.VBE.ActiveCodePane.CodeModule.Lines(sL, 1) & StringParaPasteText
        Call Application.VBE.ActiveCodePane.CodeModule.ReplaceLine(sL, StringParaPasteText)
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
' Rotina....: Public Sub CleanPasteText()
' Autor.....: RONAN VICO
' Contato...: RONANVICO@hotmail.com
' Data......: 21/05/2019
' Descricao.: 'Retira todos os PasteTexts da memoria
'---------------------------------------------------------------------------------------
Public Sub CleanPasteText()
On Error GoTo TError
    If MsgBox("Deseja mesmo Limpar todos os ""PasteText""?", vbYesNo + vbExclamation, "Atencao") = vbNo Then Exit Sub
    Dim i&
    
    For i = 1 To 10
        Call VBA.SaveSetting(C_APPNAME, C_SECTION_CopyText, i, "")
    Next i
    
Fim:
     Exit Sub
TError:
     Debug.Print "erro ao rodar CleanPasteText"
     GoTo Fim
End Sub

Public Sub GetFunctionAndSubNames()
On Error GoTo TError
    Dim item            As Variant

    For Each item In Application.VBE.ActiveVBProject.VBComponents
            ListProcedures item.Name, True
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
Dim vbProj                             As VBIDE.VBProject
Dim vbCOMP                             As VBIDE.VBComponent
Dim CodeMod                            As VBIDE.CodeModule
Dim LineNum                            As Long
Dim ProcName                           As String
Dim procKind                           As VBIDE.vbext_ProcKind
Dim strSubsInfo                        As String

    Set vbProj = Application.VBE.ActiveVBProject
    Set vbCOMP = vbProj.VBComponents(strName)
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


'---------------------------------------------------------------------------------------
' Autor.....: RONAN VICO
' Contato...: ronanvico@hotmail.com.br - Empresa: Ronan Vico - Rotina: Public Sub CloseAllWindowsCodeModule()
' Data......: 08/02/2020
' Descricao.: Close all Windows Opened (CodeModules)
'---------------------------------------------------------------------------------------
Public Sub CloseAllWindowsCodeModule()
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
Public Sub CloseProjectExplorer()
    Dim w As VBIDE.Window
    Dim i As Long
    On Error Resume Next
    For Each w In Application.VBE.Windows
        If w.Type = vbext_wt_ProjectWindow Or w.Type = vbext_wt_PropertyWindow Then w.Close
    Next
End Sub


'---------------------------------------------------------------------------------------
' Autor.....: RONAN VICO
' Contato...: ronanvico@hotmail.com.br - Empresa: Ronan Vico - Rotina: Public Sub CheckVariablesNotUsedInProcedure()
' Data......: 08/02/2020
' Descricao.: This is a beautyFul Procedure ,
' that proc will check if one Variable declared is not used , and if so , then will
' print in imediate window
'---------------------------------------------------------------------------------------
Public Sub CheckVariablesNotUsedInProcedure()
On Error GoTo Fim
If ActiveProcedure = "" Then
    Exit Sub
End If

On Error GoTo TError
Dim textoProc                          As String
Dim linhas                             As Variant
Dim Linha                              As String
Dim LinhaVerificacaoSeFoiUsada         As String
Dim contLine                           As Long
Dim contLine2                          As Long
Dim ArrAux                             As Variant
Dim AntComment                         As Variant
Dim LinhaSemComentario                 As String
Dim arrVars                            As Variant
Dim ArrVarsPlus                        As Variant
Dim LinhaDeDeclaracoes                 As Variant
Dim VarDeclaration                     As Variant
Dim strRegexPatter                     As String
Dim DicVariaveis                       As New Scripting.Dictionary
Dim VariaveisNaoUsadas                 As Variant
Dim contVarsNaoUsadas                  As Long
Dim lin                                As Long

    textoProc = GetProcedureTextWithoutBreakLines(ActiveProcedure)
    textoProc = formataTexto2(textoProc)
    linhas = VBA.Split(textoProc, vbNewLine)
    For lin = LBound(linhas, 1) To UBound(linhas, 1)
        If linhas(lin) <> vbNullString Then
            linhas(lin) = VBA.Split(linhas(lin), "'")(0)
        End If
    Next lin
    For contLine = 0 To UBound(linhas)
        Linha = linhas(contLine)
        Linha = formataTexto(Linha, ArrAux)

        If Linha = vbNullString Then Linha = " "
        LinhaSemComentario = VBA.Split(Linha, "'")(0)

        For Each AntComment In VBA.Split(LinhaSemComentario, ":")
            If IsLinhaMatch(AntComment, " Dim ") Or VBA.Left(AntComment, 4) = "Dim " Then
                If VBA.Left(AntComment, 4) = "Dim " Then AntComment = " " & AntComment
                arrVars = VBA.Split(AntComment, " Dim ")
                LinhaDeDeclaracoes = arrVars(1)
                arrVars = VBA.Split(LinhaDeDeclaracoes, ",")
                For Each VarDeclaration In arrVars
                    VarDeclaration = VBA.Split(VBA.Split(VarDeclaration, " As ")(0), "(")(0)
                    VarDeclaration = VBA.Replace(VBA.Replace(VarDeclaration, " ", ""), vbTab, "")
                    DicVariaveis.Add VarDeclaration, False
                    strRegexPatter = "(\(v\)|\(v,|, v\)|,v\)|^v = | v = |,v,|, v,| v = |^v | v |v\()"
                    strRegexPatter = VBA.Replace(strRegexPatter, "v", VarDeclaration)
                    For contLine2 = 0 To UBound(linhas)
                        LinhaVerificacaoSeFoiUsada = linhas(contLine2)
                        If contLine2 <> contLine Then
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
            If PT_BR() Then
                Debug.Print "Variavel Nao Usada: " & VariaveisNaoUsadas
            Else
                Debug.Print "Unused Variable   : " & VariaveisNaoUsadas
            End If
            contVarsNaoUsadas = contVarsNaoUsadas + 1
        End If
    Next

    If contVarsNaoUsadas = 0 Then
        If PT_BR() Then
            Debug.Print "Todas as variaveis estao sendo utilizadas , Parabens mano!"
        Else
            Debug.Print "Congratz Bro , you do not have Unused Variables in this proc!"
        End If
    End If

Fim:
     Exit Sub
     Resume
TError:
     'Call MOSTRAR_ERRO(Err.Number, Err.Description, "CheckVariablesNotUsedInProcedure()")
     Stop
     GoTo Fim
End Sub




'---------------------------------------------------------------------------------------
' Autor.....: RONAN VICO
' Contato...: ronanvico@hotmail.com.br - Empresa: Ronan Vico - Rotina: Public Function formataTexto2(ByVal TextoOriginal As String, Optional ByRef MyArrVar) As String
' Data......: 08/02/2020
'---------------------------------------------------------------------------------------
Public Function formataTexto2(ByVal TextoOriginal As String, Optional ByRef MyArrVar) As String
'#PT Transforma todas as strings em Variaveis dentro de um array , assim podemos manipular _
            todo o texto sem medo de estar mechendo com dados dentro de string , por exemplo _
            uma string pode ser MyString = "OLA : ' "  , Logo os caracteres ":" e "'" , sao _
            importantes na nossa formatacao de texto ,e nao podemos consideralo na formatacao, _
            por isso monto um array para posteriormente apos a formatacao pegar as strings e jogar de volta.

'#EN Transform all strings into Variables within an array, so we can manipulate _
                                            all the text without fear of being messing with data inside a string, for example _
                                            a string can be MyString = "OLA: '", So the characters ":" and "'", are _
                                            important in our text formatting, and we cannot consider it in formatting, _
                                            so I set up an array for later, after formatting, take the strings and play them back.
 On Error GoTo TratarErro

Dim arrVars()           As Variant: Dim contVar            As Long: Dim i                  As Long:
Dim Y                   As Long
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
         'Verifica se e um caracter Aspas Dupla ""
         If c = VBA.Chr(34) Then
             'Percorre as proximas letras ate que a string seja fechada MyString= "abc" <Fim da string
             For Y = i + 1 To VBA.Len(texto)
                 If VBA.Mid$(texto, Y, 1) = VBA.Chr(34) Then
                     If VBA.Mid$(texto, Y + 1, 1) = VBA.Chr(34) Then
                         Y = Y + 1
                     Else
                         Exit For
                     End If
                 End If
             Next Y
            'Adcionando a variavel nova ao vetor
             ReDim Preserve arrVars(0 To contVar)
             ValorVariavel = VBA.Mid(texto, i + 1, (Y) - (i + 1))
             'Aspas duplas de comentario estavam sendo apagadass
             arrVars(contVar) = VBA.Chr(34) & ValorVariavel & VBA.Chr(34)
             tag = tagVarInit & contVar & tagVarEnd
             texto = VBA.Mid(texto, 1, i - 1) & tag & VBA.Mid(texto, Y + 1, VBA.Len(texto))
             contVar = contVar + 1
             i = i + VBA.Len(tag) - 1
             LenMax = VBA.Len(texto)
         End If

     Wend

     MyArrVar = arrVars
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


Public Sub Change_color_White_Theme()
    CreateObject("WScript.Shell").RegWrite REG_BACK_COLOR, BACK_COLOR_WHITE_THEME, "REG_SZ"
    CreateObject("WScript.Shell").RegWrite Reg_FORE_COLOR, FORE_COLOR_WHITE_THEME, "REG_SZ"
If PT_BR Then
    MsgBox "Pronto!, Reabra o Excel", vbInformation, "Reabra o Excel!"
Else
    MsgBox "Finished!, Re-open the Excel", vbInformation, "Re open the Excel!"
End If
End Sub

Public Sub Change_color_Dark_Theme()
    CreateObject("WScript.Shell").RegWrite REG_BACK_COLOR, BACK_COLOR_BLACK_THEME, "REG_SZ"
    CreateObject("WScript.Shell").RegWrite Reg_FORE_COLOR, FORE_COLOR_BLACK_THEME, "REG_SZ"
If PT_BR Then
    MsgBox "Pronto!, Reabra o Excel", vbInformation, "Reabra o Excel!"
Else
    MsgBox "Finished!, Re-open the Excel", vbInformation, "Re open the Excel!"
End If
End Sub


'---------------------------------------------------------------------------------------
' Autor.....: RONAN VICO
' Contato...: ronanvico@hotmail.com.br - Empresa: Ronan Vico - Rotina: Public Sub Atualizar_RVTool()
' Data......: 08/02/2020
' Descricao.: Update RVTOOL
'---------------------------------------------------------------------------------------
Public Sub Atualizar_RVTool()

    If ThisWorkbook.VBProject.Protection = vbext_pp_locked Then
#If PT_BR Then
     MsgBox "Desproteja o VBE do RVTool!", vbExclamation, "ATENa√¢‚Ç¨¬°a√Ü‚ÄôO!"
#Else
    MsgBox "Unlock the VB Project in RVTool!", vbExclamation, "Atention!"
#End If
        Exit Sub
    End If

On Error GoTo TError
Dim FSO                                As New Scripting.FileSystemObject
Dim fl                                 As Scripting.File
Dim fold                               As Scripting.Folder
Dim vbProj                             As VBIDE.VBProject
Dim vbCOMP                             As VBIDE.VBComponent
Dim oAPP                               As Object
Dim objZipItems                        As Object
Dim objFolderDestination               As Object
Dim PASTA                              As String
Dim i                                  As Long
Dim File                               As String
Dim Folder                             As String
    Const URL_PROJETO = "https://github.com/RonanVico/RVTool/archive/master.zip"


    Folder = ThisWorkbook.Path & "\RVTool" & VBA.Format(VBA.Now(), "ddmmyyyyhhmmss") & "new"
    Call VBA.FileSystem.MkDir(Folder)
    Call URLDownloadToFile(0, URL_PROJETO, Folder & ".zip", 0, 0)
    File = Folder & ".zip\RVTool-master"
    With CreateObject("Shell.Application")
        .Namespace(Folder & "\").CopyHere .Namespace(Folder & ".zip\").Items
    End With

    For i = 1 To 1000: VBA.DoEvents: Next i

    PASTA = Folder & "\RVTool-master"
    Set fold = FSO.GetFolder(PASTA)
    Set vbProj = ThisWorkbook.VBProject
    For Each fl In fold.Files
        If IsLinhaMatch(VBA.LCase(VBA.Split(fl.Name, ".")(UBound(VBA.Split(fl.Name, ".")))), "(bas|cls|md|frm)$") Then
            For Each vbCOMP In vbProj.VBComponents
                If VBA.UCase(vbCOMP.Name) = VBA.UCase(VBA.Split(fl.Name, ".")(0)) Then
                    vbCOMP.Name = vbCOMP.Name & "DELETAR__"
                    Exit For
                End If
            Next vbCOMP
            Call vbProj.VBComponents.Import(fl.Path)
        End If
    Next fl

    For Each vbCOMP In vbProj.VBComponents
        If VBA.UCase(vbCOMP.Name) Like "*DELETAR__*" Then
            Call vbProj.VBComponents.Remove(vbCOMP)
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
    
    Set fold = FSO.GetFolder(ThisWorkbook.Path)
    
    For Each fold2 In fold.SubFolders
        If VBA.LCase(fold2.Name) Like "*rvtool*new*" Then
            fold2.DELETE
        End If
    Next fold2
End Sub




'---------------------------------------------------------------------------------------
' Autor.....: RONAN VICO
' Contato...: ronanvico@hotmail.com.br - Empresa: Ronan Vico - Rotina: Public Sub Snippets()
' Data......: 08/02/2020
' Descricao.: That function its a Auto Complete (Code Snippets) for VBA
'---------------------------------------------------------------------------------------
Public Sub Snippets()
On Error GoTo TError
Dim sL&
Dim sC&
Dim eL&
Dim eC&
Dim i&
Dim NEWsL&
Dim NEWsC&
Dim NEWeL&
Dim NEWeC&
Dim SomarSC&
Dim SomarEC&
Dim ProcInfo                           As ProcInfo
Dim procKind                           As VBIDE.vbext_ProcKind
Dim Linha                              As String
Dim TextoVerificar                     As String
Dim CodigoParaInserir                  As String
Dim novaLInha                          As String
Dim tamanhoSnippet                     As Long
Dim auxParam1                          As Variant
Dim Middle                             As Boolean

    '#PT Pega a localizacao do cursor no codigo
    '#EN  Get the cursor position on code
    Call Application.VBE.ActiveCodePane.GetSelection(sL, sC, eL, eC)
    'PEga a linha inteira ate a primeira posicao do cursor do teclado
    '#EN Get the Inteire line until the first position of cursor
    Linha = VBA.Replace(VBA.Trim(VBA.Mid(Application.VBE.ActiveCodePane.CodeModule.Lines(sL, 1), 1, eC - 1)), " (", "(")
    'Retira os tabs por precaucao
    '#EN remove vbTabs
    Linha = VBA.Replace(Linha, vbTab, " ")
    If VBA.InStr(Linha, " ") <> 0 Then
        TextoVerificar = VBA.Mid(Linha, VBA.InStrRev(Linha, " ") + 1)
    Else
        TextoVerificar = Linha
    End If

    'retorna variavel da linha
    'Return the variable of line
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
            auxParam1 = VBA.Replace(VBA.Split(TextoVerificar, "sub(", , vbTextCompare)(1), ")", "")
            tamanhoSnippet = 5 + VBA.Len(auxParam1) + IIf(VBA.LCase(TextoVerificar) Like "*sub (*", 1, 0)
            CodigoParaInserir = "Public Sub " & auxParam1 & "()" & _
            vbNewLine & vbNewLine & "End Sub"
            Middle = True
        Case VBA.LCase(TextoVerificar) Like "*fn (*" Or VBA.LCase(TextoVerificar) Like "*fn(*"
            auxParam1 = VBA.Replace(VBA.Split(TextoVerificar, "fn(", , vbTextCompare)(1), ")", "")
            tamanhoSnippet = VBA.Len(auxParam1) + 4 + IIf(TextoVerificar Like "*fn (*", 0, 1): SomarSC = 16 + VBA.Len(auxParam1) + 6: SomarEC = 7
            CodigoParaInserir = "Public Function " & auxParam1 & "() as Variant" & _
            vbNewLine & vbNewLine & "End Function"
        Case VBA.LCase(TextoVerificar) Like "*prop (*" Or VBA.LCase(TextoVerificar) Like "*prop(*"
            auxParam1 = VBA.Replace(VBA.Split(TextoVerificar, "prop(", , vbTextCompare)(1), ")", "")
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
        Call InsertErrorTreatment
        
        If Middle = True Then
            Call Application.VBE.ActiveCodePane.SetSelection(sL + 2, 1, sL + 2, 1)
        Else
            Call Application.VBE.ActiveCodePane.SetSelection(sL, sC + SomarSC - tamanhoSnippet, sL, sC + SomarEC + SomarSC - tamanhoSnippet)
        End If
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
    Dim FSO                                As Object
    Dim oAPP                               As Object
    Dim Fname                              As Variant
    Dim FileNameFolder                     As Variant
    Dim DefPath                            As String
    Dim strDate                            As String
    
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
        'oApp.Namespace(FileNameFolder).CopyTextHere _
        'oApp.Namespace(Fname).items.Item("test.txt")
        
        MsgBox "You find the files here: " & FileNameFolder
        
        On Error Resume Next
        Set FSO = CreateObject("scripting.filesystemobject")
        FSO.DeleteFolder Environ("Temp") & "\Temporary Directory*", True
    End If
End Sub

Public Sub IndentarProcedure()
    Call IdentAProcedure(ActiveProcedure())
End Sub


'-- FEITO POR RONAN VICO 04/02/2020
'-- NA LIVE DO GUIA DO EXCEL
Public Sub IdentAProcedure(Optional procedure As String = vbNullString, Optional vbCodeModule As CodeModule)
    On Error GoTo TError
Dim DicOpenIdentWords                  As New Scripting.Dictionary
Dim arrWords                           As Variant
Dim lin                                As Long
Dim arrLinhas                          As Variant
Dim strLinha                           As String
Dim Key                                As Variant
Dim bSomarTAB                          As Boolean
Dim bSubtrairTAB                       As Boolean
Dim contQuantidadeIndentar             As Long
Dim vbProj                             As VBIDE.VBProject
Dim NEW_CODE_TEXT                      As String

    If vbCodeModule Is Nothing Then
        Set vbCodeModule = Application.VBE.ActiveCodePane.CodeModule
    End If
    If procedure = vbNullString Then
        procedure = ActiveProcedure
    End If


    Const caracterDeIndentacao  As String = PARAM_CHAR_IndentacaO
    
    contQuantidadeIndentar = 0
    With DicOpenIdentWords
        .Add "Do |Do$", "(Loop)"
        .Add "For ", "(Next)"
        'This line will be ok for if and Elseif
        .Add "If .*Then$", "(Else$|ElseIf .* Then$|End If$)"
'        .Add "ElseIf ", "(Else$|ElseIf.*&Then$|End If$)"
        .Add "Else$", "(End If)"
        .Add "While ", "(Wend)"
        .Add "Select Case", "(End Select)"
        .Add "Case ", "(Case |End Select)"
        .Add "With ", "(End With$)"
        .Add "Public |Private |Friend ", "(End Sub|End Function|End Property)"
        .Add "Sub |Function |Property ", "(End Sub|End Function|End Property)"
    End With

Dim TEXTO_DA_PROC                      As String
Dim arrVars                            As Variant
    TEXTO_DA_PROC = formataTexto2(GetProcedureTextWithoutBreakLines(procedure), arrVars)
    TEXTO_DA_PROC = VBA.Replace(GetProcedureTextWithoutBreakLines(procedure), QUEBRA_DE_LINHA, "_" & vbCrLf)
    arrLinhas = VBA.Split(TEXTO_DA_PROC, vbCrLf)


    'Joga tudo para a esquerda para come√É¬ßar Indentar
    For lin = LBound(arrLinhas, 1) To UBound(arrLinhas, 1)
        'Problema com numeracao de linhas
        strLinha = VBA.LTrim(arrLinhas(lin))
        If VBA.IsNumeric(VBA.Left(strLinha, 1)) Then
            strLinha = VBA.Mid(strLinha, VBA.InStr(strLinha, " "))
        End If
        arrLinhas(lin) = strLinha
    Next lin



    For lin = LBound(arrLinhas, 1) To UBound(arrLinhas, 1)
        'devido a numeracao de linhas
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
            'Check the line to Subtract Vb Tab
            If IsLinhaMatch(strLinha, DicOpenIdentWords(arrWords(UBound(arrWords)))) Then
                bSubtrairTAB = True
                'Clausula especial para Select  Case , pois Indentacao e diferente
                'select case is only one that is special , becoz is diferent
                If arrWords(UBound(arrWords)) = "Case " And IsLinhaMatch(strLinha, "^(End Select)") Then
                    contQuantidadeIndentar = contQuantidadeIndentar - 1
                    'Pop the Stack
                    If UBound(arrWords) = 0 Then Set arrWords = Nothing Else ReDim Preserve arrWords(0 To UBound(arrWords) - 1)
                End If
                'Pop the Stack
                If UBound(arrWords) = 0 Then Set arrWords = Nothing Else ReDim Preserve arrWords(0 To UBound(arrWords) - 1)
            End If
        End If
        For Each Key In DicOpenIdentWords
            'verifica linha para somar
            'Check if is the line to Add VB TAB
            If IsLinhaMatch(strLinha, "^(" & Key & ")") And Not IsLinhaMatch(strLinha, DicOpenIdentWords(Key)) Then
                bSomarTAB = True
                'Push the Stack
                If VBA.IsArray(arrWords) Then
                    ReDim Preserve arrWords(0 To UBound(arrWords) + 1)
                Else
                    arrWords = VBA.Split("", "")
                End If

                arrWords(UBound(arrWords)) = Key
                Exit For
            End If

        Next Key

        If bSubtrairTAB Then contQuantidadeIndentar = contQuantidadeIndentar - 1

        strLinha = arrLinhas(lin)
        If VBA.IsNumeric(VBA.Left(strLinha, 1)) Then
            strLinha = VBA.Mid(strLinha, 1, VBA.InStr(strLinha, " ") - 1) & VBA.Strings.String(contQuantidadeIndentar, caracterDeIndentacao) & VBA.Mid(strLinha, VBA.InStr(strLinha, " "))
        Else
            arrLinhas(lin) = VBA.Strings.String(contQuantidadeIndentar, caracterDeIndentacao) & strLinha
        End If
        If bSomarTAB Then contQuantidadeIndentar = contQuantidadeIndentar + 1
    Next lin

    '25 02 2020
    If UBound(arrLinhas) >= 0 Then
        While arrLinhas(UBound(arrLinhas)) = vbNullString
            ReDim Preserve arrLinhas(0 To UBound(arrLinhas) - 1)
        Wend
    
        NEW_CODE_TEXT = VBA.Join(arrLinhas, vbCrLf)
    
        With pInfo
            Call vbCodeModule.DeleteLines(.ProcBodyLine, .ProcCountLines - (.ProcBodyLine - .ProcStartLine))
            Call vbCodeModule.InsertLines(.ProcBodyLine, NEW_CODE_TEXT)
        End With
    End If
    
    Call IndentVariables(procedure)

Fim:
    Exit Sub
TError:
    'Call MOSTRAR_ERRO(Err.Number, Err.Description, "dev()")
    Stop
    Resume
    GoTo Fim
End Sub


'---------------------------------------------------------------------------------------
' Autor.....: RONAN VICO
' Contato...: ronanvico@hotmail.com.br - Empresa: Ronan Vico - Rotina: Public Sub IdentModule()
' Data......: 25/02/2020
' Descricao.: Run to ident all the module
'---------------------------------------------------------------------------------------
Public Sub IdentModule()
    Dim msg As String
    If PT_BR() Then
        msg = "Deseseja realmente identar o mÛdulo inteiro ?! Pode demorar um pouco..."
    Else
        msg = "Are you sure that you want ident all the module?! This could take a while..."
    End If
    If MsgBox(msg, vbExclamation + vbYesNo, IIf(PT_BR, "Identar Modulo", "Ident Module")) = vbYes Then
        IdentTheModule Application.VBE.ActiveCodePane.CodeModule
    End If
End Sub

'---------------------------------------------------------------------------------------
' Autor.....: RONAN VICO
' Contato...: ronanvico@hotmail.com.br - Empresa: Ronan Vico - Rotina: Private Sub IdentTheModule(strName As String)
' Data......: 25/02/2020
' Descricao.: Ident All the Module
'---------------------------------------------------------------------------------------
Private Sub IdentTheModule(strName As String)
On Error GoTo TError
    'Microsoft Visual Basic for Applications Extensibility 5.3 library
Dim vbProj                             As VBIDE.VBProject
Dim vbCOMP                             As VBIDE.VBComponent
Dim CodeMod                            As VBIDE.CodeModule
Dim LineNum                            As Long
Dim ProcName                           As String
Dim procKind                           As VBIDE.vbext_ProcKind
Dim strSubsInfo                        As String
Dim i                                  As Long
Dim arrProc

    Set vbProj = Application.VBE.ActiveVBProject
    Set vbCOMP = vbProj.VBComponents(strName)
    Set CodeMod = vbCOMP.CodeModule

    With CodeMod
        LineNum = .CountOfDeclarationLines + 1

        Do Until LineNum >= .CountOfLines
            ProcName = .ProcOfLine(LineNum, procKind)
            strSubsInfo = strSubsInfo & ";" & ProcName
            LineNum = .ProcStartLine(ProcName, procKind) + .ProcCountLines(ProcName, procKind) + 1
        Loop
    End With
    
    If VBA.Len(strSubsInfo) > 1 Then
        strSubsInfo = VBA.Mid(strSubsInfo, 2)
        arrProc = VBA.Split(strSubsInfo, ";")
        For i = LBound(arrProc, 1) To UBound(arrProc, 1)
            Call IdentAProcedure(VBA.CStr(arrProc(i)), CodeMod)
        Next i
    End If
Fim:
     Exit Sub
TError:
     'Call MOSTRAR_ERRO(Err.Number, Err.Description, "ListProcedures()")
     Stop
     GoTo Fim
End Sub


Public Sub Change_Region()
    If getLenguage() = ENGLISH Or getLenguage() = NOT_DEFINED Then
        Call setLenguague(PORTUGUES)
    Else
        Call setLenguague(ENGLISH)
    End If
End Sub
Public Sub ExcelContextMenu()
    Call mod_WindowsRegistry.CreateContextMenuForExcel
End Sub

Ôæ        .                             D e s k t o p    J 1          MOS 8 	  Ôæ        .                             M O S    x 1          MOSExcelExpert2016  V 	  Ôæ        .                             M O S E x c e l E x p e r t 2 0 1 6   " ` 1          Objective1  F 	  Ôæ        .                             O b j e c t i v e 1    ¨ 2 rG  NIÂi  ExcelExpert_1-2b.xlsx \ 	  ÔæXNô†XNÙÆ.   íè                      E x c e l E x p e r t _ 1 - 2 b . x l s x   $ ,    Ôæ E x c e l . S h e e t . 1 