VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVBECommandHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents EvtHandler As VBIDE.CommandBarEvents
Attribute EvtHandler.VB_VarHelpID = -1

Private Sub EvtHandler_Click(ByVal CommandBarControl As Object, _
    Handled As Boolean, CancelDefault As Boolean)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This is called when a item is clicked. Call the
    ' procedure named in the OnAction property of the
    ' CommandBarControl passed into this procedure.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    
   ' CommandBarControl.Reset
    On Error Resume Next
    'Reinstanciar a barra
    Call CallByName(Application, "OnTime", VbMethod, VBA.Now() + VBA.TimeSerial(0, 0, 5), "InitVBRVTools")
    'Application.Run CommandBarControl.OnACtion
    Dim tool As New cTOOLS
    
    If VBA.Left(VBA.LCase(CommandBarControl.CAption), 6) = "copiar" Or VBA.Left(VBA.LCase(CommandBarControl.TooltipText), 5) = "colar" Then
        Call CallByName(tool, CommandBarControl.OnACtion, VbMethod, CommandBarControl.DescriptionText)
    Else
        Call CallByName(tool, CommandBarControl.OnACtion, VbMethod)
    End If
    Set tool = Nothing
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Tell VBIDE that we've handled the event.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    Handled = True
    CancelDefault = True
    
    'Reinstanciar a barra
    Call CallByName(Application, "OnTime", VbMethod, VBA.Now(), "InitVBRVTools")
    
    
End Sub





