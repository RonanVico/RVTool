VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRonanVico 
   Caption         =   "Ronan Vico - VBA <3  - I love you guys"
   ClientHeight    =   2625
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6015
   OleObjectBlob   =   "frmRonanVico.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRonanVico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub imG_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call VBA.Shell("Explorer.exe " & VBA.Chr(34) & txtG.Value & VBA.Chr(34))
End Sub

Private Sub imL_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call VBA.Shell("Explorer.exe " & VBA.Chr(34) & txtL.Value & VBA.Chr(34))
End Sub

Private Sub imY_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call VBA.Shell("Explorer.exe " & VBA.Chr(34) & txtY.Value & VBA.Chr(34))
End Sub

