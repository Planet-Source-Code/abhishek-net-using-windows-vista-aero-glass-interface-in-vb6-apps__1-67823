Attribute VB_Name = "Module1"
Option Explicit

'==========================================APIs==========================================
Private Declare Function InitCommonControls Lib "COMCTL32.DLL" () As Long

Public Sub Main()
    Dim X As Long
    X = InitCommonControls
    
    Form1.Show
End Sub

