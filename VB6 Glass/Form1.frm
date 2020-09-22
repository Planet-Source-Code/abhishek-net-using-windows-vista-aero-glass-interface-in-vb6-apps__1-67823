VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VB6 Vista Glass Demo"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGlass 
      Caption         =   "Apply Glass"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblReturn 
      AutoSize        =   -1  'True
      Caption         =   "Return:"
      Height          =   195
      Left            =   1680
      TabIndex        =   1
      Top             =   1440
      Width           =   555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form1
' Created   : 08-Feb-07
' Author    : abhishek abhishek007p@hotmail.com
' Purpose   : Code to Use Windows Vista Glass Interface in VB6
' System    : Windows Vista+
' IE Version: N/A
' Dependency: N/A
' Notes     :
'---------------------------------------------------------------------------------------
'
Option Explicit

'=========================================Types==========================================
Private Type tRect
    m_Left      As Long
    m_Right     As Long
    m_Top       As Long
    m_Buttom    As Long
End Type

'==========================================APIs==========================================
Private Declare Function ApplyGlass Lib "dwmapi.dll" Alias "DwmExtendFrameIntoClientArea" (ByVal hWnd As Long, rect As tRect) As Long

Private Sub cmdGlass_Click()

    Dim GRect       As tRect
    Dim lngReturn   As Long
    
    GRect.m_Buttom = 50
    GRect.m_Left = 50
    GRect.m_Right = 50
    GRect.m_Top = 50
    
    Me.BackColor = vbBlack

    lngReturn = ApplyGlass(Me.hWnd, GRect)
    
    lblReturn.Caption = "Return: " & lngReturn
    
End Sub
