VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   16200
   ClientLeft      =   -210
   ClientTop       =   -405
   ClientWidth     =   28800
   BeginProperty Font 
      Name            =   "黑体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   16200
   ScaleWidth      =   28800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " 最小化"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   15720
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "练字"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   699.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   15135
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   28815
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C9A70D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C9A70D&
      Height          =   375
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   15720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Integer
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Private Sub Label2_Click()
Me.Visible = False
End Sub
