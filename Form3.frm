VERSION 5.00
Begin VB.Form Bigplus 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00332E2B&
   BorderStyle     =   0  'None
   ClientHeight    =   5535
   ClientLeft      =   5505
   ClientTop       =   2580
   ClientWidth     =   8910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   4440
      X2              =   4440
      Y1              =   1560
      Y2              =   5400
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "  +"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   6840
      TabIndex        =   10
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "  -"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4800
      TabIndex        =   9
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "  +"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2160
      TabIndex        =   8
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "  -"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   " x"
      BeginProperty Font 
         Name            =   "方正兰亭超细黑简体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   8280
      TabIndex        =   6
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "大组加分-七二教学助手"
      BeginProperty Font 
         Name            =   "方正兰亭超细黑简体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   2775
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      Height          =   2895
      Left            =   4800
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   2895
      Left            =   120
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "南半球"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   5880
      TabIndex        =   3
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404040&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   150
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   4800
      TabIndex        =   2
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404040&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   150
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "北半球"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   375
      Left            =   8400
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label12 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9375
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C9A70D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C9A70D&
      Height          =   735
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C9A70D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C9A70D&
      Height          =   735
      Left            =   2160
      Shape           =   4  'Rounded Rectangle
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00C9A70D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C9A70D&
      Height          =   735
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00C9A70D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C9A70D&
      Height          =   735
      Left            =   6840
      Shape           =   4  'Rounded Rectangle
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "  清除"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3480
      TabIndex        =   11
      Top             =   960
      Width           =   1935
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000080FF&
      Height          =   495
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "Bigplus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Integer
Dim b As Integer
Dim xa As Single, ya As Single
Private Sub Form_Load()
a = 0
b = 0
End Sub


Private Sub Label10_Click()
a = 0
b = 0
Label4.Caption = b
Label3.Caption = a
End Sub

Private Sub Label11_Click()
Me.Visible = False
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xa = X
ya = Y
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Me.Move Me.Left + X - xa, Me.Top + Y - ya
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xa = X
ya = Y
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Me.Move Me.Left + X - xa, Me.Top + Y - ya
End Sub

Private Sub Label6_Click()
a = a - 1
Label3.Caption = a
End Sub

Private Sub Label7_Click()
a = a + 1
Label3.Caption = a
End Sub

Private Sub Label8_Click()
b = b - 1
Label4.Caption = b
End Sub

Private Sub Label9_Click()
b = b + 1
Label4.Caption = b
End Sub
