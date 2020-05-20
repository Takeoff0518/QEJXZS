VERSION 5.00
Begin VB.Form Number 
   BackColor       =   &H00332E2B&
   BorderStyle     =   0  'None
   ClientHeight    =   6240
   ClientLeft      =   7620
   ClientTop       =   2790
   ClientWidth     =   4545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   360
      Top             =   3720
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   120
      Max             =   50
      Min             =   2
      TabIndex        =   5
      Top             =   5520
      Value           =   32
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   3975
      Left            =   120
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "人数：32"
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "      抽学号"
      BeginProperty Font 
         Name            =   "华文细黑"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   4335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "抽学号-七二教学助手"
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
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3855
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
      Left            =   3720
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   200.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C9A70D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C9A70D&
      Height          =   855
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4560
      Width           =   4335
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   375
      Left            =   3840
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
      TabIndex        =   4
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "Number"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim xa As Single, ya As Single
Dim cs As String

Private Sub Form_Load()
cs = GetSetting("QEJXZS", "UserLike", "StuSum", 32) 'GetSetting(appname, section, key[, default])
b = 0
c = CStr(cs)
HScroll1.Value = c
Label2.Caption = "人数：" & HScroll1.Value
End Sub

Private Sub menu1_Click()
Me.Visible = False
Label1.Caption = 0
End Sub

Private Sub HScroll1_Change()
c = HScroll1.Value
Label2.Caption = "人数：" & HScroll1.Value
cs = str(c)
SaveSetting "QEJXZS", "UserLike", "StuSum", cs
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

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xa = X
ya = Y
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Me.Move Me.Left + X - xa, Me.Top + Y - ya
End Sub

Private Sub Label5_Click()
If Timer1.Enabled = False Then
c = HScroll1.Value
b = 0
Timer1.Enabled = True
Label5.Caption = "    加载中...."
HScroll1.Enabled = False
End If
End Sub

Private Sub Timer1_Timer()
Randomize
a = Int(Rnd * (c - 1 + 1)) + 1
Label1.Caption = a
b = b + 1
If b = 10 Then
Timer1.Enabled = False
Label5.Caption = "      抽学号"
HScroll1.Enabled = True
End If
End Sub
