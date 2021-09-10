VERSION 5.00
Begin VB.Form Home 
   BackColor       =   &H00332E2B&
   BorderStyle     =   0  'None
   ClientHeight    =   4545
   ClientLeft      =   6990
   ClientTop       =   4065
   ClientWidth     =   3405
   ControlBox      =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   4080
   End
   Begin VB.Label Label13 
      BackColor       =   &H00332E2B&
      Caption         =   "更新信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   4200
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   3240
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   2880
      TabIndex        =   9
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "七二教学助手-V2.1"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "  铃声（优化中）"
      BeginProperty Font 
         Name            =   "华文细黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "         计时器"
      BeginProperty Font 
         Name            =   "华文细黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "           黑屏"
      BeginProperty Font 
         Name            =   "华文细黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "       大组加分"
      BeginProperty Font 
         Name            =   "华文细黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "       小组加分"
      BeginProperty Font 
         Name            =   "华文细黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00C9A70D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C9A70D&
      Height          =   375
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00C9A70D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C9A70D&
      Height          =   375
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C9A70D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C9A70D&
      Height          =   375
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   2175
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C9A70D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C9A70D&
      Height          =   375
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "         抽学号"
      BeginProperty Font 
         Name            =   "华文细黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C9A70D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C9A70D&
      Height          =   375
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "V2.1"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "现在日期:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "现在时间:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C9A70D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C9A70D&
      FillColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   375
      Left            =   2880
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
      TabIndex        =   10
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim YN As String

Dim xa As Single, ya As Single

Private Sub Form_Load()
With nidProgramData
.cbSize = Len(nidProgramData)
.hwnd = Me.hwnd
.uID = vbNull
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
.uCallbackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon
.szTip = "七二教学助手-V2.1(单击恢复窗口)" & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, nidProgramData
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Form_MouseMove_err:
Dim Result, MSG As Long, I As Integer
If Me.ScaleMode = vbPixels Then
MSG = X
Else
MSG = X / Screen.TwipsPerPixelX
End If
Select Case MSG
Case WM_LBUTTONUP
SetForegroundWindow Me.hwnd '这个函数用来当你不或得焦点时弹出菜单能自动消失
Me.Show
Case WM_LBUTTONDOWN '双击托盘
SetForegroundWindow Me.hwnd
Me.Show
End Select
Exit Sub
Form_MouseMove_err:
End Sub
Private Sub Form_Unload(Cancel As Integer)
Dim blnExitMe As Boolean
If blnExitMe = False Then
Cancel = True '取消退出
Me.Hide
End If
End Sub
Private Sub MnuQuit_Click() '单击退出按钮时
Shell_NotifyIcon NIM_DELETE, nidProgramData
End
End Sub
'************************************************



Private Sub Label10_Click()
'Ring.Visible = True
End Sub

Private Sub Label10_DblClick()
MsgBox "此功能仍在开发！", vbInformation, "七二教学助手"
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

Private Sub Label13_Click()
MsgBox "七二教学助手-V2.1更新内容：" & vbCrLf & "-修复了因为大屏幕缺少“方正兰亭超细黑简体”导致的关闭叉号消失BUG；" & vbCrLf & "-贴心更新：抽学号修改人数增加标识标签，改人数更加方便。", , "更新内容"
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xa = X
ya = Y
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Me.Move Me.Left + X - xa, Me.Top + Y - ya
End Sub

Private Sub Label5_Click()
Smallplus.Visible = True
End Sub

Private Sub Label6_Click()
Number.Visible = True
End Sub

Private Sub Label7_Click()
Bigplus.Visible = True
End Sub

Private Sub Label8_Click()
Form1.Visible = True
End Sub

Private Sub Label9_Click()
Stopclock.Visible = True
End Sub

Private Sub Timer1_Timer()
Label1.Caption = "现在时间:" & Time
Label2.Caption = "现在日期:" & Date
'If Label1.Caption = "现在时间:13:48:40" Then
'Form1.Visible = True
'Label8.Caption = "     安静练字"
'Form1.Label1.ForeColor = RGB(255, 255, 255)
'End If
'If Label1.Caption = "现在时间:13:58:45" Then
'Form1.Visible = False
'Label8.Caption = "        黑屏"
'Form1.Label1.ForeColor = RGB(0, 0, 0)
'End If
End Sub

