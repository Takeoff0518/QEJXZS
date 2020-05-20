VERSION 5.00
Begin VB.Form Stopclock 
   BorderStyle     =   0  'None
   ClientHeight    =   5880
   ClientLeft      =   5925
   ClientTop       =   2790
   ClientWidth     =   8700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox Check1 
      BackColor       =   &H00332E2B&
      Caption         =   "结束提示音"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   480
      MaskColor       =   &H00000000&
      TabIndex        =   0
      Top             =   4080
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8040
      Top             =   4080
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00332E2B&
      Height          =   7695
      Left            =   0
      ScaleHeight     =   7635
      ScaleWidth      =   14715
      TabIndex        =   1
      Top             =   0
      Width           =   14775
      Begin VB.Label Label26 
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
         Left            =   8040
         TabIndex        =   27
         Top             =   0
         Width           =   495
      End
      Begin VB.Shape Shape22 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         Height          =   375
         Left            =   8160
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "小组加分-七二教学助手"
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
         Left            =   0
         TabIndex        =   26
         Top             =   120
         Width           =   3855
      End
      Begin VB.Label Label24 
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
         Left            =   -480
         TabIndex        =   25
         Top             =   0
         Width           =   9495
      End
      Begin VB.Label Command15 
         BackStyle       =   0  'Transparent
         Caption         =   " 停止"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   42
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   5760
         TabIndex        =   24
         Top             =   4560
         Width           =   2655
      End
      Begin VB.Shape Shape21 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   5760
         Shape           =   4  'Rounded Rectangle
         Top             =   4560
         Width           =   2655
      End
      Begin VB.Label Command14 
         BackStyle       =   0  'Transparent
         Caption         =   " 暂停"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   42
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   3000
         TabIndex        =   23
         Top             =   4560
         Width           =   2655
      End
      Begin VB.Shape Shape20 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   3000
         Shape           =   4  'Rounded Rectangle
         Top             =   4560
         Width           =   2655
      End
      Begin VB.Label Command13 
         BackStyle       =   0  'Transparent
         Caption         =   " 开始"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   42
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   240
         TabIndex        =   22
         Top             =   4560
         Width           =   2655
      End
      Begin VB.Shape Shape19 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   4560
         Width           =   2655
      End
      Begin VB.Shape Shape18 
         BorderColor     =   &H00E0E0E0&
         Height          =   1935
         Left            =   7080
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Shape Shape17 
         BorderColor     =   &H00E0E0E0&
         Height          =   1935
         Left            =   5880
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Shape Shape16 
         BorderColor     =   &H00E0E0E0&
         Height          =   1935
         Left            =   4320
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Shape Shape15 
         BorderColor     =   &H00E0E0E0&
         Height          =   1935
         Left            =   3120
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Shape Shape14 
         BorderColor     =   &H00E0E0E0&
         Height          =   1935
         Left            =   1560
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Shape Shape13 
         BorderColor     =   &H00E0E0E0&
         Height          =   1935
         Left            =   360
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Command6 
         BackStyle       =   0  'Transparent
         Caption         =   " ↑"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   7080
         TabIndex        =   21
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Command5 
         BackStyle       =   0  'Transparent
         Caption         =   " ↑"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   5880
         TabIndex        =   20
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Command3 
         BackStyle       =   0  'Transparent
         Caption         =   " ↑"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   4320
         TabIndex        =   19
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Command4 
         BackStyle       =   0  'Transparent
         Caption         =   " ↑"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   3120
         TabIndex        =   18
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Command2 
         BackStyle       =   0  'Transparent
         Caption         =   " ↑"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   1560
         TabIndex        =   17
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Command1 
         BackStyle       =   0  'Transparent
         Caption         =   " ↑"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   360
         TabIndex        =   16
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Command7 
         BackStyle       =   0  'Transparent
         Caption         =   " ↓"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   7080
         TabIndex        =   15
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Command8 
         BackStyle       =   0  'Transparent
         Caption         =   " ↓"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   5880
         TabIndex        =   14
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Command10 
         BackStyle       =   0  'Transparent
         Caption         =   " ↓"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   4320
         TabIndex        =   13
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Command9 
         BackStyle       =   0  'Transparent
         Caption         =   " ↓"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   3120
         TabIndex        =   12
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Command11 
         BackStyle       =   0  'Transparent
         Caption         =   " ↓"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   1560
         TabIndex        =   11
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Command12 
         BackStyle       =   0  'Transparent
         Caption         =   " ↓"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Shape Shape12 
         BackColor       =   &H00C9A70D&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C9A70D&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C9A70D&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   7080
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   1095
      End
      Begin VB.Shape Shape11 
         BackColor       =   &H00C9A70D&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C9A70D&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C9A70D&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   5880
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   1095
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H00C9A70D&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C9A70D&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C9A70D&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   4320
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   1095
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H00C9A70D&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C9A70D&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C9A70D&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   3120
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   1095
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H00C9A70D&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C9A70D&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C9A70D&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   1560
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   1095
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00C9A70D&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C9A70D&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C9A70D&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   1095
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00C9A70D&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C9A70D&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C9A70D&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   7080
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00C9A70D&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C9A70D&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C9A70D&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   5880
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00C9A70D&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C9A70D&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C9A70D&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   4320
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00C9A70D&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C9A70D&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C9A70D&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   3120
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00C9A70D&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C9A70D&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C9A70D&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   1560
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C9A70D&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C9A70D&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C9A70D&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   99.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1935
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   99.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1935
         Left            =   1560
         TabIndex        =   8
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   99.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1935
         Left            =   3120
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404040&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   99.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1935
         Left            =   4320
         TabIndex        =   6
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   99.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1935
         Left            =   5880
         TabIndex        =   5
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00404040&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   99.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1935
         Left            =   7080
         TabIndex        =   4
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Height          =   1335
         Left            =   2640
         TabIndex        =   3
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Height          =   1215
         Left            =   5400
         TabIndex        =   2
         Top             =   1680
         Width           =   375
      End
   End
End
Attribute VB_Name = "Stopclock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
    Option Explicit
    Dim s1 As Integer
    Dim s2 As Integer
    Dim m1 As Integer
    Dim m2 As Integer
    Dim h1 As Integer
    Dim h2 As Integer
    Dim xa As Single, ya As Single
    
Private Sub Check1_Click()
If Check1.Value = 0 Then
    SaveSetting "QEJXZS", "UserLike", "StopClockRing", "0"
End If

If Check1.Value = 1 Then
    SaveSetting "QEJXZS", "UserLike", "StopClockRing", "1"
End If

End Sub

Private Sub Command1_Click()   'h2+
    Command12.Enabled = True
    h2 = h2 + 1
    Label1.Caption = h2
    If h2 >= 5 Then
        Command1.Enabled = False
    End If
End Sub

Private Sub Command10_Click()   'm1-
    Command3.Enabled = True
    m1 = m1 - 1
    Label4.Caption = m1
    If m1 <= 0 Then
        Command10.Enabled = False
    End If
End Sub

Private Sub Command11_Click()   'h1-
    Command2.Enabled = True
    h1 = h1 - 1
    Label2.Caption = h1
    If h1 <= 0 Then
        Command11.Enabled = False
    End If
End Sub

Private Sub Command12_Click()   'h2-
    Command1.Enabled = True
    h2 = h2 - 1
    Label1.Caption = h2
    If h2 <= 0 Then
        Command12.Enabled = False
    End If
End Sub

Private Sub Command13_Click()   'Start
    If s1 = 0 And s2 = 0 And m1 = 0 And m2 = 0 And h1 = 0 And h2 = 0 Then
        MsgBox "时间不能为0！", vbExclamation, "警告"
        Else
    Timer1.Enabled = True
    Check1.Enabled = False
    Command13.Enabled = False
    Command14.Enabled = True
    Command15.Enabled = False
    
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = False
    Command8.Enabled = False
    Command9.Enabled = False
    Command10.Enabled = False
    Command11.Enabled = False
    Command12.Enabled = False
    End If
End Sub

Private Sub Command14_Click()   'Loading
    Command15.Enabled = True
    Command14.Enabled = False
    Timer1.Enabled = False
    Command13.Enabled = True
End Sub

Private Sub Command15_Click()   'Stop
    s1 = 0
    s2 = 0
    m1 = 0
    m2 = 0
    h1 = 0
    h2 = 0
    Label1.Caption = h2
    Label2.Caption = h1
    Label3.Caption = m2
    Label4.Caption = m1
    Label5.Caption = s2
    Label6.Caption = s1
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command15.Enabled = False
    Command14.Enabled = False
    Command13.Enabled = True
     Check1.Enabled = True
End Sub

Private Sub Command2_Click()   'h1+
    Command11.Enabled = True
    h1 = h1 + 1
    Label2.Caption = h1
    If h1 >= 9 Then
        Command2.Enabled = False
    End If
End Sub

Private Sub Command3_Click()   'm1+
    Command10.Enabled = True
    m1 = m1 + 1
    Label4.Caption = m1
    If m1 >= 9 Then
        Command3.Enabled = False
    End If
End Sub

Private Sub Command4_Click()   'm2+
    Command9.Enabled = True
    m2 = m2 + 1
    Label3.Caption = m2
    If m2 >= 5 Then
        Command4.Enabled = False
    End If
End Sub

Private Sub Command5_Click()    's2+
    Command8.Enabled = True
    s2 = s2 + 1
    Label5.Caption = s2
    If s2 >= 5 Then
    Command5.Enabled = False
    End If
End Sub

Private Sub Command6_Click()    's1+
    Command7.Enabled = True
    s1 = s1 + 1
    Label6.Caption = s1
    If s1 >= 9 Then
        Command6.Enabled = False
    End If
End Sub

Private Sub Command7_Click()    's1-
    Command6.Enabled = True
    s1 = s1 - 1
    Label6.Caption = s1
    If s1 <= 0 Then
        Command7.Enabled = False
    End If
End Sub

Private Sub Command8_Click()    's2-
    Command5.Enabled = True
    s2 = s2 - 1
    Label5.Caption = s2
    If s2 <= 0 Then
        Command8.Enabled = False
    End If
End Sub

Private Sub Command9_Click()    'm2-
    Command4.Enabled = True
    m2 = m2 - 1
    Label3.Caption = m2
    If m2 <= 0 Then
        Command9.Enabled = False
    End If
End Sub

Private Sub Form_Load()     'Load'
If GetSetting("QEJXZS", "UserLike", "StopClockRing", 1) = "0" Then
    Check1.Value = 0
End If
If GetSetting("QEJXZS", "UserLike", "StopClockRing", 1) = "1" Then
    Check1.Value = 1
End If
    Me.Visible = False
    Home.Visible = True
    s1 = 0
    s2 = 0
    m1 = 0
    m2 = 0
    h1 = 0
    h2 = 0
End Sub

Private Sub Menu_1_Click()  'End
Me.Visible = False
Home.Visible = True
End Sub






Private Sub Label24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xa = X
ya = Y
End Sub

Private Sub Label24_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Me.Move Me.Left + X - xa, Me.Top + Y - ya
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xa = X
ya = Y
End Sub

Private Sub Label25_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Me.Move Me.Left + X - xa, Me.Top + Y - ya
End Sub

Private Sub Label26_Click()
Me.Visible = False
End Sub





Private Sub Timer1_Timer()
s1 = s1 - 1
If s1 = -1 Then '秒
    s1 = 9
    s2 = s2 - 1
End If
If s2 = -1 Then
    s2 = 5
    m1 = m1 - 1
End If

If m1 = -1 Then
    m1 = 9
    m2 = m2 - 1
End If
If m2 = -1 Then
    m2 = 5
    h1 = h1 - 1
End If
If h1 = -1 Then
    h1 = 9
    h2 = h2 - 1
End If
Label1.Caption = h2
Label2.Caption = h1
Label3.Caption = m2
Label4.Caption = m1
Label5.Caption = s2
Label6.Caption = s1
If s1 = 0 And s2 = 0 And m1 = 0 And m2 = 0 And h1 = 0 And h2 = 0 Then
        
        Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    s1 = 0
    s2 = 0
    m1 = 0
    m2 = 0
    h1 = 0
    h2 = 0
     Command15.Enabled = False
    Command14.Enabled = False
    Command13.Enabled = True
    Timer1.Enabled = False
    Check1.Enabled = True
     If Check1.Value = 1 Then
    
        Beep 523.3, 450   '~1
        Beep 523.3, 450   '~1
        Beep 523.3, 450   '~1
        Beep 523.3, 450   '~1
        Beep 523.3, 450   '~1
        Beep 523.3, 450   '~1
    End If
     MsgBox "时间到！", , "提示"
    End If
   
End Sub
