VERSION 5.00
Begin VB.Form Smallplus 
   BackColor       =   &H00332E2B&
   BorderStyle     =   0  'None
   ClientHeight    =   6405
   ClientLeft      =   4230
   ClientTop       =   1920
   ClientWidth     =   9465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox label11 
      Height          =   375
      Left            =   7200
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   7560
      Width           =   1815
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00332E2B&
      BorderStyle     =   0  'None
      Caption         =   "加分组"
      Height          =   2295
      Left            =   120
      TabIndex        =   16
      Top             =   5520
      Width           =   9015
      Begin VB.OptionButton Option7 
         BackColor       =   &H00332E2B&
         Caption         =   "七组"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   3120
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   5760
         PasswordChar    =   "*"
         TabIndex        =   34
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00332E2B&
         Caption         =   "八组"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   3120
         TabIndex        =   24
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00332E2B&
         Caption         =   "六组"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00332E2B&
         Caption         =   "五组"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00332E2B&
         Caption         =   "四组"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1080
         TabIndex        =   20
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00332E2B&
         Caption         =   "三组"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1080
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00332E2B&
         Caption         =   "二组"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00332E2B&
         Caption         =   "一组"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   " +"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   4080
         TabIndex        =   35
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   " -"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   4920
         TabIndex        =   36
         Top             =   240
         Width           =   735
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H00C9A70D&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C9A70D&
         FillColor       =   &H00C9A70D&
         Height          =   495
         Left            =   4920
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   735
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H00C9A70D&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C9A70D&
         FillColor       =   &H00C9A70D&
         Height          =   495
         Left            =   4080
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "  登录"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   7200
         TabIndex        =   38
         Top             =   240
         Width           =   1695
      End
      Begin VB.Shape Shape12 
         BackColor       =   &H00C9A70D&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C9A70D&
         FillColor       =   &H00C9A70D&
         Height          =   495
         Left            =   7200
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00332E2B&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   2400
      TabIndex        =   14
      Top             =   720
      Width           =   2175
      Begin VB.Shape Shape2 
         BorderColor     =   &H00E0E0E0&
         Height          =   1455
         Left            =   120
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   80.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "二组"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   600
         TabIndex        =   15
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00332E2B&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   2175
      Begin VB.Shape Shape1 
         BorderColor     =   &H00E0E0E0&
         Height          =   1455
         Left            =   120
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   80.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "一组"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   0
         Left            =   600
         TabIndex        =   13
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00332E2B&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   6960
      TabIndex        =   10
      Top             =   3120
      Width           =   2175
      Begin VB.Shape Shape8 
         BorderColor     =   &H00E0E0E0&
         Height          =   1455
         Left            =   120
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackColor       =   &H00404040&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   80.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "八组"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   6
         Left            =   600
         TabIndex        =   11
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00332E2B&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   4680
      TabIndex        =   8
      Top             =   3120
      Width           =   2175
      Begin VB.Shape Shape7 
         BorderColor     =   &H00E0E0E0&
         Height          =   1455
         Left            =   120
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackColor       =   &H00404040&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   80.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "七组"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   5
         Left            =   600
         TabIndex        =   9
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00332E2B&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   2400
      TabIndex        =   6
      Top             =   3120
      Width           =   2175
      Begin VB.Shape Shape6 
         BorderColor     =   &H00E0E0E0&
         Height          =   1455
         Left            =   120
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackColor       =   &H00404040&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   80.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "六组"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   4
         Left            =   600
         TabIndex        =   7
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00332E2B&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   2175
      Begin VB.Shape Shape5 
         BorderColor     =   &H00E0E0E0&
         Height          =   1455
         Left            =   120
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404040&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   80.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "五组"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   3
         Left            =   600
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00332E2B&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   6960
      TabIndex        =   2
      Top             =   720
      Width           =   2175
      Begin VB.Shape Shape4 
         BorderColor     =   &H00E0E0E0&
         Height          =   1455
         Left            =   120
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H00404040&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   80.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "四组"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   2
         Left            =   600
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00332E2B&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   4680
      TabIndex        =   0
      Top             =   720
      Width           =   2175
      Begin VB.Shape Shape3 
         BorderColor     =   &H00E0E0E0&
         Height          =   1455
         Left            =   120
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   80.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "三组"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   1
         Left            =   600
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
   End
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
      Left            =   8880
      TabIndex        =   41
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape Shape13 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   375
      Left            =   9000
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label16 
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
      Left            =   0
      TabIndex        =   39
      Top             =   120
      Width           =   3855
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H00C9A70D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C9A70D&
      FillColor       =   &H00C9A70D&
      Height          =   495
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   " +"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6600
      TabIndex        =   37
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label Label17 
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
      TabIndex        =   40
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "Smallplus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a1 As Integer
Dim b1 As String

Dim a2 As Integer
Dim b2 As Integer

Dim a3 As Integer
Dim b3 As Integer

Dim a4 As Integer
Dim b4 As Integer

Dim a5 As Integer
Dim b5 As Integer

Dim a6 As Integer
Dim b6 As Integer

Dim a7 As Integer
Dim b7 As Integer

Dim a8 As Integer
Dim b8 As Integer
Dim pp As String

Dim str As Boolean

Dim xa As Single, ya As Single


Private Sub Form_Load()
b1 = GetSetting("QEJXZS", "Stu", "1", 0)
b2 = GetSetting("QEJXZS", "Stu", "2", 0)
b3 = GetSetting("QEJXZS", "Stu", "3", 0)
b4 = GetSetting("QEJXZS", "Stu", "4", 0)
b5 = GetSetting("QEJXZS", "Stu", "5", 0)
b6 = GetSetting("QEJXZS", "Stu", "6", 0)
b7 = GetSetting("QEJXZS", "Stu", "7", 0)
b8 = GetSetting("QEJXZS", "Stu", "8", 0)
a1 = CInt(b1)
a2 = CInt(b2)
a3 = CInt(b3)
a4 = CInt(b4)
a5 = CInt(b5)
a6 = CInt(b6)
a7 = CInt(b7)
a8 = CInt(b8)
Label2.Caption = a1
Label3.Caption = a2
Label5.Caption = a3
Label6.Caption = a4
Label7.Caption = a5
Label8.Caption = a6
Label9.Caption = a7
Label10.Caption = a8
Label11.Text = a1 & "," & a2 & "," & a3 & "," & a4 & "," & a5 & "," & a6 & "," & a7 & "," & a8
Close #1
str = False
End Sub



Private Sub Label12_Click()
If str = False Then
 MsgBox "请先登录！", vbExclamation, "七二教学助手"
Else
If Option1.Value = False And Option2.Value = False And Option3.Value = False And Option4.Value = False And Option5.Value = False And Option6.Value = False And Option7.Value = False And Option8.Value = False Then
    MsgBox "请至少选择一个组!", vbExclamation, "警告"
Else
    If Option1.Value = True Then
        a1 = a1 + 1
        Label2.Caption = a1
    End If
    If Option2.Value = True Then
        a2 = a2 + 1
        Label3.Caption = a2
    End If
    If Option3.Value = True Then
        a3 = a3 + 1
        Label5.Caption = a3
    End If
    If Option4.Value = True Then
        a4 = a4 + 1
        Label6.Caption = a4
    End If
    If Option5.Value = True Then
        a5 = a5 + 1
        Label7.Caption = a5
    End If
    If Option6.Value = True Then
        a6 = a6 + 1
        Label8.Caption = a6
    End If
    If Option7.Value = True Then
        a7 = a7 + 1
        Label9.Caption = a7
    End If
    If Option8.Value = True Then
        a8 = a8 + 1
        Label10.Caption = a8
    End If
    Label11.Text = a1 & "," & a2 & "," & a3 & "," & a4 & "," & a5 & "," & a6 & "," & a7 & "," & a8
    Open "D:\start\text.scrt" For Output As #2
    Print #2, Label11.Text
    Close #2
End If
End If
End Sub

Private Sub Label13_Click()
If str = False Then
 MsgBox "请先登录！", vbExclamation, "七二教学助手"
Else
If Option1.Value = False And Option2.Value = False And Option3.Value = False And Option4.Value = False And Option5.Value = False And Option6.Value = False And Option7.Value = False And Option8.Value = False Then
    MsgBox "请至少选择一个组!", vbExclamation, "七二教学助手"
Else
    If Option1.Value = True Then
        a1 = a1 - 1
        Label2.Caption = a1
    End If
    If Option2.Value = True Then
        a2 = a2 - 1
        Label3.Caption = a2
    End If
    If Option3.Value = True Then
        a3 = a3 - 1
        Label5.Caption = a3
    End If
    If Option4.Value = True Then
        a4 = a4 - 1
        Label6.Caption = a4
    End If
    If Option5.Value = True Then
        a5 = a5 - 1
        Label7.Caption = a5
    End If
    If Option6.Value = True Then
        a6 = a6 - 1
        Label8.Caption = a6
    End If
    If Option7.Value = True Then
        a7 = a7 - 1
        Label9.Caption = a7
    End If
    If Option8.Value = True Then
        a8 = a8 - 1
        Label10.Caption = a8
    End If
    Label11.Text = a1 & "," & a2 & "," & a3 & "," & a4 & "," & a5 & "," & a6 & "," & a7 & "," & a8
    Open "D:\start\text.scrt" For Output As #2
    Print #2, Label11.Text
    Close #2
End If
End If
End Sub

Private Sub Label13_DblClick()
pp = MsgBox("是否清除?清除后将无法恢复!", vbYesNo + vbExclamation, "警告")
 If Option1.Value = True Then
        a1 = a1 + 1
        Label2.Caption = a1
    End If
    If Option2.Value = True Then
        a2 = a2 + 1
        Label3.Caption = a2
    End If
    If Option3.Value = True Then
        a3 = a3 + 1
        Label5.Caption = a3
    End If
    If Option4.Value = True Then
        a4 = a4 + 1
        Label6.Caption = a4
    End If
    If Option5.Value = True Then
        a5 = a5 + 1
        Label7.Caption = a5
    End If
    If Option6.Value = True Then
        a6 = a6 + 1
        Label8.Caption = a6
    End If
    If Option7.Value = True Then
        a7 = a7 + 1
        Label9.Caption = a7
    End If
    If Option8.Value = True Then
        a8 = a8 + 1
        Label10.Caption = a8
    End If
    Label11.Text = a1 & "," & a2 & "," & a3 & "," & a4 & "," & a5 & "," & a6 & "," & a7 & "," & a8
    Open "D:\start\text.scrt" For Output As #2
    Print #2, Label11.Text
    Close #2
If pp = vbYes Then
a1 = 0
a2 = 0
a3 = 0
a4 = 0
a5 = 0
a6 = 0
a7 = 0
a8 = 0
b1 = "0"
b2 = "0"
b3 = "0"
b4 = "0"
b5 = "0"
b6 = "0"
b7 = "0"
b8 = "0"
Label2.Caption = a1
Label3.Caption = a2
Label5.Caption = a3
Label6.Caption = a4
Label7.Caption = a5
Label8.Caption = a6
Label9.Caption = a7
Label10.Caption = a8
 Label11.Text = a1 & "," & a2 & "," & a3 & "," & a4 & "," & a5 & "," & a6 & "," & a7 & "," & a8
    Open "D:\start\text.scrt" For Output As #2
    Print #2, Label11.Text
    Close #2
End If
End Sub

Private Sub Label15_Click()
If str = False Then
    If Text1.Text = "qe72" Or Text1.Text = "qejxzs" Then
    Text1.Enabled = False
    str = True
    Label15.Caption = "  注销"
    Else
    MsgBox "密码错误！", vbExclamation, "警告"
    End If
Else
    str = False
    Text1.Enabled = True
    Text1.Text = ""
    Label15.Caption = "  登录"
End If
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xa = X
ya = Y
End Sub

Private Sub Label16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Me.Move Me.Left + X - xa, Me.Top + Y - ya
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xa = X
ya = Y
End Sub

Private Sub Label17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Me.Move Me.Left + X - xa, Me.Top + Y - ya
End Sub

Private Sub Label26_Click()
Me.Visible = False
End Sub
