VERSION 5.00
Begin VB.Form Ring 
   BorderStyle     =   0  'None
   Caption         =   "广播"
   ClientHeight    =   315
   ClientLeft      =   22965
   ClientTop       =   4725
   ClientWidth     =   1245
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   315
   ScaleWidth      =   1245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1200
      Top             =   0
   End
   Begin VB.CheckBox Check1 
      Caption         =   "铃声开启"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.Menu menu1 
      Caption         =   "最小化(&X)"
   End
End
Attribute VB_Name = "Ring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Sub Check1_Click()
If Check1.Value = 1 Then
Timer1.Enabled = True
Else
Timer1.Enabled = False
End If
End Sub

Private Sub menu1_Click()
Me.Visible = False
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Time
If Label1.Caption = "7:28:45" Or Label1.Caption = "8:19:45" Or Label1.Caption = "9:09:45" Or Label1.Caption = "10:19:45" Or Label1.Caption = "11:09:45" Or Label1.Caption = "13:59:45" Or Label1.Caption = "14:49:45" Or Label1.Caption = "15:59:45" Or Label1.Caption = "16:49:45" Or Label1.Caption = "17:44:45" Then
Beep 262, 550      '1
        Beep 329.7, 550  '3
        Beep 392, 550    '5
        Beep 329.7, 550  '3
        Beep 392, 550    '5
        Beep 523.3, 550   '~1
End If
If Label1.Caption = "8:09:45" Or Label1.Caption = "8:59:45" Or Label1.Caption = "9:49:45" Or Label1.Caption = "10:59:45" Or Label1.Caption = "11:54:45" Or Label1.Caption = "14:39:45" Or Label1.Caption = "15:29:45" Or Label1.Caption = "16:39:45" Or Label1.Caption = "17:29:45" Or Label1.Caption = "18:29:45" Then
Beep 523.3, 550      '~1
        Beep 392, 550  '5
        Beep 329.7, 550    '3
        Beep 392, 550    '5
        Beep 329.7, 550  '3
        Beep 262, 550   '~1
End If
End Sub
