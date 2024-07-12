VERSION 5.00
Begin VB.Form special_gap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "引导_特殊"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4470
   Icon            =   "special.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   2430
   ScaleWidth      =   4470
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "50"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "0~0"
      Top             =   1440
      Width           =   3375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "取消"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ENTER"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CTRL+ENTER"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1380
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ALT+S"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   960
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "刷屏期间不要点击其它任何窗口！"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "刷屏延迟："
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "刷屏间隔："
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "请选择消息发送方式:"
      Height          =   255
      Left            =   240
      MousePointer    =   1  'Arrow
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "刷屏次数:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "special_gap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Me.Hide
    MsgBox "请点击要刷屏的文本框", 4096, "引导"
    For i = 1 To Me.Text1.Text
        Randomize
        w = Int((gap.Text2.Text - gap.Text1.Text + 1) * Rnd + gap.Text1.Text) * 1000
        SendKeys "^v"
        Call Sleep(Me.Text3.Text)
        SendKeys "{ENTER}"
        Call Sleep(w)
    Next
    Call Sleep(500)
    gap.Show
    Unload Me
End Sub
Private Sub Command2_Click()
    Me.Hide
    MsgBox "请点击要刷屏的文本框", 4096, "引导"
    For i = 1 To Me.Text1.Text
        Randomize
        w = Int((gap.Text2.Text - gap.Text1.Text + 1) * Rnd + gap.Text1.Text) * 1000
        SendKeys "^v"
        Call Sleep(Me.Text3.Text)
        SendKeys "^{ENTER}"
        Call Sleep(w)
    Next
    Call Sleep(500)
    gap.Show
    Unload Me
End Sub
Private Sub Command3_Click()
    Me.Hide
    MsgBox "请点击要刷屏的文本框", 4096, "引导"
    For i = 1 To Me.Text1.Text
        Randomize
        w = Int((gap.Text2.Text - gap.Text1.Text + 1) * Rnd + gap.Text1.Text) * 1000
        SendKeys "^v"
        Call Sleep(Me.Text3.Text)
        SendKeys "%s"
        Call Sleep(w)
    Next
    Call Sleep(500)
    gap.Show
    Unload Me
End Sub
Private Sub Command4_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    Dim rtn As Long
    rtn = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, aero, LWA_ALPHA
    If gif = True Then
        Call rgn_gif(Me, 25, 3, 3, 3)
    End If
    Call rgnform(Me, rou, rou, 25, 3, 3, 3)
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
End Sub
Private Sub Form_Load()
    Me.Text1.Text = gap.num.Text
    Me.Text2.Text = "在 " & gap.Text1.Text & " 到 " & gap.Text2.Text & " 间取随机数"
    Me.Text3.Text = gap.Text3.Text
End Sub
Private Sub Form_Unload(Cancel As Integer)
    gap.Show
    DeleteObject outrgn
End Sub
