VERSION 5.00
Begin VB.Form special_quickest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "引导_特殊"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4455
   Icon            =   "special_quickest.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4455
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command4 
      Caption         =   "取消"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   1440
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
      Width           =   3375
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
   Begin VB.Label Label5 
      Caption         =   "刷屏期间不要点击其它任何窗口！"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "刷屏次数:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "请选择消息发送方式:"
      Height          =   255
      Left            =   240
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "special_quickest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Me.Hide
    MsgBox "请点击要刷屏的文本框", 4096, "引导"
    For i = 1 To Me.Text1.Text
        SendKeys "^v"
        SendKeys "{ENTER}"
    Next
    Call Sleep(500)
    quickest.Show
    Unload Me
End Sub
Private Sub Command2_Click()
    Me.Hide
    MsgBox "请点击要刷屏的文本框", 4096, "引导"
    For i = 1 To Me.Text1.Text
        SendKeys "^v"
        SendKeys "^{ENTER}"
    Next
    Call Sleep(500)
    quickest.Show
    Unload Me
End Sub
Private Sub Command3_Click()
    Me.Hide
    MsgBox "请点击要刷屏的文本框", 4096, "引导"
    For i = 1 To Me.Text1.Text
        SendKeys "^v"
        SendKeys "%s"
    Next
    Call Sleep(500)
    quickest.Show
    Unload Me
End Sub
Private Sub Command4_Click()
    If gif = True Then
        aero_tmp = aero
        For i = 1 To aero_tmp / 9
            SetLayeredWindowAttributes hwnd, 0, aero_tmp, LWA_ALPHA
            aero_tmp = aero_tmp - 8
            Call Sleep(1)
        Next
    End If
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
    Me.Text1.Text = quickest.num.Text
End Sub
Private Sub Form_Unload(Cancel As Integer)
    quickest.Show
End Sub
