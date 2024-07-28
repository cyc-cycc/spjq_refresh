VERSION 5.00
Begin VB.Form gap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "随机间隔"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3495
   Icon            =   "gap.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   5  'Size
   ScaleHeight     =   3150
   ScaleWidth      =   3495
   Begin VB.CommandButton Command3 
      Caption         =   "开始刷屏（请手动复制刷屏内容）"
      Height          =   495
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   8
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "数据设置"
      Height          =   1815
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   10
      Top             =   120
      Width           =   3255
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         MaxLength       =   4
         MousePointer    =   3  'I-Beam
         TabIndex        =   5
         Text            =   "50"
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   "+"
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Caption         =   "-"
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   1320
         MaxLength       =   3
         MousePointer    =   3  'I-Beam
         TabIndex        =   4
         Text            =   "0.5"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   480
         MaxLength       =   3
         MousePointer    =   3  'I-Beam
         TabIndex        =   3
         Text            =   "0.1"
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000004&
         Caption         =   "+"
         Height          =   255
         Left            =   1800
         MousePointer    =   1  'Arrow
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000004&
         Caption         =   "-"
         Height          =   255
         Left            =   2280
         MousePointer    =   1  'Arrow
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox num 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         MaxLength       =   5
         MousePointer    =   3  'I-Beam
         TabIndex        =   0
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "刷屏延迟："
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "毫秒"
         Height          =   255
         Left            =   2760
         TabIndex        =   17
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "在"
         Height          =   255
         Left            =   240
         MousePointer    =   1  'Arrow
         TabIndex        =   16
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "秒 间取随机数"
         Height          =   255
         Left            =   1920
         MousePointer    =   1  'Arrow
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "到"
         Height          =   255
         Left            =   1080
         MousePointer    =   1  'Arrow
         TabIndex        =   14
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "刷屏间隔："
         Height          =   255
         Left            =   120
         MousePointer    =   1  'Arrow
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000004&
         Caption         =   "刷屏次数："
         Height          =   255
         Left            =   120
         MousePointer    =   1  'Arrow
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Caption         =   "次"
         Height          =   255
         Left            =   2880
         MousePointer    =   1  'Arrow
         TabIndex        =   11
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.CommandButton exit 
      Caption         =   "关闭"
      Height          =   375
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   9
      Top             =   2640
      Width           =   3255
   End
End
Attribute VB_Name = "gap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Me.num.Text = Me.num.Text + 1
    If Me.num.Text <= 0 Then
        Me.num.Text = 1
    End If
End Sub
Private Sub Command2_Click()
    Me.num.Text = Me.num.Text - 1
    If Me.num.Text <= 0 Then
        Me.num.Text = 1
    End If
End Sub
Private Sub Command3_Click()
    gif_gap = False
    Me.Hide
    special_gap.Show
End Sub
Private Sub Command4_Click()
    Me.Text3.Text = Me.Text3.Text - 10
    If Me.Text3.Text < 0 Then
        Me.Text3.Text = 0
    End If
End Sub
Private Sub Command5_Click()
    Me.Text3.Text = Me.Text3.Text + 10
    If Me.Text3.Text < 0 Then
        Me.Text3.Text = 0
    End If
End Sub
Private Sub exit_Click()
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
        If gif_gap <> True Then
            Call rgn_gif(Me, 25, 3, 3, 3)
            gif_gap = True
        End If
    End If
    Call rgnform(Me, rou, rou, 25, 3, 3, 3)
End Sub
Private Sub Form_Load()
    If sfly = 1 And sfl = "all" Then
        Me.left = GetSetting(App.ProductName, "location", "left_gap", Screen.Width / 2 - Me.Width / 2)
        Me.top = GetSetting(App.ProductName, "location", "top_gap", Screen.Height / 2 - Me.Height / 2)
    Else
        Me.left = Screen.Width / 2 - Me.Width / 2
        Me.top = Screen.Height / 2 - Me.Height / 2
    End If
    Me.num.Text = Form1.num.Text
    Me.Text3.Text = Form1.Text2.Text
    Me.Text1.Text = GetSetting(App.ProductName, "settings", "gap_small", 0.1)
    Me.Text2.Text = GetSetting(App.ProductName, "settings", "gap_big", 0.5)
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sx = X
    sy = Y
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Me.left = Me.left + (X - sx)
        Me.top = Me.top + (Y - sy)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    gif_gap = False
    Form1.Show
    SaveSetting App.ProductName, "location", "left_gap", Me.left
    SaveSetting App.ProductName, "location", "top_gap", Me.top
    SaveSetting App.ProductName, "settings", "gap_small", Me.Text1.Text
    SaveSetting App.ProductName, "settings", "gap_big", Me.Text2.Text
    DeleteObject outrgn
End Sub
Private Sub num_LostFocus()
    Me.num.Text = Fix(Me.num.Text)
    If Me.num.Text <= 0 Then
        Me.num.Text = 1
    End If
End Sub
