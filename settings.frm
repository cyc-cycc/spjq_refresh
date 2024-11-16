VERSION 5.00
Begin VB.Form settings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÉèÖÃ"
   ClientHeight    =   4695
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   3975
   Icon            =   "settings.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   5  'Size
   ScaleHeight     =   4695
   ScaleWidth      =   3975
   Begin VB.Frame Frame3 
      Caption         =   "±£´æÉèÖÃ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   19
      Top             =   2880
      Width           =   3735
      Begin VB.OptionButton Option4 
         Caption         =   "±£´æËùÓÐ"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton Option3 
         Caption         =   "½ö±£´æÖ÷´°¿Ú"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "ÍË³öÊ±±£´æ´°¿ÚÎ»ÖÃ"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000004&
      Caption         =   "ÖØÖÃ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      MousePointer    =   1  'Arrow
      TabIndex        =   12
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000004&
      Caption         =   "Ó¦ÓÃ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   11
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "ÏÔÊ¾Æô¶¯ÆÁÄ»"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      MousePointer    =   1  'Arrow
      TabIndex        =   7
      Top             =   2400
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Æô¶¯ÉèÖÃ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   17
      Top             =   1680
      Width           =   3735
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   2520
         MousePointer    =   3  'I-Beam
         TabIndex        =   6
         Text            =   "2000"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Æô¶¯ÆÁÄ»ÏÔÊ¾Ê±³¤£¨ms£©£º"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ÍË³ö"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      MousePointer    =   1  'Arrow
      TabIndex        =   13
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Caption         =   "Íâ¹ÛÉèÖÃ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   1455
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   14
      Top             =   120
      Width           =   3735
      Begin VB.CheckBox Check3 
         Caption         =   "ÆôÓÃ¶¯»­"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000004&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   4
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000004&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   5
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H80000004&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H80000004&
         Caption         =   "-"
         Height          =   255
         Left            =   3240
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   1920
         MaxLength       =   3
         MousePointer    =   3  'I-Beam
         TabIndex        =   0
         Text            =   "210"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   1920
         MaxLength       =   3
         MousePointer    =   3  'I-Beam
         TabIndex        =   3
         Text            =   "20"
         Top             =   705
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000004&
         Caption         =   "²»Í¸Ã÷¶È0~255£º"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   120
         MousePointer    =   1  'Arrow
         TabIndex        =   16
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000004&
         Caption         =   "Ô²½Ç´óÐ¡£º"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   120
         MousePointer    =   1  'Arrow
         TabIndex        =   15
         Top             =   720
         Width           =   975
      End
   End
End
Attribute VB_Name = "settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command10_Click()
    Me.Text1.Text = Me.Text1.Text + 1
    If Me.Text1.Text > 255 Then
        Me.Text1.Text = 255
    End If
End Sub
Private Sub Command11_Click()
    Me.Text1.Text = Me.Text1.Text - 1
    If Me.Text1.Text < 10 Then
        Me.Text1.Text = 10
    End If
End Sub
Private Sub Command2_Click()
    Me.Text2.Text = Me.Text2.Text + 1
End Sub
Private Sub Command1_Click()
    Me.Text2.Text = Me.Text2.Text - 1
    If Me.Text2.Text < 0 Then
        Me.Text2.Text = 0
    End If
End Sub
Private Sub Command3_Click()
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
Private Sub Command5_Click()
    aero = Me.Text1.Text
    rou = Me.Text2.Text
    Dim rtn As Long
    rtn = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes Me.hwnd, 0, aero, LWA_ALPHA
'    Call rgnform(Me, rou, rou, 25, 3, 3, 3)
    splt = Me.Text3.Text
    spl = Me.Check1.Value
    If Me.Option3.Value = True Then sfl = "o_main" Else sfl = "all"
    sfly = Me.Check2.Value
    If Me.Check3.Value = 1 Then
        gif = True
        Call rgn_gif(Me, 25, 3, 3, 3)
    Else
        gif = False
    End If
End Sub
Private Sub Command6_Click()
    Me.Text1.Text = 210
    Me.Text2.Text = 20
    rou = Me.Text2.Text
    aero = Me.Text1.Text
    Dim rtn As Long
    rtn = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes Me.hwnd, 0, aero, LWA_ALPHA
'    Call rgnform(Me, rou, rou, 25, 3, 3, 3)
    Me.Text3.Text = 2000
    Me.Check1.Value = 1
    splt = Me.Text3.Text
    spl = Me.Check1.Value
    Me.Option3.Value = True
    Me.Check2.Value = 1
    If Me.Option3.Value = True Then sfl = "o_main" Else sfl = "all"
    sfly = Me.Check2.Value
    Me.Check3.Value = 1
    gif = True
    Call rgn_gif(Me, 25, 3, 3, 3)
End Sub
Private Sub Form_Activate()
    Me.Text1.Text = aero
    Me.Text2.Text = rou
    Me.Text3.Text = splt
    Me.Check1.Value = spl
    If gif = True Then Me.Check3.Value = 1 Else Me.Check3.Value = 0
    Dim rtn As Long
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, aero, LWA_ALPHA
    If gif = True Then
        If gif_settings <> True Then
            Call rgn_gif(Me, 25, 3, 3, 3)
            gif_settings = True
        End If
    End If
    Call rgnform(Me, rou, rou, 25, 3, 3, 3)
End Sub
Private Sub Form_Load()
    Me.Check2.Value = sfly
    If sfl = "o_main" Then
        Me.Option3.Value = True
    Else
        Me.Option4.Value = True
    End If
    If sfly = 1 And sfl = "all" Then
        Me.left = GetSetting(App.ProductName, "location", "left_settings", Screen.Width / 2 - Me.Width / 2)
        Me.top = GetSetting(App.ProductName, "location", "top_settings", Screen.Height / 2 - Me.Height / 2)
    Else
        Me.left = Screen.Width / 2 - Me.Width / 2
        Me.top = Screen.Height / 2 - Me.Height / 2
    End If
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sx = X
    sy = Y
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        settings.left = settings.left + (X - sx)
        settings.top = settings.top + (Y - sy)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    gif_settings = False
    SaveSetting App.ProductName, "location", "left_settings", Me.left
    SaveSetting App.ProductName, "location", "top_settings", Me.top
    DeleteObject outrgn
End Sub
Private Sub Text1_LostFocus()
    If Me.Text1.Text < 10 Then
        Me.Text1.Text = 10
    End If
End Sub
Private Sub Text2_LostFocus()
    If Me.Text2.Text < 0 Then
        Me.Text2.Text = 0
    End If
    If Me.Text2.Text > 90 Then
        Me.Text2.Text = 90
    End If
End Sub
