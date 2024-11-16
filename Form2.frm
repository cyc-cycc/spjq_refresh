VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Òýµ¼"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4455
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   1920
   ScaleWidth      =   4455
   StartUpPosition =   1  'ËùÓÐÕßÖÐÐÄ
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   960
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "È¡Ïû"
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
      Left            =   3240
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ALT+S"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CTRL+ENTER"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1380
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ENTER"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Ë¢ÆÁÆÚ¼ä²»Òªµã»÷ÆäËüÈÎºÎ´°¿Ú£¡"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "Ë¢ÆÁÑÓ³Ù:"
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
      Left            =   3000
      TabIndex        =   9
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Ë¢ÆÁ¼ä¸ô:"
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
      Left            =   1560
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Ë¢ÆÁ´ÎÊý:"
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
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "ÇëÑ¡ÔñÏûÏ¢·¢ËÍ·½Ê½:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Me.Hide
    Call sp(Form1.num.Text, Form1.wai.Text, Form1.Text2.Text, 1)
    Form1.Show
    Unload Me
End Sub
Private Sub Command2_Click()
    Me.Hide
    Call sp(Form1.num.Text, Form1.wai.Text, Form1.Text2.Text, 2)
    Form1.Show
    Unload Me
End Sub
Private Sub Command3_Click()
    Me.Hide
    Call sp(Form1.num.Text, Form1.wai.Text, Form1.Text2.Text, 3)
    Form1.Show
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
    Me.Text1.Text = Form1.num.Text
    Me.Text2.Text = Form1.wai.Text
    Me.Text3.Text = Form1.Text2.Text
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
Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
    DeleteObject outrgn
End Sub
