VERSION 5.00
Begin VB.Form 黑屏 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "嗨嗨嗨"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   12  'No Drop
   Moveable        =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   1  '所有者中心
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "黑屏"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Sub Form_Load()
    Dim rtn As Long
    rtn = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, 10, LWA_ALPHA
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
End Sub
