VERSION 5.00
Begin VB.Form frmBounce 
   BorderStyle     =   0  'None
   Caption         =   "Animated Mouse's Tail"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "About"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Left            =   2520
      Top             =   1440
   End
   Begin VB.Label Label1 
      Caption         =   "Double Click to Exit"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
   Begin VB.Image ImgBall 
      Height          =   165
      Index           =   0
      Left            =   1560
      Picture         =   "frmBounce.frx":0000
      Top             =   960
      Width           =   165
   End
End
Attribute VB_Name = "frmBounce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    frmAbout.Show vbModal
End Sub

Private Sub Form_DblClick()
    Dim i As Integer
    For i = ImgBall.LBound + 1 To ImgBall.UBound
        Unload ImgBall(i)
    Next i
    Unload Me
    End
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = ImgBall.UBound + 1 To 7
        Load ImgBall(i)
        ImgBall(i).Visible = True
        ImgBall(i).Top = ImgBall(i - 1).Top + 11
    Next i
    ImgBall(0).Visible = False
    Call InitVal
    Call InitBall
    Timer1.Interval = 20
    Timer1.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveHandler CLng(X), CLng(Y)
    Animate
End Sub

Private Sub Timer1_Timer()
    Animate
End Sub
