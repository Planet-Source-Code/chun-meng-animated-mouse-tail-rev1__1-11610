VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "About Animated Mouse's Tail"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   3795
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   120
      Top             =   2520
      Width           =   3495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "The original JaveScript source code is available free online at  http://javascript.internet.com "
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Chun Meng (chunmeng@mailcityasia.com)"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Port to VB by"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "http://members.xoom.com/ebullets"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Original Code in JavaScript is written by Philip Winston (pwinston@yahoo.com) "
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3255
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3600
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Version : 0.1 (22/09/2000)"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdok_Click()
    Unload Me
   
End Sub
