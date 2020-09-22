VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3D - Pong"
   ClientHeight    =   7500
   ClientLeft      =   6315
   ClientTop       =   4395
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   603
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox P2Score 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3360
      TabIndex        =   4
      Text            =   "0"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox P1Score 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "0"
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox Display 
      BackColor       =   &H00000000&
      Height          =   6975
      Left            =   0
      ScaleHeight     =   6915
      ScaleWidth      =   8955
      TabIndex        =   0
      Top             =   480
      Width           =   9015
      Begin VB.Label Loading 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3960
         TabIndex        =   6
         Top             =   3000
         Width           =   1095
      End
   End
   Begin VB.Label Tips 
      BackStyle       =   0  'Transparent
      Caption         =   "[Hit Enter to Start the ball]"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   5
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Player 2:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Player 1:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Menu mnuGOpts 
      Caption         =   "Game Options"
      Begin VB.Menu mnuRot 
         Caption         =   "Rotate Camera to Red Player"
      End
      Begin VB.Menu mnuRestart 
         Caption         =   "Restart Game"
      End
      Begin VB.Menu mnuSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit Game"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Show
    DoEvents
    
    Init Display.HWND
    
    Loading.Visible = False
    Do While Finish = False
        CheckKey
        RenderAll
        DoEvents
    Loop
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    KillApp
End Sub

Private Sub mnuExit_Click()
    KillApp
End Sub

Private Sub mnuHelp_Click()
    frmHelp.Show
End Sub

Private Sub mnuRestart_Click()
    ResetBall
    P1Points = 0
    P2Points = 0
    P1Score.Text = "0"
    P2Score.Text = "0"
End Sub

Private Sub mnuRot_Click()
    If Rot = 1 Then
        Rot = -1
        mnuRot.Caption = "Rotate Camera to Blue Player"
    Else
        Rot = 1
        mnuRot.Caption = "Rotate Camera to Red Player"
    End If
    EyePosition.Z = -EyePosition.Z
    SetupMatrices
End Sub

