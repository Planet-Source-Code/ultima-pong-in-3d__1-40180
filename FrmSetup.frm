VERSION 5.00
Begin VB.Form FrmSetup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Settings"
   ClientHeight    =   3210
   ClientLeft      =   10395
   ClientTop       =   6585
   ClientWidth     =   2625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   2625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Ball Speed:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   2415
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   120
         Max             =   5
         Min             =   1
         TabIndex        =   5
         Top             =   360
         Value           =   1
         Width           =   2175
      End
      Begin VB.Label SpeedScroll 
         Caption         =   "1"
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
         Left            =   1200
         TabIndex        =   6
         Top             =   600
         Width           =   255
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Type of game:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.OptionButton P2 
         Caption         =   "&Human vs. Human"
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
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton AI 
         Caption         =   "Human vs. &Computer"
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
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   2175
      End
   End
End
Attribute VB_Name = "FrmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AI_Click()
    P2Con = False
End Sub

Private Sub Command1_Click()
    BallSpeed = HScroll1.Value / 1.5
    FrmSetup.Hide
    FrmMain.Show
End Sub

Private Sub Form_Load()
    P2Con = True
End Sub

Private Sub HScroll1_Change()
    SpeedScroll.Caption = HScroll1.Value
End Sub

Private Sub P2_Click()
    P2Con = True
End Sub
