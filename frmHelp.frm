VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Help"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox MainText 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Left = FrmMain.Left + FrmMain.Width
    Me.Top = FrmMain.Top
    Me.Height = FrmMain.Height
    
    MainText.Text = "Player 1 will have to use the 2 arrow keys to move " + _
                    "it's player and Player 2 will have to use the A and W " + _
                    "Keys for his player. The game goes on for ever. To change " + _
                    "the perspective to Reds perspective, go to the menu and click " + _
                    "on: Rotate Camera to Red player. Click again on this button to " + _
                    "Go back to blues perspective. This program was written by " + _
                    "SturmNacht (www.SturmNacht.de.vu)"
    
End Sub

