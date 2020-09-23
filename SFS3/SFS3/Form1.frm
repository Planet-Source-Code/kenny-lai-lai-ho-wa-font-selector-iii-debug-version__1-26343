VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   3900
   StartUpPosition =   3  '¨t²Î¹w³]­È
   Begin VB.CommandButton Command2 
      Caption         =   "Random Set Text"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin Project1.Font fs1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      _extentx        =   5741
      _extenty        =   556
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command2_Click()
fs1.Text = Screen.Fonts(Int(Rnd * Screen.FontCount))
End Sub

Private Sub fs1_Click()
MsgBox fs1.Text
End Sub
