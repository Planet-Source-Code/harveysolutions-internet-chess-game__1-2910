VERSION 5.00
Begin VB.Form frmnewgame 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Game"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   159
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000002&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000002&
      Caption         =   "Ok"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   1800
      Width           =   855
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H80000002&
      Caption         =   "Visitor"
      Enabled         =   0   'False
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H80000002&
      Caption         =   "Host"
      Enabled         =   0   'False
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000002&
      Caption         =   "Oline Game"
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000002&
      Caption         =   "Offline Game"
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Value           =   -1  'True
      Width           =   1815
   End
End
Attribute VB_Name = "frmnewgame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'On Error Resume Next
If Not Connected And Not PlayOffline Then GoTo j2
 If Connected Then
  If PlayOffline Then
    frmChess.SockClient.SendData "#020code#quitegame"
    frmChess.HChessBoard1.EraseBoard
    GoTo j2
  End If
  Dim d
  If JoueurHote Then
   d = "Hote"
  Else
   d = "No"
  End If
  'MsgBox "Envoi les donnees"
  frmChess.SockClient.SendData "#020code#wantnewga" & d
  frmChess.HChessBoard1.EraseBoard
  Unload Me
  frmWaitting.Show 1
  Exit Sub
 End If
j2:
Unload Me
NameNIP.Show 1
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Load()
OutSquare 0, 0, ScaleWidth - 1, ScaleHeight, Me
InSquare Option1.Left, Option1.Top, Option1.Width - 1, Option1.Height, Me
InSquare Option2.Left, Option2.Top, Option2.Width - 1, Option2.Height, Me
InSquare Option3.Left, Option3.Top, Option3.Width - 1, Option3.Height, Me
InSquare Option4.Left, Option4.Top, Option4.Width - 1, Option4.Height, Me
InSquare Command1.Left, Command1.Top, Command1.Width - 1, Command1.Height, Me
InSquare Command2.Left, Command2.Top, Command2.Width - 1, Command2.Height, Me
JoueurHote = True
PlayOffline = True
Left = frmChess.Left + ((frmChess.Width - Width) / 2)
Top = frmChess.Top + ((frmChess.Height - Height) / 2)
End Sub
Private Sub Option1_Click()
'On Error Resume Next
Option3.Enabled = False: Option4.Enabled = False
JoueurHote = True: PlayOffline = True
End Sub
Private Sub Option2_Click()
'On Error Resume Next
Option3.Enabled = True: Option4.Enabled = True
PlayOffline = False
End Sub
Private Sub Option3_Click()
JoueurHote = True
End Sub
Private Sub Option4_Click()
JoueurHote = False
End Sub

