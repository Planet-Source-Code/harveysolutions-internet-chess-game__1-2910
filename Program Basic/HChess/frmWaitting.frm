VERSION 5.00
Begin VB.Form frmWaitting 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Waitting"
   ClientHeight    =   810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   54
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   317
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Waitting for Opponent answer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   4455
   End
End
Attribute VB_Name = "frmWaitting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
OutSquare 0, 0, ScaleWidth - 1, ScaleHeight, Me
InSquare Label1.Left, Label1.Top, Label1.Width - 1, Label1.Height, Me
Left = frmChess.Left + ((frmChess.Width - Width) / 2)
Top = frmChess.Top + ((frmChess.Height - Height) / 2)
End Sub

