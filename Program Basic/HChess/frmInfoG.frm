VERSION 5.00
Begin VB.Form frmInfoG 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Info"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   248
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   198
   Begin VB.ListBox List1 
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   2430
      Left            =   150
      TabIndex        =   0
      Top             =   1140
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Active Game"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   75
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Start a New Game"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   495
      Left            =   150
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
End
Attribute VB_Name = "frmInfoG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
OutSquare 0, 0, ScaleWidth - 1, ScaleHeight, Me
InSquare List1.Left, List1.Top, List1.Width - 1, List1.Height, Me
InSquare Label1.Left, Label1.Top, Label1.Width - 1, Label1.Height, Me
PaintPicture frmChess.Image2.Picture, 15, 5
PaintPicture frmChess.Image2.Picture, ScaleWidth - 35, 5
Left = frmChess.Left + ((frmChess.Width - Width) / 2)
Top = frmChess.Top + ((frmChess.Height - Height) / 2)
End Sub

