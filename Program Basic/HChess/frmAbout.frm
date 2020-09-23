VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Harvey Chess"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   298
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "About Project1"
   Begin VB.CommandButton command1 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1560
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   2400
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit to : Carl Harvey elterrorista@videotron.ca"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Harvey Chess"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   1
      Left            =   750
      TabIndex        =   3
      Tag             =   "Application Title"
      Top             =   150
      Width           =   2325
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Harvey Chess"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   360
      Index           =   0
      Left            =   720
      TabIndex        =   2
      Tag             =   "Application Title"
      Top             =   120
      Width           =   2325
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   16
      X2              =   241
      Y1              =   128
      Y2              =   128
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   16
      X2              =   240
      Y1              =   136
      Y2              =   136
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   1
      Tag             =   "Version"
      Top             =   1680
      Width           =   1125
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Reg Key Security Options...

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
OutSquare 0, 0, ScaleWidth - 1, ScaleHeight, Me
InSquare Command1.Left, Command1.Top, Command1.Width - 1, Command1.Height, Me
InSquare 24, 16, frmChess.Image2.Width - 1, frmChess.Image2.Height, Me
Dim c, l
c = 10: l = 10
For i = 1 To 3
 For i2 = 1 To 12
  PaintPicture frmChess.Image2.Picture, c, l
   c = c + 20
 Next i2
c = 10
l = l + 22
Next i
 
Left = frmChess.Left + ((frmChess.Width - Width) / 2)
End Sub

