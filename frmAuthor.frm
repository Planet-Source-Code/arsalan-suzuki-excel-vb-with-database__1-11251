VERSION 5.00
Begin VB.Form frmAuthor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About The Author"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4800
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Contacts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1200
      Left            =   105
      TabIndex        =   3
      Top             =   2160
      Width           =   4590
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   750
         Left            =   165
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "frmAuthor.frx":0000
         Top             =   285
         Width           =   4245
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "About the author"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2010
      Left            =   90
      TabIndex        =   1
      Top             =   45
      Width           =   4590
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   1575
         Left            =   165
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "frmAuthor.frx":00A0
         Top             =   285
         Width           =   4245
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Bye "
      Default         =   -1  'True
      Height          =   330
      Left            =   1470
      TabIndex        =   0
      Top             =   3435
      Width           =   2055
   End
End
Attribute VB_Name = "frmAuthor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Hcolor = vbBlue
Const Color = vbBlack

Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text2.ForeColor = Color
Text1.ForeColor = Color
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text2.ForeColor = Hcolor
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Text1.ForeColor = Hcolor
End Sub

