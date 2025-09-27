VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  '2D
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "DoDi's VB Tools"
   ClientHeight    =   1635
   ClientLeft      =   2880
   ClientTop       =   2235
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "F0001.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   1635
   ScaleWidth      =   4950
   Begin VB.PictureBox Pic1 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   408
      Left            =   240
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   5
      Top             =   780
      Width           =   408
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  '2D
      Caption         =   "wow !"
      Default         =   -1  'True
      Height          =   552
      Left            =   3993
      TabIndex        =   4
      Top             =   605
      Width           =   852
   End
   Begin VB.Label control3 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   253
      Left            =   1936
      TabIndex        =   9
      Top             =   1331
      Width           =   1947
   End
   Begin VB.Label control4 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   253
      Left            =   242
      TabIndex        =   8
      Top             =   1331
      Width           =   1342
   End
   Begin VB.Label control5 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      Caption         =   "D-70180 Stuttgart"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1920
      TabIndex        =   7
      Top             =   960
      Width           =   1932
   End
   Begin VB.Label control6 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      Caption         =   "Weissenburgstr. 2 c"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1920
      TabIndex        =   6
      Top             =   720
      Width           =   1932
   End
   Begin VB.Label control7 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      Caption         =   "2.0"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   1080
      TabIndex        =   3
      Top             =   180
      Width           =   2292
   End
   Begin VB.Label control8 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      Caption         =   "Dr. H.-P. Diettrich"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   1932
   End
   Begin VB.Label control9 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      Caption         =   "Copyright 1995,96"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1572
   End
   Begin VB.Label control10 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Version"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   852
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' F0001.FRM
Option Explicit

Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Form_Load()
Dim l0020 As String
'Pic1.Picture = Screen.ActiveForm.Icon
  Caption = "DoDi's " & App.Title
  control7 = VERSION & VERSION_TYPE & l0020
    control4 = Descramble("Ztoohpzbbzf)")
    control3 = Descramble("_nt;Hv~vy")
End Sub

