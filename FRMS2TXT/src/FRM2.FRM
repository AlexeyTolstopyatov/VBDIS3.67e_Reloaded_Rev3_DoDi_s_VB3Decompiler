VERSION 5.00
Begin VB.Form frm2 
   Appearance      =   0  '2D
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "About Forms2Text"
   ClientHeight    =   2385
   ClientLeft      =   1695
   ClientTop       =   2685
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
   Icon            =   "FRM2.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   2385
   ScaleWidth      =   4950
   Begin VB.Timer control1 
      Interval        =   500
      Left            =   2520
      Top             =   1440
   End
   Begin VB.CommandButton control2 
      Appearance      =   0  '2D
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Image control3 
      Appearance      =   0  '2D
      Height          =   480
      Index           =   3
      Left            =   1920
      Picture         =   "FRM2.frx":000C
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image control3 
      Appearance      =   0  '2D
      Height          =   480
      Index           =   2
      Left            =   1320
      Picture         =   "FRM2.frx":0316
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image control3 
      Appearance      =   0  '2D
      Height          =   480
      Index           =   1
      Left            =   720
      Picture         =   "FRM2.frx":0620
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image control3 
      Appearance      =   0  '2D
      Height          =   480
      Index           =   0
      Left            =   120
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image control4 
      Appearance      =   0  '2D
      Height          =   480
      Left            =   120
      Top             =   120
      Width           =   480
   End
   Begin VB.Label control5 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "WWW: http://dspace.dial.pipex.com/johnc/"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   720
      TabIndex        =   4
      Top             =   960
      Width           =   3780
   End
   Begin VB.Label control6 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Email: <johnc@dial.pipex.com>"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label control7 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Copyright 1996 John Cartwright"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   2670
   End
   Begin VB.Label control8 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Version 1.0e"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "frm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frm2.FRM
Option Explicit
Dim m0024 As Integer

Private Sub control2_Click()
  Unload Me
End Sub


Private Sub Form_Load()
  control3(0).Picture = frm1.Icon
  Me.Height = 1740
  DoEvents
End Sub

Private Sub control1_Timer()
  control4.Picture = control3(m0024).Picture
  m0024 = m0024 + 1
  If m0024 = 4 Then m0024 = 0
End Sub

