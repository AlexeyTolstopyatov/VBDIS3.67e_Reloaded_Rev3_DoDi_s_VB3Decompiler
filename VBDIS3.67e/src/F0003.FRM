VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  '2D
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "DoDi's VB Tools"
   ClientHeight    =   2190
   ClientLeft      =   4815
   ClientTop       =   1860
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
   Icon            =   "F0003.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   2190
   ScaleWidth      =   4950
   Begin VB.PictureBox dum203 
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
      Top             =   960
      Width           =   408
   End
   Begin VB.CommandButton dum303 
      Appearance      =   0  '2D
      Caption         =   "wow !"
      Default         =   -1  'True
      Height          =   552
      Left            =   3960
      TabIndex        =   4
      Top             =   180
      Width           =   852
   End
   Begin VB.Label dum403 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      Caption         =   "http://members.aol.com/vbdis/"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   1800
      TabIndex        =   12
      Top             =   1260
      Width           =   3075
   End
   Begin VB.Label dum503 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      Caption         =   "URL"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   960
      TabIndex        =   11
      Top             =   1260
      Width           =   735
   End
   Begin VB.Label dum503 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      Caption         =   "e-mail"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   960
      TabIndex        =   6
      Top             =   1020
      Width           =   735
   End
   Begin VB.Label dum603 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      Caption         =   "<your e-mail address>"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   1860
      Width           =   3075
   End
   Begin VB.Label dum703 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      Caption         =   "<your name>"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1800
      TabIndex        =   9
      Top             =   1635
      Width           =   3030
   End
   Begin VB.Label dum803 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      Caption         =   "Translated by"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1635
      Width           =   1335
   End
   Begin VB.Label dum403 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      Caption         =   "DoDi@compuserve.com"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   1800
      TabIndex        =   7
      Top             =   1020
      Width           =   3075
   End
   Begin VB.Label dum903 
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
   Begin VB.Label duma03 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      Caption         =   "Dr. H.-P. Diettrich"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label dumb03 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      Caption         =   "Copyright 1997"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label dumc03 
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
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' F0003.FRM
Option Explicit

Private Sub dum303_Click()
44000
  Unload Me
End Sub

Private Sub Form_Load()
Dim l001A As String

44010
  dum203.Picture = CurForm_gv0BC0.Icon
  Caption = "DoDi's " & App.Title
  l001A = " cracked Demo + Src :)"
  dum903 = gc031A & gc031E & l001A
  If Left$(dum703, 1) <> "<" Then
  Else
    Height = Height - (ScaleHeight - dum803.Top)
  End If
End Sub
