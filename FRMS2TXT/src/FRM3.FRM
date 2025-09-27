VERSION 5.00
Begin VB.Form frm3 
   Appearance      =   0  '2D
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Forms2Text"
   ClientHeight    =   6000
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   7305
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   6000
   ScaleWidth      =   7305
   Begin VB.CommandButton control1 
      Appearance      =   0  '2D
      Caption         =   "&Click here when all information is correct"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   5400
      Width           =   5295
   End
   Begin VB.Frame control2 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      Caption         =   "Settings"
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   6855
      Begin VB.TextBox control3 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   6375
      End
      Begin VB.TextBox control3 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   2760
         Width           =   6375
      End
      Begin VB.TextBox control3 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   2040
         Width           =   6375
      End
      Begin VB.TextBox control3 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   6375
      End
      Begin VB.Label control4 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Enter shortcut for 'File - Save File As ... - Save as Text': (default = fas)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   6045
      End
      Begin VB.Label control4 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Enter shortcut for 'Yes': (default=y)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   2520
         Width           =   2985
      End
      Begin VB.Label control4 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Enter shortcut for 'File - Exit': (default = fx)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   3645
      End
      Begin VB.Label control4 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Enter shortcut for 'Window - Project: (default = wr)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   4305
      End
   End
   Begin VB.Label control4 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   $"FRM3.frx":0000
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   6855
      WordWrap        =   -1  'True
   End
   Begin VB.Label control4 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   $"FRM3.frx":00F1
      ForeColor       =   &H80000008&
      Height          =   855
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   6855
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frm3.FRM
Option Explicit

Private Sub control1_Click()
Dim l001A As Integer
  KeyCombos.PrjManager = control3(0)
  KeyCombos.FileSaveAsText = control3(1)
  KeyCombos.ExitVB3 = control3(2)
  KeyCombos.Yes = control3(3)
  l001A = WritePrivateProfileString("Forms2Text", "WindowProject", KeyCombos.PrjManager, App.Path & "\FRMS2TXT.INI")
  l001A = WritePrivateProfileString("Forms2Text", "FileSaveAsText", KeyCombos.FileSaveAsText, App.Path & "\FRMS2TXT.INI")
  l001A = WritePrivateProfileString("Forms2Text", "FileExit", KeyCombos.ExitVB3, App.Path & "\FRMS2TXT.INI")
  l001A = WritePrivateProfileString("Forms2Text", "Yes", KeyCombos.Yes, App.Path & "\FRMS2TXT.INI")
  Unload Me
End Sub

Private Sub Form_Load()
  control3(0) = KeyCombos.PrjManager
  control3(1) = KeyCombos.FileSaveAsText
  control3(2) = KeyCombos.ExitVB3
  control3(3) = KeyCombos.Yes
End Sub
