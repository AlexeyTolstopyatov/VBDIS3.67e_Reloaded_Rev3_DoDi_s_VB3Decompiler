VERSION 5.00
Begin VB.Form frm1 
   Appearance      =   0  '2D
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Forms2Text"
   ClientHeight    =   2400
   ClientLeft      =   1350
   ClientTop       =   2580
   ClientWidth     =   6120
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
   Icon            =   "FRM1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   2400
   ScaleWidth      =   6120
   Begin VB.CommandButton control1 
      Appearance      =   0  '2D
      Caption         =   "&Convert"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.FileListBox control2 
      Appearance      =   0  '2D
      Height          =   1980
      Left            =   2880
      Pattern         =   "*.MAK"
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton control3 
      Appearance      =   0  '2D
      Caption         =   "&About"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton control4 
      Appearance      =   0  '2D
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.DirListBox control5 
      Appearance      =   0  '2D
      Height          =   1605
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   2535
   End
   Begin VB.DriveListBox control6 
      Appearance      =   0  '2D
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   2535
   End
   Begin VB.Menu control7 
      Caption         =   "&File"
      Begin VB.Menu control8 
         Caption         =   "&Info ..."
      End
      Begin VB.Menu control9 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frm1.FRM
Option Explicit

Private Sub control3_Click()
  frm2.Show
End Sub

Private Sub control1_Click()
  sFile.Mak_Path = control5.Path + "\"
  sFile.Mak_File = control2.Filename
  If FileExists(sFile.Mak_Path & sFile.Mak_File) Then
    Screen.MousePointer = 11
    If ReadMakFile() = True Then DoConvention
  Else
    ShowStatus "Error: Invalid .MAK file." & Chr$(10) & Chr$(13) & "Cause: Non-VB file selected."
  End If
End Sub

Private Sub control4_Click()
  AppExit (EXIT_CODE_EXIT_CANCEL)
End Sub

Private Sub control5_Change()
  control2.Path = control5.Path
End Sub

Private Sub control6_Change()
  Screen.MousePointer = 11
  control5.Path = control6.Drive
  Screen.MousePointer = 0
End Sub

Private Sub Form_Activate()
  Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
  control6.Drive = CurDir$
  control5.Path = CurDir$
  control2.Path = CurDir$
End Sub

Private Sub control9_Click()
  control4_Click
End Sub

Private Sub control8_Click()
  frm2.Show
End Sub

