VERSION 5.00
Begin VB.Form frm6 
   Appearance      =   0  '2D
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Notice"
   ClientHeight    =   2055
   ClientLeft      =   1695
   ClientTop       =   2685
   ClientWidth     =   5580
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
   ScaleHeight     =   2055
   ScaleWidth      =   5580
   Begin VB.CommandButton control1 
      Appearance      =   0  '2D
      Caption         =   "Convert all files to &text format"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   5055
   End
   Begin VB.CommandButton control2 
      Appearance      =   0  '2D
      Caption         =   "Convert all files to &binary format"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   5055
   End
   Begin VB.Label control3 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   $"FRM6.frx":0000
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   5175
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frm6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frm6.FRM
Option Explicit
Dim OldMousePointer As Integer

Private Sub control2_Click()
  KeyCombos.TargetType = TYPE_BIN_FRM
  Unload Me
End Sub

Private Sub control1_Click()
  KeyCombos.TargetType = TYPE_TXT_FRM
  Unload Me
End Sub

Private Sub Form_Load()
  OldMousePointer = Screen.MousePointer
  Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = OldMousePointer
End Sub
