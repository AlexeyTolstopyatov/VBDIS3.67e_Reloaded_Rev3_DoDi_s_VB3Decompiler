VERSION 5.00
Begin VB.Form frm5 
   Appearance      =   0  '2D
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Notice"
   ClientHeight    =   1425
   ClientLeft      =   1860
   ClientTop       =   2175
   ClientWidth     =   5265
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
   ScaleHeight     =   1425
   ScaleWidth      =   5265
   Begin VB.CommandButton control1 
      Appearance      =   0  '2D
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
   Begin VB.Label control2 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   75
   End
End
Attribute VB_Name = "frm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frm5.FRM
Option Explicit
Dim m001A As Integer

Private Sub control1_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  m001A = Screen.MousePointer
  Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = m001A
End Sub
