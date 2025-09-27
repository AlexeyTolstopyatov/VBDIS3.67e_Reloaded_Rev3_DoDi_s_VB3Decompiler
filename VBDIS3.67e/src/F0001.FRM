VERSION 5.00
Begin VB.Form FrmPrjDir 
   Appearance      =   0  '2D
   BackColor       =   &H80000005&
   Caption         =   "Select Project Path"
   ClientHeight    =   3990
   ClientLeft      =   6660
   ClientTop       =   975
   ClientWidth     =   4200
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
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   3990
   ScaleWidth      =   4200
   Begin VB.CommandButton dum201 
      Appearance      =   0  '2D
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   372
      Left            =   3000
      TabIndex        =   4
      Top             =   1380
      Width           =   1092
   End
   Begin VB.CommandButton Cmd_New 
      Appearance      =   0  '2D
      Caption         =   "New"
      Height          =   372
      Left            =   3000
      TabIndex        =   3
      Top             =   600
      Width           =   1092
   End
   Begin VB.TextBox txt_newDir 
      Appearance      =   0  '2D
      Height          =   288
      Left            =   2880
      TabIndex        =   2
      Top             =   180
      Width           =   1212
   End
   Begin VB.DirListBox DirListBox 
      Appearance      =   0  '2D
      Height          =   3264
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2652
   End
   Begin VB.DriveListBox DrivelistBox 
      Appearance      =   0  '2D
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   2652
   End
End
Attribute VB_Name = "FrmPrjDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' F0001.FRM
Option Explicit


Private Sub dum201_Click()
36000
  gWorkDir = DirListBox.Path
  If Right$(gWorkDir, 1) <> "\" Then gWorkDir = gWorkDir & "\"
  Hide
End Sub

Private Sub Cmd_New_Click()
Dim newDir As String
Dim CurDir As String

36010
  newDir = txt_newDir
  If Len(newDir) = 0 Then Beep: Exit Sub
  
  CurDir = DirListBox.List(DirListBox.ListIndex)
  If Right$(CurDir, 1) <> "\" Then CurDir = CurDir & "\"
  CurDir = CurDir & newDir

On Error Resume Next
  MkDir CurDir
  
  DirListBox.Path = CurDir
  
  If Err Then Beep: Exit Sub
  
  gWorkDir = CurDir & "\"
End Sub

Private Sub DrivelistBox_Change()
36020
  On Error Resume Next
  DirListBox.Path = DrivelistBox.Drive
End Sub

