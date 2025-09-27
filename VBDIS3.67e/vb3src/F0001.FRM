VERSION 2.00
Begin Form FrmPrjDir 
   Caption         =   "Select Project Path"
   ClientHeight    =   3990
   ClientLeft      =   6660
   ClientTop       =   975
   ClientWidth     =   4200
   Height          =   4500
   Left            =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   4200
   Top             =   525
   Width           =   4320
   Begin CommandButton dum201 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   372
      Left            =   3000
      TabIndex        =   4
      Top             =   1380
      Width           =   1092
   End
   Begin CommandButton Cmd_New 
      Caption         =   "New"
      Height          =   372
      Left            =   3000
      TabIndex        =   3
      Top             =   600
      Width           =   1092
   End
   Begin TextBox txt_newDir 
      Height          =   288
      Left            =   2880
      TabIndex        =   2
      Top             =   180
      Width           =   1212
   End
   Begin DirListBox DirListBox 
      Height          =   3264
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2652
   End
   Begin DriveListBox DrivelistBox 
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   2652
   End
End
'Attribute VB_Name = "FrmPrjDir"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = False
' F0001.FRM
Option Explicit

Private Sub Cmd_New_Click ()
Dim newDir As String
Dim CurrentDir As String

36010
  newDir = txt_newDir
  If Len(newDir) = 0 Then Beep: Exit Sub
  
  CurrentDir = DirListBox.List(DirListBox.ListIndex)
  If Right$(CurrentDir, 1) <> "\" Then CurrentDir = CurrentDir & "\"
  CurrentDir = CurrentDir & newDir

On Error Resume Next
  MkDir CurrentDir
  
  DirListBox.Path = CurrentDir
  
  If Err Then Beep: Exit Sub
  
  gWorkDir = CurrentDir & "\"
End Sub

Private Sub DrivelistBox_Change ()
36020
  On Error Resume Next
  DirListBox.Path = DrivelistBox.Drive
End Sub

Private Sub dum201_Click ()
36000
  gWorkDir = DirListBox.Path
  If Right$(gWorkDir, 1) <> "\" Then gWorkDir = gWorkDir & "\"
  Hide
End Sub

