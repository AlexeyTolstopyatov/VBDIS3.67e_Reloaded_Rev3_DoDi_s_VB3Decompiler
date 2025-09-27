VERSION 2.00
Begin Form frm2 
   Caption         =   "File Selection"
   ClientHeight    =   6195
   ClientLeft      =   900
   ClientTop       =   1560
   ClientWidth     =   4920
   Height          =   6705
   Icon            =   FRM2.FRX:0000
   Left            =   840
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   4920
   Top             =   1110
   Width           =   5040
   Begin ComboBox control1 
      Height          =   300
      Left            =   1080
      TabIndex        =   9
      Top             =   60
      Width           =   3735
   End
   Begin CommandButton control2 
      Caption         =   "Mask"
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   60
      Width           =   855
   End
   Begin FileListBox control3 
      Height          =   2904
      Index           =   1
      Left            =   2520
      Pattern         =   "*.bas;*.frm;*.txt"
      TabIndex        =   6
      Top             =   720
      Width           =   2292
   End
   Begin DirListBox control4 
      Height          =   2055
      Index           =   1
      Left            =   2520
      TabIndex        =   5
      Top             =   3720
      Width           =   2295
   End
   Begin DriveListBox control5 
      Height          =   288
      Index           =   1
      Left            =   2520
      TabIndex        =   4
      Top             =   5820
      Width           =   2292
   End
   Begin FileListBox control3 
      Height          =   2904
      Index           =   0
      Left            =   120
      Pattern         =   "*.bas;*.frm;*.txt"
      TabIndex        =   2
      Top             =   720
      Width           =   2292
   End
   Begin DirListBox control4 
      Height          =   2055
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   2295
   End
   Begin DriveListBox control5 
      Height          =   288
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   5820
      Width           =   2292
   End
   Begin Label control6 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   7
      Top             =   420
      Width           =   2295
   End
   Begin Label control6 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   420
      Width           =   2295
   End
End
' frm2.FRM
Option Explicit
Const mc001A = "*.*"
Const mc001E = "*.bas;*.frm;*.txt;*.mak"
Const mc0022 = "*.c*;*.h*;*.rc;*.def"

Sub control2_Click ()
  control3(0).Pattern = control1.Text
  control3(1).Pattern = control1.Text
End Sub

Sub control1_Click ()
Dim l0032 As Integer
  l0032 = control1.ListIndex
  If  l0032 < 0 Then Beep: Exit Sub
  control2_Click
End Sub

Sub control4_Change (Index%)
  control3(Index).Path = control4(Index).Path
  control6(Index) = control3(Index).Path
End Sub

Sub control5_Change (Index%)
  control4(Index).Path = control5(Index).Drive
  If  Index = 0 Then
    control5(1).Drive = control5(0).Drive
  End If
End Sub

Sub control3_Click (Index%)
Dim l004C As Integer
Dim l004E As String
  If  Index Then Exit Sub
  l004C = control3(0).ListIndex
  If  l004C < 0 Then Exit Sub
  l004E = control3(0).List(l004C)
  For l004C = control3(1).ListCount - 1 To 0 Step -1
    If  control3(1).List(l004C) = l004E Then Exit For
  Next
  control3(1).ListIndex = l004C
End Sub

Sub control3_DblClick (Index%)
Dim l0052
  For l0052 = 0 To 1
    If  control3(l0052).ListIndex < 0 Then
      Beep
      control3(l0052).SetFocus
      Exit Sub
    End If
    gv0022(l0052) = control3(l0052).Path & "\" & control3(l0052).FileName
  Next
  If  gv0022(0) = gv0022(1) Then
    MsgBox gc0012, 16
    Exit Sub
  End If
  frm1.Show
  frm1.SetFocus
  frm1.control4.Value = 1
End Sub

Sub Form_Load ()
  control1.AddItem mc001A
  control1.AddItem mc001E
  control1.AddItem mc0022
  control1.ListIndex = 1
End Sub
