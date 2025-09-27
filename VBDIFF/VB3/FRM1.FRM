VERSION 2.00
Begin Form frm1 
   Caption         =   "DoDi's VB FileCompare"
   ClientHeight    =   3945
   ClientLeft      =   870
   ClientTop       =   1845
   ClientWidth     =   5535
   Height          =   4755
   Icon            =   FRM1.FRX:0000
   Left            =   810
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   5535
   Top             =   1095
   Width           =   5655
   Begin ListBox control2 
      Height          =   1590
      Index           =   1
      Left            =   2760
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   1980
      Width           =   2775
   End
   Begin ListBox control2 
      Height          =   1590
      Index           =   0
      Left            =   0
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   1980
      Width           =   2775
   End
   Begin ListBox control3 
      Height          =   1590
      Left            =   0
      TabIndex        =   1
      Top             =   60
      Width           =   5535
   End
   Begin CommandButton control4 
      Caption         =   "start"
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   732
   End
   Begin Label control5 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "File 2"
      Height          =   252
      Index           =   1
      Left            =   2760
      TabIndex        =   5
      Top             =   1680
      Width           =   2772
   End
   Begin Label control5 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "File 1"
      Height          =   252
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   1680
      Width           =   2772
   End
   Begin Menu control6 
      Caption         =   "&File"
      Begin Menu control7 
         Caption         =   "&Select..."
      End
      Begin Menu control8 
         Caption         =   "&Compare"
      End
      Begin Menu control9 
         Caption         =   "Con&tinue"
      End
      Begin Menu control10 
         Caption         =   "-"
      End
      Begin Menu control11 
         Caption         =   "on &top"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin Menu control12 
         Caption         =   "&Info..."
      End
      Begin Menu control13 
         Caption         =   "E&xit"
      End
   End
   Begin Menu control14 
      Caption         =   "&Edit"
      Begin Menu control15 
         Caption         =   "&Copy"
         Shortcut        =   ^{INSERT}
      End
   End
   Begin Menu control16 
      Caption         =   "&Options"
      Begin Menu control17 
         Caption         =   "ignore &case"
      End
      Begin Menu control17_2 
         Caption         =   "ignore &blanks"
      End
   End
End
' frm1.FRM
Option Explicit

Sub control11_Click ()
  ZOrder 0
End Sub


Sub Form_Resize ()
Dim l0036 As Integer
  If WindowState = 1 Then
    frm2.Hide
    Exit Sub
  End If
  l0036 = ScaleWidth
  control3.Width = l0036
  l0036 = l0036 \ 2
  control5(0).Width = l0036
  control5(1).Left = l0036
  control5(1).Width = l0036
  control2(0).Width = l0036
  control2(0).Height = ScaleHeight - control2(0).Top
  control2(1).Left = l0036
  control2(1).Width = l0036
  control2(1).Height = control2(0).Height
End Sub

Sub control12_Click ()
  frm3.Show 1
End Sub

Sub control13_Click ()
  End
End Sub

Sub control15_Click ()
Dim l0064 As String
Dim l0066
Dim l006A As Control
On Error GoTo L24CA
  Set l006A = ActiveControl
  For l0066 = 0 To l006A.ListCount - 1
    If l006A.Selected(l0066) Then l0064 = l0064 & l006A.List(l0066) & Chr$(13) & Chr$(10)
  Next l0066
  Clipboard.Clear
  Clipboard.SetText l0064
Exit Sub
L24CA:
  MsgBox gc000A, 32
End Sub

Sub control16_Click ()
  control17.Checked = gv0096
  control17_2.Checked = gv0098
End Sub

Sub control17_2_Click ()
  gv0098 = Not gv0098
End Sub

Sub control17_Click ()
  gv0096 = (Not gv0096) And 1
End Sub

Sub control2_DblClick (Index%)
  sub0314 Index
End Sub

Sub control4_Click ()
  control8.Enabled = True
  control9.Enabled = True
  control5(0) = gv0022(0): gv003C = 1
  control5(1) = gv0022(1): gv0040 = 1
  sub021E
End Sub

Sub control7_Click ()
  frm2.Show
End Sub

Sub control8_Click ()
  control4.Value = 1
End Sub

Sub Form_Load ()
  control8.Enabled = False
  control9.Enabled = False
End Sub


Sub sub021E ()
Dim l008A As Long
Dim l008C
Dim l008E As Integer
Dim l0090 As String
  control3.Clear
  control2(0).Clear
  control2(1).Clear
  Refresh
  gv0038 = FreeFile
  Open control5(0) For Input As gv0038
  Seek gv0038, gv003C
  gv003A = FreeFile
  Open control5(1) For Input As gv003A
  Seek gv003A, gv0040
  sub0081
  control8.Enabled = True
  If gv0094 < 0 Then
    control3.AddItem gc0016
    Exit Sub
  End If
  control9.Enabled = True
  l008A = gv0066(gv0094).M0467
  For l008C = -15 To -1
    l008E = (gv0094 + l008C) And &HF
    If gv0066(l008E).M0467 Then control3.AddItem gv0066(l008E).M046E
  Next
  control3.ListIndex = control3.ListCount - 1
  Seek gv0038, gv0066(gv0094).M0467
  For l008C = 1 To 100
    If EOF(gv0038) Then
      control2(0).AddItem "<eof>"
      control9.Enabled = False
      Exit For
    End If
    l008A = Seek(gv0038)
    Line Input #gv0038, l0090
    control2(0).AddItem l0090
    control2(0).ItemData(control2(0).NewIndex) = l008A
  Next l008C
  Seek gv003A, gv007E(gv0094).M0467
  For l008C = 1 To 100
    If EOF(gv003A) Then
      control2(1).AddItem "<eof>"
      control9.Enabled = False
      Exit For
    End If
    l008A = Seek(gv003A)
    Line Input #gv003A, l0090
    control2(1).AddItem l0090
    control2(1).ItemData(control2(1).NewIndex) = l008A
  Next l008C
  Close
End Sub

Sub sub0314 (p00AA As Integer)
Dim l00AC As Integer
Dim l00AE As Integer
Dim l00B0 As String
Dim l00B2 As String
Dim l00B4 As Integer
Dim l00B6 As Integer
  l00AC = control2(0).ListIndex
  l00AE = control2(1).ListIndex
  If l00AC < 0 Or l00AE < 0 Then GoTo L2A78
  If fn008E((control2(0).List(l00AC)), (control2(1).List(l00AE))) Then GoTo L2A78
  gv003C = control2(0).ItemData(l00AC)
  gv0040 = control2(1).ItemData(l00AE)
  sub021E
Exit Sub
L2A78:
  If p00AA Then
    l00B0 = control2(1).List(l00AE)
  Else
    l00B0 = control2(0).List(l00AC)
  End If
  l00B6 = InStr(l00B0, ")")
  If l00B6 = 0 Or (Left$(l00B0, 4) <> "Sub " And Left$(l00B0, 9) <> "Function ") Then
    MsgBox gc000E
    Exit Sub
  End If
  l00B0 = Left$(l00B0, l00B6)
  Screen.MousePointer = 11
  l00B4 = FreeFile
  Open control5(1 - p00AA) For Input As l00B4
  While Not EOF(l00B4)
    gv003C = Seek(l00B4)
    Line Input #l00B4, l00B2
    If Left$(l00B2, l00B6) = l00B0 Then
      Close l00B4
      Screen.MousePointer = 0
      If p00AA = 0 Then
        gv0040 = gv003C
        gv003C = control2(0).ItemData(l00AC)
      Else
        gv0040 = control2(1).ItemData(l00AE)
      End If
      sub021E
      Exit Sub
    End If
  Wend
  Close l00B4
  Screen.MousePointer = 0
  MsgBox l00B0, 48, gc001A
End Sub

