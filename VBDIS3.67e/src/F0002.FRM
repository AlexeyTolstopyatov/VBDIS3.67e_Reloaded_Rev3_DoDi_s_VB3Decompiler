VERSION 5.00
Begin VB.Form FrmOpenFile 
   Appearance      =   0  '2D
   BackColor       =   &H80000005&
   Caption         =   "Open File"
   ClientHeight    =   3195
   ClientLeft      =   2745
   ClientTop       =   1365
   ClientWidth     =   4410
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
   LinkTopic       =   "Form4"
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   3195
   ScaleWidth      =   4410
   Begin VB.CommandButton dum202 
      Appearance      =   0  '2D
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   372
      Left            =   3480
      TabIndex        =   5
      Top             =   2700
      Width           =   852
   End
   Begin VB.CommandButton dum302 
      Appearance      =   0  '2D
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   2520
      TabIndex        =   4
      Top             =   2700
      Width           =   852
   End
   Begin VB.ComboBox dum402 
      Appearance      =   0  '2D
      Height          =   288
      Left            =   2520
      Style           =   2  'Dropdown-Liste
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.FileListBox dum502 
      Appearance      =   0  '2D
      Height          =   2370
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1812
   End
   Begin VB.DirListBox dum602 
      Appearance      =   0  '2D
      Height          =   2616
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2292
   End
   Begin VB.DriveListBox DrivelistBox 
      Appearance      =   0  '2D
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2292
   End
End
Attribute VB_Name = "FrmOpenFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' F0002.FRM
Option Explicit

Private Sub dum302_Click()
43000
  FrmOpenFile.Tag = ""
  Hide
End Sub

Private Sub dum202_Click()
Dim l001E As String
Dim l0020 As String

43010
  l001E = dum502.Filename
  If l001E = "" Then Beep: dum502.SetFocus: Exit Sub
  l0020 = dum502.Path
  If Right$(l0020, 1) <> "\" Then l0020 = l0020 & "\"
  FrmOpenFile.Tag = l0020 & l001E
  Hide
End Sub

Private Sub dum402_Click()
Dim l002A As Integer

43020
  l002A = dum402.ListIndex
  If l002A < 0 Then Exit Sub
  dum502.Pattern = dum402.List(l002A)
End Sub

Private Sub dum602_Change()
43030
  dum502.Path = dum602.List(dum602.ListIndex)
End Sub

Private Sub dum602_Click()
43040
  dum502.Path = dum602.List(dum602.ListIndex)
End Sub

Private Sub DrivelistBox_Change()
43050
  dum602.Path = DrivelistBox.Drive
End Sub

Private Sub dum502_DblClick()
43060
  dum202.Value = 1
End Sub

Private Sub Form_Activate()
Dim l0044 As String
Dim l0046 As Integer

43070
  l0044 = Tag
  If Len(l0044) = 0 Then
    If dum402.ListCount Then
      l0046 = dum402.ListIndex
      If l0046 < 0 Then l0046 = 0: dum402.ListIndex = 0
      l0044 = dum402.List(l0046)
    End If
    If Len(l0044) = 0 Then l0044 = "*.*"
  End If
  Caption = gc03CE & l0044
  dum502.Path = dum602.List(dum602.ListIndex)
  dum502.Pattern = l0044
End Sub

Private Sub Form_Resize()
43080
  If dum402.ListCount Then
    dum502.Top = dum602.Top
    dum502.Height = dum202.Top - dum502.Top
    dum402.Visible = True
  End If
End Sub
