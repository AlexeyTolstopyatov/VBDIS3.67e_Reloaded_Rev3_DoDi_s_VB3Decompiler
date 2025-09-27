VERSION 2.00
Begin Form Form1 
   BackColor       =   &H8000000F&
   Caption         =   "DoDi's VBX Fool"
   ClientHeight    =   2910
   ClientLeft      =   4740
   ClientTop       =   6855
   ClientWidth     =   6540
   Height          =   3420
   Left            =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   6540
   Top             =   6405
   Width           =   6660
   Begin CommandButton control2 
      Caption         =   "Unfool"
      Height          =   312
      Left            =   5280
      TabIndex        =   0
      Top             =   60
      Width           =   852
   End
   Begin CommandButton control3 
      Caption         =   "Exit"
      Height          =   312
      Left            =   2400
      TabIndex        =   6
      Top             =   60
      Width           =   732
   End
   Begin CommandButton control4 
      Caption         =   "Fool"
      Height          =   312
      Left            =   3600
      TabIndex        =   5
      Top             =   60
      Width           =   732
   End
   Begin FileListBox control5 
      Height          =   2175
      Left            =   4920
      Pattern         =   "*.vby"
      TabIndex        =   4
      Top             =   420
      Width           =   1575
   End
   Begin FileListBox control6 
      Height          =   2175
      Left            =   3240
      Pattern         =   "*.vbx"
      TabIndex        =   3
      Top             =   420
      Width           =   1575
   End
   Begin DirListBox control7 
      Height          =   2400
      Left            =   120
      TabIndex        =   2
      Top             =   420
      Width           =   3012
   End
   Begin DriveListBox control8 
      Height          =   288
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   2172
   End
End
' F0000.FRM
Option Explicit

Sub control2_Click ()
Dim FileBoxListIndex As Integer
Dim FileNameOld As String
Dim FileNameNew As String
  FileBoxListIndex = control5.ListIndex
  If FileBoxListIndex < 0 Then Beep: Exit Sub
On Error GoTo LE70
  FileNameOld = control5.FileName
  FileNameNew = Left$(FileNameOld, Len(FileNameOld) - 1) & "x"
  Kill FileNameNew
  Name FileNameOld As FileNameNew
  control6.Refresh
  control5.Refresh
Exit Sub

LE70:
  MsgBox "File in use", , "Cannot Unfool"
Exit Sub
End Sub

Sub control3_Click ()
  End
End Sub

Sub control4_Click ()
Dim ListIndex As Integer
Dim FileNameOld As String
Dim VBX_Name As String
Dim FileNameNew As String
Dim hFile As Integer
Dim FileBuff As String
  ListIndex = control6.ListIndex
  If ListIndex < 0 Then Exit Sub

  FileNameOld = control6.FileName
  FileNameNew = Left$(FileNameOld, Len(FileNameOld) - 1) & "y"

  If Len(Dir$(FileNameNew)) Then
    MsgBox FileNameNew, , "Already Fooled"
    Exit Sub
  End If

  MsgBox FileNameOld, , "Fooling"

On Error GoTo LD6C
  Name FileNameOld As FileNameNew

  hFile = FreeFile
  
  Open App.Path & "\" & "vbxfool.dll" For Binary As hFile
  FileBuff = Space$(LOF(hFile))
  Get hFile, , FileBuff
  Close hFile
  
  ListIndex = control5.ListCount
  VBX_Name = "VBX" & Format$(ListIndex) & "FOOL"
  
  ListIndex = InStr(FileBuff, "VBX_FOOL")
  Mid$(FileBuff, ListIndex, 8) = VBX_Name

  ListIndex = InStr(FileBuff, "VBY_FOOL")
  Mid$(FileBuff, ListIndex, 12) = FileNameNew & Chr$(0)
  
  Open FileNameOld For Binary As hFile
  Put hFile, , FileBuff
  Close hFile
  
  control6.Refresh
  control5.Refresh
Exit Sub

LD6C:
  MsgBox "File in use", , "Cannot fool"
  Exit Sub
End Sub

Sub control7_Click ()
Dim FileBoxListIndex2 As Integer
Dim NewFilePath As String
  
  FileBoxListIndex2 = control7.ListIndex
  NewFilePath = control7.List(FileBoxListIndex2)
  
  ChDrive NewFilePath
  ChDir NewFilePath
  
  control6.Path = NewFilePath
  control5.Path = NewFilePath
End Sub

Sub control8_Change ()
  control7.Path = control8
End Sub

