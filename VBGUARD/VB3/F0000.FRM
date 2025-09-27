VERSION 2.00
Begin Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "DoDi's VB Guard"
   ClientHeight    =   3705
   ClientLeft      =   3600
   ClientTop       =   5160
   ClientWidth     =   5190
   Height          =   4215
   Icon            =   F0000.FRX:0000
   Left            =   3540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5190
   Top             =   4710
   Width           =   5310
   Begin SSFrame SSFrame1 
      Caption         =   "Step 3: Create protected copy"
      Font3D          =   0  'None
      ForeColor       =   &H00000000&
      Height          =   836
      Left            =   121
      TabIndex        =   0
      Top             =   2200
      Width           =   4972
      Begin SSCommand control2 
         Caption         =   "Destination"
         Font3D          =   0  'None
         Height          =   286
         Left            =   121
         TabIndex        =   1
         Top             =   330
         Width           =   1221
      End
      Begin Label txt_FileOut 
         BorderStyle     =   1  'Fixed Single
         Height          =   253
         Left            =   1452
         TabIndex        =   2
         Top             =   330
         Width           =   3377
      End
   End
   Begin SSFrame control4 
      Caption         =   "Step 2: Add Version Info ressource (optional)"
      Font3D          =   0  'None
      ForeColor       =   &H00000000&
      Height          =   836
      Left            =   121
      TabIndex        =   7
      Top             =   1155
      Width           =   4972
      Begin SSCommand control5 
         Caption         =   "VersionInfo"
         Font3D          =   0  'None
         Height          =   286
         Left            =   121
         TabIndex        =   8
         Top             =   330
         Width           =   1221
      End
      Begin Label txt_FileRes 
         BorderStyle     =   1  'Fixed Single
         Height          =   253
         Left            =   1452
         TabIndex        =   9
         Top             =   330
         Width           =   3377
      End
   End
   Begin SSFrame control7 
      Caption         =   "Step 1: Select the program to protect"
      Font3D          =   0  'None
      Height          =   836
      Left            =   121
      TabIndex        =   4
      Top             =   110
      Width           =   4972
      Begin SSCommand control8 
         Caption         =   "Program"
         Font3D          =   0  'None
         Height          =   286
         Left            =   121
         TabIndex        =   5
         Top             =   330
         Width           =   1221
      End
      Begin Label txt_FileIn 
         BorderStyle     =   1  'Fixed Single
         Height          =   286
         Left            =   1452
         TabIndex        =   6
         Top             =   330
         Width           =   3377
      End
   End
   Begin CommonDialog OpenFileDlg 
      Left            =   3751
      Top             =   3190
   End
   Begin CommandButton control11 
      Caption         =   "Done"
      Height          =   432
      Left            =   1936
      TabIndex        =   3
      Top             =   3190
      Width           =   1332
   End
End
' F0000.FRM
Option Explicit
Dim FileInput As String
Dim FileOutput As String
Dim FilenameRes As String
Dim ResDataBuff As String
Const COMMDLG_ACTION_OPEN = 1 ' &H1%
Const COMMDLG_ACTION_SAVEAS = 2 ' &H2%
Const OFN_FILEMUSTEXIST = 4096 ' &H1000%
Const VB_GUARD = 16 ' &H10%
Const mc004E = 64 ' &H40%

Sub control11_Click ()
  Unload Me
End Sub

Sub control2_Click ()
'On Error GoTo ERR_CANCEL
  OpenFileDlg.Flags = OFN_FILEMUSTEXIST
  OpenFileDlg.DialogTitle = SAVE_AS
  OpenFileDlg.Filter = PROGRAM_EXECUTABLES
  OpenFileDlg.Filename = FileInput
  OpenFileDlg.CancelError = True
  
  OpenFileDlg.Action = COMMDLG_ACTION_SAVEAS
  FileOutput = LCase$(OpenFileDlg.Filename)
'FileInput = "M:\t\VBDIS3~1.59E\VBGUARD\VBXFOOL.EXE"
'FileOutput = "m:\t\vbdis3~1.59e\vbguard\vbxfoo1.exe"

ERR_Continue:
  If Len(FileOutput) Then
    
    If FileInput = FileOutput Then
      MsgBox ERR_CannotCopytoSameFile, VB_GUARD
      Exit Sub
    End If
    
    txt_FileOut = FileOutput
    
   'Load Filedata
    ReadFileData FileInput
    
    If NE_FileHeader.Signature = 0 Then
      MsgBox FileInput, VB_GUARD, ERR_FileNotFound
    
    Else
      
     'Store
      WriteFileData FileOutput, ResDataBuff
      
      MsgBox FileOutput, mc004E, PROGRAM_PROTECED
      
      txt_Filein = ""
      txt_FileOut = ""
      ResDataBuff = ""
      control2.Enabled = False
    End If
    
    Close hFile

  End If
Exit Sub

ERR_CANCEL:
  FileOutput = ""
  Resume ERR_Continue
End Sub

Sub control5_Click ()
On Error GoTo L42C0
  OpenFileDlg.Flags = OFN_FILEMUSTEXIST
  OpenFileDlg.DialogTitle = VERSIONINFO_RES
  OpenFileDlg.Filter = PROGRAM_RESOURCES
  If Right$(FileInput, 4) = ".exe" Then
    FilenameRes = FileInput
    Mid$(FilenameRes, Len(FilenameRes) - 2) = "res"
  Else
    FilenameRes = ""
  End If
  OpenFileDlg.Filename = FilenameRes
  OpenFileDlg.CancelError = True
  OpenFileDlg.Action = COMMDLG_ACTION_OPEN
  FilenameRes = LCase$(OpenFileDlg.Filename)
L427E:
  ResDataBuff = FilenameRes
  ResDataBuff = ""
  If Len(FilenameRes) Then LoadResData FilenameRes
Exit Sub

L42C0:
  FilenameRes = ""
  Resume L427E
End Sub

Sub control8_Click ()
On Error GoTo L43E2
  OpenFileDlg.Flags = OFN_FILEMUSTEXIST
  OpenFileDlg.DialogTitle = PROGRAM_TO_PROTECT
  OpenFileDlg.Filter = PROGRAM_EXECUTABLES
  OpenFileDlg.Filename = ""
  OpenFileDlg.CancelError = True
  OpenFileDlg.Action = COMMDLG_ACTION_OPEN
  FileInput = LCase$(OpenFileDlg.Filename)
L4360:
  If Len(FileInput) Then
    txt_Filein = FileInput
    txt_FileRes = ""
    txt_FileOut = ""
    ResDataBuff = ""
    control2.Enabled = True
  Else
    control2.Enabled = False
  End If
Exit Sub

L43E2:
  FileInput = ""
  Resume L4360
End Sub

Sub Form_Load ()
'Form2.Show 1
  Started = True

'control2_Click
End Sub

Sub Form_QueryUnload (Cancel As Integer, UnloadMode As Integer)
'Form2.Show 1
End Sub

Sub LoadResData (Filename As String)
Dim hFile As Integer
Dim ResInfoStruct As TYPE_RES_INFO
  If LCase$(Right$(Filename, 3)) <> "res" Then
    MsgBox Filename, VB_GUARD, ERR_NOT_A_RES_FILE
    Exit Sub
  End If

  hFile = FreeFile
  Open Filename For Binary As hFile
  
  Do Until EOF(hFile)
    Get hFile, , ResInfoStruct
    If Asc(ResInfoStruct.M07E0.M07D1) <> &HFF Then Exit Do
    If ResInfoStruct.M07E0.M0463 = &H10 Then
      ResDataBuff = Space$(ResInfoStruct.M0717)
      Get hFile, , ResDataBuff
      Close hFile
      GoTo QuitProc
    End If
    Seek hFile, Seek(hFile) + ResInfoStruct.M0717
  Loop
  
  Close hFile
  
  MsgBox Filename, mc004E, ERR_NO_VER_INFO_FOUND

QuitProc:
End Sub

