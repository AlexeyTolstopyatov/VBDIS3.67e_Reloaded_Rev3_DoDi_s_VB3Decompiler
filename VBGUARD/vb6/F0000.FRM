VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Appearance      =   0  '2D
   BackColor       =   &H00C0C0C0&
   Caption         =   "DoDi's VB Guard"
   ClientHeight    =   3705
   ClientLeft      =   3600
   ClientTop       =   5160
   ClientWidth     =   5190
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
   Icon            =   "F0000.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   3705
   ScaleWidth      =   5190
   Begin VB.CheckBox Chk_DeProtect 
      Caption         =   "DeProtect"
      Height          =   375
      Left            =   240
      MaskColor       =   &H8000000F&
      TabIndex        =   10
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   3255
   End
   Begin VB.Frame SSFrame1 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      Caption         =   "Step 3: Create protected/unprotected copy"
      ForeColor       =   &H00000000&
      Height          =   836
      Left            =   121
      TabIndex        =   0
      Top             =   2200
      Width           =   4972
      Begin VB.CommandButton control2 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         Caption         =   "Destination"
         Height          =   286
         Left            =   121
         TabIndex        =   1
         Top             =   330
         Width           =   1221
      End
      Begin VB.Label txt_FileOut 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   253
         Left            =   1452
         TabIndex        =   2
         Top             =   330
         Width           =   3377
      End
   End
   Begin VB.Frame control4 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      Caption         =   "Step 2: Add Version Info ressource (optional)"
      ForeColor       =   &H00000000&
      Height          =   836
      Left            =   121
      TabIndex        =   7
      Top             =   1155
      Width           =   4972
      Begin VB.CommandButton control5 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         Caption         =   "VersionInfo"
         Height          =   286
         Left            =   121
         TabIndex        =   8
         Top             =   330
         Width           =   1221
      End
      Begin VB.Label txt_FileRes 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   253
         Left            =   1452
         TabIndex        =   9
         Top             =   330
         Width           =   3377
      End
   End
   Begin VB.Frame control7 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      Caption         =   "Step 1: Select the program to protect"
      ForeColor       =   &H80000008&
      Height          =   836
      Left            =   121
      TabIndex        =   4
      Top             =   110
      Width           =   4972
      Begin VB.CommandButton control8 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         Caption         =   "Program"
         Height          =   286
         Left            =   121
         TabIndex        =   5
         Top             =   330
         Width           =   1221
      End
      Begin VB.Label txt_FileIn 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   286
         Left            =   1452
         TabIndex        =   6
         Top             =   330
         Width           =   3377
      End
   End
   Begin MSComDlg.CommonDialog OpenFileDlg 
      Left            =   3751
      Top             =   3190
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.CommandButton control11 
      Appearance      =   0  '2D
      Caption         =   "Done"
      Height          =   432
      Left            =   3720
      TabIndex        =   3
      Top             =   3190
      Width           =   1332
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' F0000.FRM
Option Explicit
Dim FileInput As String
Dim FileOutput As String
Dim FileNameRes As String
Dim ResDataBuff As String
Const COMMDLG_ACTION_OPEN = 1 ' &H1%
Const COMMDLG_ACTION_SAVEAS = 2 ' &H2%
Const OFN_FILEMUSTEXIST = 4096 ' &H1000%
Const VB_GUARD = 16 ' &H10%
Const mc004E = 64 ' &H40%

Private Sub control11_Click()
  Unload Me
End Sub

Private Sub control2_Click()
On Error GoTo ERR_CANCEL
  OpenFileDlg.Flags = OFN_FILEMUSTEXIST
  OpenFileDlg.DialogTitle = SAVE_AS
  OpenFileDlg.Filter = PROGRAM_EXECUTABLES
  OpenFileDlg.FileName = FileOutput
  OpenFileDlg.CancelError = True
  
   OpenFileDlg.Action = COMMDLG_ACTION_SAVEAS
   FileOutput = LCase$(OpenFileDlg.FileName)
   ProtectionEnabled = Chk_DeProtect.Value = vbUnchecked

'Stop
'FileInput = "Z:\vb4\VBGUAR2.EXE" '"Z:\vb4\Vbopt4_.exe"
'FileOutput = "Z:\vb4\VBGUARD_P.EXE" '"Z:\vb4\Vbopt4_2.exe"
'ProtectionEnabled = 0



ERR_Continue:
  If Len(FileOutput) Then
    
    If FileInput = FileOutput Then
      MsgBox ERR_CannotCopytoSameFile, VB_GUARD
      Exit Sub
    End If
    
    txt_FileOut = FileOutput
    
   'Load Filedata
    ReadFileData FileInput
    
    If NE.Signature = 0 Then
      MsgBox FileInput, VB_GUARD, ERR_FileNotFound
    
    Else
      
      On Error Resume Next
      Kill FileOutput
      On Error GoTo ERR_CANCEL
      
     'Store
      WriteFileData FileOutput, ResDataBuff
      
      Dim tmpstr$
      If ProtectionEnabled Then
         tmpstr = PROGRAM_PROTECED
      Else
         tmpstr = "Programm Deprotected!"
      End If
      
         MsgBox FileOutput, mc004E, tmpstr
      
'      txt_FileIn = ""
'      txt_FileOut = ""
'      ResDataBuff = ""
'      control2.Enabled = False
    End If
    
    Close hInFile
    
'Unload Me

  End If
Exit Sub

ERR_CANCEL:
  FileOutput = ""
  Resume ERR_Continue
End Sub

Private Sub control5_Click()
On Error GoTo L42C0
  OpenFileDlg.Flags = OFN_FILEMUSTEXIST
  OpenFileDlg.DialogTitle = VERSIONINFO_RES
  OpenFileDlg.Filter = PROGRAM_RESOURCES
  If Right$(FileInput, 4) = ".exe" Then
    FileNameRes = FileInput
    Mid$(FileNameRes, Len(FileNameRes) - 2) = "res"
  Else
    FileNameRes = ""
  End If
  OpenFileDlg.FileName = FileNameRes
  OpenFileDlg.CancelError = True
  OpenFileDlg.Action = COMMDLG_ACTION_OPEN
  FileNameRes = LCase$(OpenFileDlg.FileName)
L427E:
  txt_FileRes = FileNameRes
  ResDataBuff = ""
  If Len(FileNameRes) Then LoadResData FileNameRes
Exit Sub

L42C0:
  FileNameRes = ""
  Resume L427E
End Sub

Private Sub control8_Click()
On Error GoTo L43E2
  OpenFileDlg.Flags = OFN_FILEMUSTEXIST
  OpenFileDlg.DialogTitle = PROGRAM_TO_PROTECT
  OpenFileDlg.Filter = PROGRAM_EXECUTABLES
  OpenFileDlg.FileName = ""
  OpenFileDlg.CancelError = True
  OpenFileDlg.Action = COMMDLG_ACTION_OPEN
  FileInput = LCase$(OpenFileDlg.FileName)
L4360:
  If Len(FileInput) Then
    txt_FileIn = FileInput
'    txt_FileRes = ""
'    txt_FileOut = ""
'    ResDataBuff = ""
'    control2.Enabled = True
'  Else
'    control2.Enabled = False
  End If
Exit Sub

L43E2:
  FileInput = ""
  Resume L4360
End Sub


Private Sub processCommandline()
   CommandlineMode = Len(Command)
   
   If CommandlineMode Then
      txt_FileIn = Command
      txt_FileOut = Command

   
      control2_Click
   End If
   
   CommandlineMode = 0
   
End Sub


Private Sub Form_Load()
' Form2.Show 1
  Started = True
  
  processCommandline
  
'control2_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Form2.Show 1
End Sub

Private Sub LoadResData(FileName As String)
Dim hInFile As Integer
Dim ResInfoStruct As TYPE_RES_INFO
 
 'Check Extension
  If LCase$(Right$(FileName, 3)) <> "res" Then
    MsgBox FileName, VB_GUARD, ERR_NOT_A_RES_FILE
    Exit Sub
  End If

  hInFile = FreeFile
  Open FileName For Binary As hInFile
  
  Do Until EOF(hInFile)
    Get hInFile, , ResInfoStruct
    
    If Asc(ResInfoStruct.ResID.Type) <> &HFF Then Exit Do
    
    If ResInfoStruct.ResID.TypeAsInt = &H10 Then
      
     'Read In ResData
      ResDataBuff = Space$(ResInfoStruct.ResSize)
      Get hInFile, , ResDataBuff
      Close hInFile
      
      GoTo QuitProc
    End If
    Seek hInFile, Seek(hInFile) + ResInfoStruct.ResSize
  Loop
  
  Close hInFile
  
  MsgBox FileName, mc004E, ERR_NO_VER_INFO_FOUND

QuitProc:
End Sub

