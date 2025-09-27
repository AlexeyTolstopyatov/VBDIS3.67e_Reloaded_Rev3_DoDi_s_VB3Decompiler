VERSION 5.00
Begin VB.Form FrmMain 
   Appearance      =   0  '2D
   BackColor       =   &H80000005&
   Caption         =   "DoDi's VB Discompiler"
   ClientHeight    =   3630
   ClientLeft      =   2610
   ClientTop       =   1080
   ClientWidth     =   9285
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
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   3630
   ScaleWidth      =   9285
   Begin VB.CommandButton cmd_Browse 
      Caption         =   "..."
      Height          =   345
      Left            =   8520
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox Txt_File 
      Appearance      =   0  '2D
      Height          =   285
      Left            =   3240
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.ComboBox dum200 
      Appearance      =   0  '2D
      Height          =   3105
      Left            =   3240
      Style           =   1  'Einfaches Kombinationsfeld
      TabIndex        =   3
      Top             =   540
      Width           =   1812
   End
   Begin VB.CommandButton dum300 
      Appearance      =   0  '2D
      Caption         =   "+"
      Height          =   252
      Left            =   8760
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.CommandButton dum400 
      Appearance      =   0  '2D
      Caption         =   "+"
      Height          =   252
      Left            =   6720
      TabIndex        =   11
      Top             =   300
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.CommandButton dum500 
      Appearance      =   0  '2D
      Caption         =   "+"
      Height          =   252
      Left            =   4800
      TabIndex        =   10
      Top             =   300
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.ListBox lb_Log 
      Appearance      =   0  '2D
      Height          =   3345
      Left            =   120
      TabIndex        =   9
      Top             =   60
      Width           =   3015
   End
   Begin VB.CommandButton dum700 
      Appearance      =   0  '2D
      Caption         =   "done"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3840
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.CommandButton dum800 
      Appearance      =   0  '2D
      Caption         =   "scan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3240
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.ComboBox dum900 
      Appearance      =   0  '2D
      Height          =   3105
      Left            =   7080
      Style           =   1  'Einfaches Kombinationsfeld
      TabIndex        =   5
      Top             =   540
      Width           =   2052
   End
   Begin VB.ComboBox duma00 
      Appearance      =   0  '2D
      Height          =   3105
      Left            =   5160
      Style           =   1  'Einfaches Kombinationsfeld
      TabIndex        =   4
      Top             =   540
      Width           =   1812
   End
   Begin VB.Label dumb00 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      Caption         =   "Controls"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   7080
      TabIndex        =   2
      Top             =   300
      Width           =   1812
   End
   Begin VB.Label dumc00 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Subroutines"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   5160
      TabIndex        =   1
      Top             =   300
      Width           =   1524
   End
   Begin VB.Label dumq00 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      Caption         =   "Modules"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   3240
      TabIndex        =   0
      Top             =   300
      Width           =   1344
   End
   Begin VB.Menu mi_Open 
      Caption         =   "&Open..."
   End
   Begin VB.Menu dumw00 
      Caption         =   "&File"
      Begin VB.Menu mi_Improve 
         Caption         =   "Im&prove"
         Enabled         =   0   'False
      End
      Begin VB.Menu mi_CombineForms 
         Caption         =   "&Combine Forms"
         Enabled         =   0   'False
      End
      Begin VB.Menu a200 
         Caption         =   "Start &VB"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu a300 
         Caption         =   "-"
      End
      Begin VB.Menu a400 
         Caption         =   "&Info..."
      End
      Begin VB.Menu a500 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu Options 
      Caption         =   "O&ptions"
      Visible         =   0   'False
      Begin VB.Menu mi_VB6 
         Caption         =   "Mak &VB6-Project_files"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' F0000.FRM
Option Explicit
Dim BoxesCleared As Integer



Private Sub dum500_Click()
42010
  dum200.text = ModulFileNames(NewFileExt)
  DoEvents
End Sub

Private Sub dum400_Click()
42020
  duma00.AddItem getFrmNamebyOffbyIdx(g_CurSub)
  DoEvents
End Sub

Private Sub dum700_Click()
42030
  BoxesCleared = False
End Sub

Private Sub dum800_Click()
42040
  BoxesCleared = 1
  Caption = gIn_FileFulPath
  dum200.Clear
  duma00.Clear
  dum900.Clear
  DoEvents
End Sub

Private Sub sub1080(p0044 As ComboBox)
Static m0046 As Integer

42000
  If m0046 Then Exit Sub
  m0046 = True
  If BoxesCleared Then
    p0044.AddItem p0044.text
    p0044.ListIndex = p0044.NewIndex
    DoEvents
  Else
    Beep
  End If
  m0046 = False
End Sub

Private Sub dum900_Change()
42050
  sub1080 dum900
End Sub

Private Sub dum200_Change()
42060
  duma00.Clear
  sub1080 dum200
End Sub

Private Sub duma00_Change()
42070
  sub1080 duma00
End Sub
Private Sub processCommandline()
   CommandlineMode = Len(Command)
   
   If CommandlineMode Then
      mi_Open_Click
   End If
   
   CommandlineMode = 0
   
End Sub

Private Sub Form_Load()

'frm_ProvideCtrlNames.Show
'Stop
'Exit Sub
 Dim LogListCount%

42080
  Set CurForm_gv0BC0 = Me
  
  Show
  
  DispStatus LogListCount, sInitializing
  
  BoxesCleared = 1
  sub0E83 LogListCount
  ClearLogUntil LogListCount
  
  processCommandline
'  mi_Open_Click
End Sub

Private Sub Form_Unload(gCancel As Integer)

42090
  End
End Sub

Private Sub a400_Click()

42100
  frmAbout.Show 1
End Sub

Private Sub mi_CombineForms_Click()

42110
  StartDecompile b0010_0000_FormsCombined, Me
End Sub

Private Sub a500_Click()

42120
  End
End Sub

Private Sub mi_Open_Click()

42130
  mi_CombineForms.Enabled = False
  
  gDecompileStatus = 0
  
  StartDecompile b0001_0000_DecompilationDone, Me
  
  If gDecompileStatus And b0001_0000_DecompilationDone Then
    mi_Improve.Enabled = True
    mi_CombineForms.Enabled = True
  End If
End Sub

Private Sub mi_Improve_Click()
42140
' Redo Decompilation & SubroutinesScan
  gDecompileStatus = gDecompileStatus And Not (b0001_0000_DecompilationDone Or b1000_SubroutinesScanned)
  
  StartDecompile b0001_0000_DecompilationDone, Me
  
End Sub

Private Sub mi_VB6_Click()
   mi_VB6.Checked = Not (mi_VB6.Checked)
End Sub
