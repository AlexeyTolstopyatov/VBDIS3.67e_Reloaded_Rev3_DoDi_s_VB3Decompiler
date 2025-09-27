'Attribute VB_Name = "Helper_MODULE17"
' Module17
Option Explicit
Declare Function GetModuleUsage Lib "Kernel" (ByVal hModule As Integer) As Integer

Const mc0092 = 256 ' &H100%

Const mc00B6 = 4 ' &H4%
Const mc00C6 = 16 ' &H10%


Const mc00EC = 1 ' &H1%
Const mc00EE = 2 ' &H2%
Const mc00F2 = 4 ' &H4%
Const mc00F4 = 5 ' &H5%
Const mc00F6 = 8 ' &H8%
Const mc00F8 = 9 ' &H9%
Const mc00FA = 6 ' &H6%
Const mc00FC = 7 ' &H7%
Const mc00FE = 10 ' &HA%
Const mc0100 = 11 ' &HB%
Const mc0102 = 12 ' &HC%
Const mc0104 = 13 ' &HD%
Const mc0106 = 14 ' &HE%
Const mc0108 = 15 ' &HF%
Const mc010A = 16 ' &H10%
Const mc010C = 17 ' &H11%
Const mc010E = 18 ' &H12%
Const mc0110 = 19 ' &H13%
Const mc0112 = 20 ' &H14%

Sub ClearLogUntil (ByVal p00D6 As Integer)
25070
  Screen.MousePointer = vbDefault
  While FrmMain.lb_Log.ListCount >= p00D6
    If FrmMain.lb_Log.ListCount = 0 Then Exit Sub
    FrmMain.lb_Log.RemoveItem FrmMain.lb_Log.ListCount - 1
  Wend
End Sub

Sub DispLog2 (p00D0 As Integer, ByVal p00D2 As String)
Dim LogCount As Integer

25060
  LogCount = FrmMain.lb_Log.ListCount
  If p00D0 > LogCount Then
    p00D0 = LogCount
  Else
    While p00D0 < FrmMain.lb_Log.ListCount
      FrmMain.lb_Log.RemoveItem p00D0
    Wend
  End If
  FrmMain.lb_Log.List(p00D0 - 1) = p00D2
  sub0E79 p00D2
End Sub

Sub DispStatus (LogListCount As Integer, Text As String)
25080
  FrmMain.lb_Log.AddItem Text
  LogListCount = FrmMain.lb_Log.ListCount
  
'  Stop
  sub0E79 Text
End Sub

Private Sub sub0E79 (Dummy As String)
25050
  Screen.MousePointer = vbHourglass
  DoEvents
End Sub

Sub sub0E83 (LogListCount As Integer)
Static m00E4 As Integer
 
25090
  If Len(AppPathwithSlash) Then Exit Sub
  InitPaths LogListCount
  sub0BE9 LogListCount
  sub0B3E LogListCount
  gv0EE2(1) = -22 - 2
  gv0EE2(2) = -22 - 4
  gv0EE2(3) = gv0EE2(2)
  gv0EE2(4) = -22 - 8
  gv0EE2(5) = gv0EE2(4)
  gv0EE2(6) = -22 - &H10
  gv0EE2(7) = 1
  fnPreFix(m00E4) = "g?"
  fnPreFix(&H1) = "gn"
  fnPreFix(&H2) = "gv"
  fnPreFix(&H3) = "gc"
  fnPreFix(&H4) = "T"
  fnPreFix(&H5) = "M"
  fnPreFix(&H8) = "extsub"
  fnPreFix(&H9) = "extfn"
  fnPreFix(&H6) = "sub"
  fnPreFix(&H7) = "fn"
  fnPreFix(&HA) = "event"
  fnPreFix(&HB) = "v?"
  fnPreFix(&HC) = "v"
  fnPreFix(&HD) = "m"
  fnPreFix(&HE) = "g"
  fnPreFix(&HF) = "f"
  fnPreFix(&H10) = "l"
  fnPreFix(&H11) = "p"
  fnPreFix(&H12) = "pv"
  fnPreFix(&H13) = "mc"
  fnPreFix(&H14) = "o?"
  fnPreFix(&H15) = "f?"
End Sub

Sub MakeErrBeep ()
25000
    Beep
End Sub

'33   0000 0000 0010 0001
' rol 3               <<<
'264  0000 0001 0000 1000
Function RotateLeft3 (ByVal pv0056 As Integer) As Integer
25030
  RotateLeft3 = ToSInt16(Uint32(pv0056 And &H1FFF) * 8 + ((pv0056 \ &H1000) And 6))
End Function

Function RotateRight3 (pv0048%) As Integer
'Exsample
' 136 0000 0000 1000 1000  <- Input
'       ror 3    >>>
' 17  0000 0000 0001 0001  -> Output

25010
  '              >>3        and 0001 1111  +           ...0110 <<12
  RotateRight3 = (UDiv(pv0048, 8) And &H1FFF&) + ((pv0048 And 6) * &H1000)
End Function

Sub StartDecompile (Cmd_p005C As Integer, frm_p0060 As Form)
Dim RetVal As Integer
Dim Mak_Filename As String
Dim statflags As Integer
Dim Disp_Lvl%

On Error GoTo RunTErr

25040
  Select Case Cmd_p005C
  
  Case b0001_FileNamesSet
    gDecompileStatus = 0
  
  Case b0100_ScanDone
    gDecompileStatus = gDecompileStatus And Not (b0100_ScanDone Or b0001_0000_DecompilationDone Or b0010_0000_FormsCombined)
  
  Case b0001_0000_DecompilationDone
    'Redo FormsCombined Decompilation
    gDecompileStatus = gDecompileStatus And Not (b0010_0000_FormsCombined Or b0001_0000_DecompilationDone)
  
  Case b0010_0000_FormsCombined
    If gDecompileStatus And Cmd_p005C Then
      gDecompileStatus = gDecompileStatus And Not b0001_0000_DecompilationDone
    End If
  
  Case Else
    If Cmd_p005C And gDecompileStatus Then
      MsgBox "Already done"
      Exit Sub
    End If
  
  End Select
  
  statflags = Cmd_p005C Or (Cmd_p005C - 1)
  statflags = statflags And Not gDecompileStatus
25041
  If statflags And b0001_FileNamesSet Then
    sub06C8
    gWorkDir = ""
    frm_p0060.Caption = "Select EXE file"
    Screen.MousePointer = vbDefault
    FrmOpenFile.Tag = "*.EXE"
'***********************************************

   If CommandlineMode Then
      Mak_Filename = Command$
   Else
     FrmOpenFile.Show 1
     Mak_Filename = FrmOpenFile.Tag
   End If

'    Mak_Filename = "M:\t\VBDIS3.59e Pro\VBPRO3.EXE"
'    Mak_Filename = "M:\t\VBXFOOL.exe"
    
    If Len(Mak_Filename) = 0 Then Exit Sub
    
    gIn_FileFulPath = Mak_Filename
    
    RetVal = ExtractPath(Mak_Filename, "\")
    If RetVal Then
      gCurPath = Left$(Mak_Filename, RetVal)
      Mak_Filename = Mid$(Mak_Filename, RetVal + 1)
    Else
      MakeErrBeep
    End If
    
    gIn_FileNameAndExt = Mak_Filename
    RetVal = InStr(Mak_Filename, ".")
    
    If RetVal Then Mak_Filename = Left$(Mak_Filename, RetVal - 1)
    NewPrjName = Mak_Filename
    
    gDecompileStatus = gDecompileStatus Or b0001_FileNamesSet
    
    
  ' Set new Caption
    frm_p0060.Caption = "VBDis " & NewPrjName
    
    
    FrmMain.lb_Log.Clear
    
  End If
25042
  
  If statflags And b0010_ExeLoaded Then
    FrmMain.Show
    FrmMain.dum800.Value = 1
    DispStatus Disp_Lvl, "Loading " & gIn_FileNameAndExt
    
  ' Read MZ and NE Structs
    ReadFileData
25043
'    If VBver_11 Then
'
'      If Segs(4).Flags And RELOCINFO Then VBver_11 = 0
'
'    ' Segments must be smaller than 100
'      For RetVal = 3 To Segments
'        If Segs(RetVal).size < 0 Then
'          VBver_11 = 0: Exit For
'        End If
'      Next
'
'    ' Not bigger than 32kb
'      For RetVal = Res_ID_gv0DC0 To ResObjectsCount
'        If PageToOffset(ResObjects1(RetVal).size, ResAlignment) > &H7FFF Then
'          VBver_11 = 0: Exit For
'        End If
'      Next
'
'      If VBver_11 = 0 Then
'        ShowErrMsg "The program is too big for this version of VB Discompiler"
'        End
'      End If
'    End If



25044
    If VBver_11 Then
      
      LoadVBDisDatFiles Disp_Lvl
      
      If Cancel Then
        VBver_11 = 0
        gDecompileStatus = 0
      End If
      
      gv120C = False
      
      LoadVBX Disp_Lvl, "VBRUN" & Format(VBver_11) & "00.VBX"
    
    Else
      MsgBox "Not a Visual Basic program"
      gDecompileStatus = 0
    
    End If
    
    Screen.MousePointer = vbDefault
    ClearLogUntil Disp_Lvl
2
5045
    If gDecompileStatus Then
      gDecompileStatus = gDecompileStatus Or b0010_ExeLoaded
      frm_p0060.Caption = "VBDis " & gIn_FileNameAndExt
      FrmPrjDir.txt_newDir = NewPrjName
      FrmPrjDir.DirListBox = gCurPath

'***************************
      If CommandlineMode Then
         On Error Resume Next
         Dim tmpint&
         ChDrive gCurPath
         ChDir gCurPath
         MkDir Mak_Filename
         gWorkDir = gCurPath & Mak_Filename & "\"
         ChDir gWorkDir
         On Error GoTo RunTErr
      Else
         FrmPrjDir.Show 1
      End If
     
     ' gWorkDir = "M:\t\1\VBPDIS3\"
    
    
    Else
      frm_p0060.Caption = "Select EXE file"
      Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
  End If
  
  If statflags And b0100_ScanDone Then
    DispStatus Disp_Lvl, Scanning & gIn_FileNameAndExt
    ScanExe
    FrmMain.dum700.Value = 1

25046
    ClearLogUntil Disp_Lvl
    If gv7B32 And mc00B6 Then
      gDecompileStatus = 0
      Screen.MousePointer = vbDefault
      Exit Sub
    Else
      gDecompileStatus = gDecompileStatus Or b0100_ScanDone
    End If
  End If
25047
  
  On Error GoTo 0
  If statflags And b1000_SubroutinesScanned Then
    If (gDecompileStatus And b1000_SubroutinesScanned) = 0 Then
      FrmMain.lb_Log.Clear
      DispStatus Disp_Lvl, Scanning & CurFrmName
      ScanSubroutineCalls
      gDecompileStatus = gDecompileStatus Or b1000_SubroutinesScanned
      ClearLogUntil Disp_Lvl
    End If
  End If
  
25048
  If statflags And b0001_0000_DecompilationDone Then
    
    DispStatus Disp_Lvl, "Creating project files"
    
    CreatePrjFiles
    
    DispLog2 Disp_Lvl, "Creating project files"
    gDecompileStatus = gDecompileStatus Or b0001_0000_DecompilationDone
    bPrjFilesCreated = True
    
    Mak_Filename = PathAddExt("", "mak")
    
    RetVal = 0
    On Error Resume Next
     
    ' Launch frms2txt
      If Len(bPrjFilesCreated) Then
        
        RetVal = Shell(bPrjFilesCreated & " " & Mak_Filename, 1)
        
      Else
        
        RetVal = Shell(AppPathwithSlash & "frms2txt " & Mak_Filename)
        
        If RetVal Then
          If GetModuleUsage(RetVal) = 1 Then
            
            While GetModuleUsage(RetVal) = 1
              DoEvents
            Wend
            
          Else
            RetVal = 0
          End If
          
        End If
        
      End If
      
      
    On Error GoTo 0
    If RetVal = 0 Or Err Then
      statflags = 0
      frm_p0060.Caption = Mak_Filename
      
      MsgBox Mak_Filename & Chr$(10) & "Now save *.FRM as text, please", mc00C6, "Project created"
      
      
    Else
      statflags = b0010_0000_FormsCombined
    End If
  End If
25049
  
 'Combine Forms
  If statflags And b0010_0000_FormsCombined Then
    FrmMain.lb_Log.Clear
    DispStatus Disp_Lvl, gc03C6
    gDecompileStatus = gDecompileStatus Or b0010_0000_FormsCombined
    CombineFrms
    bPrjFilesCreated = True
    
    If gDecompileStatus And b0010_0000_FormsCombined Then
      DispLog2 Disp_Lvl, gc03CA
    Else
      DispLog2 Disp_Lvl, "Now save *.FRM as text, please"
    End If
  
  End If
  Screen.MousePointer = vbDefault
Exit Sub

RunTErr:
  ShowRuntimeErr
Exit Sub
End Sub

Function VBAbsOffs (ByVal Value As Integer) As Long
Dim FirstThreeByte As Integer

25020
 '               >>13       and ..0011
  FirstThreeByte = (Value \ &H2000) And 3 '         and 0001 1111 << 3
  '                                    +
  VBAbsOffs = 1 + GetSegOffset(FirstThreeByte + VBStartSeg) + CLng(Value And &H1FFF) * 8
End Function

