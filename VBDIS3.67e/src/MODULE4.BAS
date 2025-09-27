Attribute VB_Name = "MODULE4"
' Module4
Option Explicit
Dim m000E As T149E
Dim FileBuff_Mod As String
Dim FrmStruct2 As T14C9
Dim hFile_Mod As Integer
Const mc004E = 64 ' &H40%

Const mc0086 = 128 ' &H80%
Const Type_Form = 67 ' &H43%
Const Type_VBX = 64 ' &H40%
Const mc0114 = 8 ' &H8%
Const mc0152 = 36 ' &H24%

Const mc018C = 1 ' &H1%
Const mc0194 = -2 ' &HFFFE%
Const mc0198 = 224 ' &HE0%
Const mc019A = 9 ' &H9%
'Const mc01DA = 1 ' &H1%

Sub CombineFrms()
Dim i%
Dim ModDataLine As String
Dim hFileLocal As Integer
Dim hModFile As Integer
Dim ModFilename As String
Dim FrmFilename As String
Dim l0044%
Dim l0046%


6000
  For i = 1 To FrmFilesCount
    ModDataLine = ModulFileNames(i)
    
   'Process only ".FRM" files
    If Right$(ModDataLine, 4) = ".FRM" Then
      FrmFilename = PathAddExt(ModDataLine, "FRM")
      ModFilename = PathAddExt(ModDataLine, "BAS")
      
      If Len(Dir$(ModFilename)) Then
        hFileLocal = FreeFile
        DispStatus l0044, FrmFilename
        
       'Check is frm was saved as 'text'
        Open FrmFilename For Binary As hFileLocal
        ModDataLine = Input$(7, hFileLocal)
        If ModDataLine = "VERSION" Then

        ' Open Frm from Append
          Close hFileLocal
          Open FrmFilename For Append As hFileLocal
          hModFile = FreeFile
          DispStatus l0046, " + " & ModFilename
          
        ' Append ModFileData to Frm
          Open ModFilename For Input As hModFile
          While Not EOF(hModFile)
            Line Input #hModFile, ModDataLine
            Print #hFileLocal, ModDataLine
          Wend
          Close hModFile
         
        ' and ModFile
          Kill ModFilename
          
        Else
        
          MsgBox FrmFilename, 64, "Form must be saved 'As Text'"
          
         'Reset gDecompileStatus
          gDecompileStatus = gDecompileStatus And Not b0010_0000_FormsCombined
        End If
        
        Close hFileLocal
        ClearLogUntil l0046
        ClearLogUntil l0044
      End If
    End If
  Next
End Sub

Sub DoDecompile(bCreateOutput As Integer)
Dim lFileName As String
Dim i%
Dim j%
Dim Controls As Integer
Dim Displvl1%
Dim Displvl2%



6010
  If VBver_11 <> 2 And VBver_11 <> 3 Then
    ShowErrMsg 4
    Exit Sub
  End If
  
  If bCreateOutput Then
    DispStatus Displvl1, "Creating " & "*.FRM"
    DispStatus Displvl2, "Loading " & "empty.frm"
    
  Dim Buff$
  Buff = AppPathwithSlash & "empty.frm"
  
  While Len(Dir$(Buff, vbNormal)) = 0
    
    'If (FileLen(buff) = 0) Then Kill buff
    
    Dim oldMousePointer%
    oldMousePointer = Screen.MousePointer
    Screen.MousePointer = vbDefault
      
      MsgBox Buff, vbExclamation, "File not found"  '48
      
      FrmOpenFile.Tag = "empty.frm"
      FrmOpenFile.Show 1
      
    Screen.MousePointer = oldMousePointer
    Buff = FrmOpenFile.Tag
    
    If Len(Buff) = 0 Then End
    
  Wend
    
    
    
    
    
    
    hFile_Mod = FreeFile
    Open Buff For Binary Access Read As hFile_Mod
    
    Get hFile_Mod, 1, m000E
    
    FileBuff_Mod = Space$(LOF(hFile_Mod) - m000E.M14C1 - 5)
    
    Get hFile_Mod, m000E.M14C1 + 6, FileBuff_Mod
    Close hFile_Mod
    
    
    
    
    
    
    If VBver_11 = 2 Then
      Mid$(m000E.M14AE, 3, 1) = Chr$(&H2B)
    End If
    
    ClearLogUntil Displvl2
    
  Else
    DispStatus Displvl1, Scanning & "*.FRM"
  End If
  
  For i = 1 To gFormsCount
    
    If gControlItems(i).Type_3 = &H46 Then
      DispStatus Displvl2, "Loading " & "RC" & Hex$(gControlItems(i).ResIDAssoc)
      SeekToRCData gControlItems(i).ResIDAssoc
      
      If curResID = 0 Then
        std_p_e = VBControlNames(i)
        ShowErrMsg 128
      Else
        
        Get hVBFile, CurResOffsetStart, m000E
        
        Controls = Asc(m000E.M14B9)
        
        If bCreateOutput Then
          gFrmCount2 = gVBObjsDefs(gControlItems(i).M_2_curr).mFrmCount
          lFileName = gWorkDir & VBControlNames(i)
          DispLog2 Displvl2, "Creating " & VBControlNames(i)
          
          Open lFileName For Binary As hFile_Mod
          Put hFile_Mod, , m000E
        
        Else
          DispLog2 Displvl2, Scanning & VBControlNames(i)
          
          gControlCount1 = gControlCount1 + 1
          gFrmCount2 = gControlCount1
          
          ReDim Preserve gv0FF6(LBound(gv0FF6) To gControlCount1 + Controls)
          If gVBObjsDefs(gControlItems(i).M_2_curr).mFrmCount <> gControlItems(i).ResIDAssoc Then
            MakeErrBeep
          End If
          
          gVBObjsDefs(gControlItems(i).M_2_curr).mFrmCount = gControlCount1
          
        ' Check if the FrmRes2 (the one with the FormName is present)
          SeekToRCData gControlItems(i).ResIDAssoc + 1


          If curResID = 0 Then
           '... so 'VBGuard' has delete FrmRes2
            'ShowErrMsg "Cheat!"
           
'            End ' <- :P
          
L_CreateControllNames:

            ' Wups they got deleted so Generate ControlNames
            
              On Error Resume Next
              Dim oldMousePointer2
              oldMousePointer2 = Screen.MouseIcon
              Screen.MousePointer = MousePointerConstants.vbDefault

                frm_CtrlNames.ctrls_Filename = gWorkDir & VBControlNames(i) & ".txt" 'AppPathwithSlash & "myControlNames.txt"
                frm_CtrlNames.ControlsCount = Controls
               
               'Attention frm_CtrlNames.caption will invokes Form_Load !
                frm_CtrlNames.FormName = VBControlNames(i)
                frm_CtrlNames.Show vbModal

              Screen.MousePointer = oldMousePointer2
              On Error GoTo 0
            
            
'              For j = 0 To Controls
'                lFileName = "control" & Format$(j)
'
'                gv0FF6(gControlCount1).Name_4 = lFileName
'                gControlCount1 = gControlCount1 + 1
'              Next j
          
          Else
            
            ' Load  ControlNames from FrmRes2
              For j = 0 To Controls
                
                Get hVBFile, , FixedByteBuff
                If IsInRange(Asc(FixedByteBuff), 40, 1) Then
                
                  lFileName = Space$(Asc(FixedByteBuff))
                  Get hVBFile, , lFileName

                  
                Else
                  
                ' A VBName shouldn't be bigger than 40 Chars
                  If Asc(FixedByteBuff) > 40 Then
                     lFileName = Space$(Asc(10))
                     Get hVBFile, , lFileName

                     MsgBox "Whoops Crappy ControlName in Exe @ 0x" & Hex(Loc(hVBFile) - 1) & " Resource[" & curResID & "]" & vbCrLf & "Data: " & lFileName, vbCritical, "Fallingback to no FormName Resources..."
                     gControlCount1 = gControlCount1 - j
                     GoTo L_CreateControllNames
                  Else
                     lFileName = "control" & Format$(j)
                  End If
                  
                End If
                
                gv0FF6(gControlCount1).Name_4 = lFileName
                gControlCount1 = gControlCount1 + 1
              Next j
          End If
          
          gVBObjsDefs(gControlItems(i).M_2_curr).Name_4 = gv0FF6(gFrmCount2).Name_4
          SeekToRCData gControlItems(i).ResIDAssoc
        End If
        
        Get hVBFile, CurResOffsetStart, m000E
        
        sub0385 bCreateOutput
        
        Get hVBFile, , FixedByteBuff
        If bCreateOutput Then Put hFile_Mod, , FixedByteBuff
        
        If Asc(FixedByteBuff) = 1 Then
          sub0385 bCreateOutput
          
          Do
            Get hVBFile, , FixedByteBuff
            
            If bCreateOutput Then Put hFile_Mod, , FixedByteBuff
            
            Select Case Asc(FixedByteBuff)
            Case 1, 3
              sub0385 bCreateOutput
            Case 2
            Case 4, 5
              Exit Do
            Case Else
              std_p_e = VBControlNames(i)
              ShowErrMsg mc0086
              Exit Do
            End Select
          Loop
        End If
        
        If Asc(FixedByteBuff) = 5 Then
          
          sub0385 bCreateOutput
          
          Do
            Get hVBFile, , FixedByteBuff
            If bCreateOutput Then Put hFile_Mod, , FixedByteBuff
            Select Case Asc(FixedByteBuff)
            Case 2
              sub0385 bCreateOutput
            Case 3
            Case 4
              Exit Do
            Case Else
              std_p_e = VBControlNames(i)
              ShowErrMsg mc0086
              Exit Do
            End Select
          Loop
        End If
        
        If Asc(FixedByteBuff) <> 4 Then
          ShowErrMsg mc0086
        End If
        
        If bCreateOutput Then
          m000E.M14C1 = Seek(hFile_Mod) - 6
          Put hFile_Mod, , FileBuff_Mod
          Put hFile_Mod, 1, m000E
          Close hFile_Mod
        End If
      End If
      
      ClearLogUntil Displvl2
    
    End If
    
  Next i
  
  ClearLogUntil Displvl1

End Sub

Sub sub03C1()
End Sub

Function fn03D3(p00AE%) As String
Dim l00B0 As String
Dim l00B4 As T2237

6040
  Get hVBFile, gv0B92 + Uint32(p00AE), l00B4
  l00B0 = Space$(Asc(l00B4.M2249))
  Get hVBFile, , l00B0
  fn03D3 = l00B0
End Function

Function fn03DF(p00BA As Integer) As String
Dim l00BC As String

6050
  l00BC = fn034F(gv0B9E, p00BA - 1)
  If InStr(l00BC, " ") Or InStr("%&!#@?$", Right$(l00BC, 1)) Then
    l00BC = "[" & l00BC & "]"
  End If
  fn03DF = l00BC
End Function

Sub ScanRCFrmData()
Dim IntBuff2 As Integer
Dim ControlType As Integer
Dim l00C8
Dim IntBuff As Integer
Dim VBXName As String
Dim bytelen As Integer
Dim l00D0 As Integer
Dim CurFilePos As Long
Dim VBPrjStruct As T1267
Dim VBMain As VB_Main_Struct

Const VBForm = &H46    'F
Const VBX = &H43       'C
Const VBControl = &H58 'X
Const Type_VBX = 64 ' &H40%

Dim FrmRes1Pos As Long
 
 
 Const mc0114 = 8 ' &H8%
Dim l0116() As Integer

6060
 'GetMain
  Get hVBFile, CurResOffsetStart, VBPrjStruct
  Init_gVBObjsDefs
  
  gFormsCount = 0: VBX_Names_Count = 0: gControlItemsCount = 0
  
  Start_Form_Index = 0
  
  If VBPrjStruct.M1287 = -1 Then
    gSubMainIdx = 0
  Else
    gSubMainIdx = VBPrjStruct.M1287 And -2
  End If
  
  gv0FD2 = 0
  gVBTitle = strNullCut(VBPrjStruct.M128E)
  Proj_Title = ReadFixedString(VBPrjStruct.M1299)
  Proj_HelpFile = ReadFixedString(VBPrjStruct.M12A4)
  
  ReDim gv1042(0), gv1076(0)
  Do
    Get hVBFile, , VBMain
    
   'Finished ?
    If Asc(VBMain.Type) = 0 Then Exit Sub
    
   'Store ctrlItems
    gControlItemsCount = gControlItemsCount + 1
    LSet gControlItems(gControlItemsCount) = VBMain

    With gControlItems(gControlItemsCount)
      .Type_3 = Asc(VBMain.Type)
      .ResIDAssoc = VBMain.ResIDAssoc
      .M_2_curr = VBMain.M_2_curr
      .M_3_next = VBMain.M_3_next
      .M_4_sub = VBMain.M_4_sub
      .M_5_Size = VBMain.M_5_Size
 
     
    End With

    
   'get Type
    ControlType = gControlItems(gControlItemsCount).Type_3 And &HFF
    gControlItems(gControlItemsCount).Type_3 = ControlType
    
    Select Case ControlType
    
    
    Case VBX, VBForm
      
      gFormsCount = gFormsCount + 1
      VBXName = ReadFixedString(Asc(VBMain.length))
      
      If ControlType = VBForm Then
        
      ' Store CurrentPos
        CurFilePos = Seek(hVBFile)
        
       'Handle Protected Files, which have Blanked FormNames
        Dim FormCounter%
        FormCounter = FormCounter + 1
        If VBXName = "" Then
           VBXName = "frm" & FormCounter & ".FRM"
        End If
        
        
       'Seek to assoc Res and to the following Res (the one with the FormNames)
        SeekToRCData VBMain.ResIDAssoc: FrmRes1Pos = CurResOffsetStart
        SeekToRCData VBMain.ResIDAssoc + 1
       
      ' If seek Fail or on Res1 don't follow Res2...
        If curResID = 0 Or FrmRes1Pos >= CurResOffsetStart Then
         'Error
          'FrmFilesCount = 0
        End If
        
      ' Restore CurrentPos
        Seek hVBFile, CurFilePos
        
        
       'SecurityCheck: VBXName ends with ".FRM"?
        If UCase$(Right$(VBXName, 4)) <> ".FRM" Then
          std_p_e = VBXName
          ShowErrMsg mc0086
        End If
        
        If Start_Form_Index = 0 Then
          Start_Form_Index = gControlItemsCount
          Start_Form_Res = VBMain.M_5_Size
        End If
        
        If VBMain.M_2_curr > gv0FD2 Then gv0FD2 = VBMain.M_2_curr
        
        If Asc(VBMain.M_4_sub) And 1 Then
          gVBObjsDefs(VBMain.M_2_curr).id_5 = &H14 + 1
        Else
          gVBObjsDefs(VBMain.M_2_curr).id_5 = &HD + 1
        End If
        
        gVBObjsDefs(VBMain.M_2_curr).mFrmCount = VBMain.ResIDAssoc
        
      Else
        
        VBX_Names_Count = VBX_Names_Count + 1
        
        sub0418
        
        VBXName = UCase$(VBXName)
        
        LoadVBX 0, VBXName
        
      End If
      
      VBControlNames(gControlItemsCount) = VBXName
      
    Case VBControl
      
      bytelen = Asc(VBMain.length)
      
      If bytelen = &HFF Then
      
        Get hVBFile, , FixedByteBuff
        VBControlNames(gControlItemsCount) = ReadFixedString(Asc(FixedByteBuff))
        
        For bytelen = gKnowControlsCount To 1 Step -1
          If gKnowControlsData(bytelen).ControlTypeName = VBControlNames(gControlItemsCount) Then Exit For
        Next
        
       'Was Controlname found in gKnowControlsData
        If bytelen = 0 Then
          std_p_e = VBControlNames(gControlItemsCount)
          gv120C = True
          ShowErrMsg Type_VBX
          gKnowControlsData(0).ControlTypeName = "?"
        End If
        
      Else
        bytelen = bytelen + 1
        VBControlNames(gControlItemsCount) = gKnowControlsData(bytelen).ControlTypeName
      End If
      
      gControlItems(gControlItemsCount).M2551 = bytelen
      If gVBObjsDefs(VBMain.M_2_curr).id_5 Then
        If gVBObjsDefs(VBMain.M_2_curr).id_5 <> bytelen Then
          std_p_e = VBControlNames(gControlItemsCount)
          ShowErrMsg Type_VBX
        End If
      Else
        gVBObjsDefs(VBMain.M_2_curr).id_5 = bytelen
        gVBObjsDefs(VBMain.M_2_curr).Name_4 = gKnowControlsData(bytelen).ControlTypeName
      End If
      
      If VBMain.M_3_next <> &H1FF Then
        
        gVBObjsDefs(VBMain.M_3_next).id_5 = -bytelen
        gVBObjsDefs(VBMain.M_3_next).Name_4 = gKnowControlsData(bytelen).ControlTypeName
      
      End If
    Case Else
      ShowErrMsg mc0114
      Exit Do
    End Select
    gControlItems(gControlItemsCount).Type_3 = ControlType
    
    Select Case ControlType
    
    Case VBForm, VBControl
      
      Get hVBFile, , IntBuff
      
      If IntBuff > 0 Then
        ReDim l0116%(IntBuff)
        For l00C8 = 1 To IntBuff
          Get hVBFile, , l0116(l00C8)
        Next
        For l00C8 = 1 To IntBuff
          l00D0 = l0116(l00C8)
          Get hVBFile, , IntBuff2
          
          If ControlType = VBForm Then
            If l00D0 > UBound(gv1076) Then ReDim Preserve gv1076(l00D0)
            If gv1076(l00D0).M2651 = 0 Then
              gv1076(l00D0).M2658 = IntBuff2
              gv1076(l00D0).M2651 = VBMain.M_2_curr
            End If
          Else
            If l00D0 > UBound(gv1042) Then ReDim Preserve gv1042(l00D0)
            If gv1042(l00D0) = 0 Then
              gv1042(l00D0) = gv1478(gKnowControlsData(bytelen).M2BBB + IntBuff2)
            End If
          End If
        Next
      End If
    End Select
  Loop
End Sub

Sub ScanningFrms()
Dim i
Dim rawSize As Integer
Dim l0126 As Integer
Dim size As Integer
Dim dispstat%

6070
  DispStatus dispstat, "Scanning forms"
  
  For i = 1 To gFormsCount
    
    If gControlItems(i).Type_3 <> Type_Form Then
      
      rawSize = gControlItems(i).M_5_Size
      size = RotateRight3(rawSize)
      
      l0126 = NamesLookUpTbl(size + 1)
      SetFrmName l0126, (VBControlNames(i))
      
      gFrmStruct(l0126).M24D0 = gControlItems(i).Type_3
      gFrmStruct(l0126).M24DA = Asc(gControlItems(i).M_4_sub)
      gFrmStruct(l0126).ResIDAssoc = gControlItems(i).ResIDAssoc
      gFrmStruct(l0126).M_2_curr = gControlItems(i).M_2_curr
      gControlItems(i).M2548 = l0126
    
    Else
      FrmMain.dum900.Text = VBControlNames(i)
    
    End If
  
  Next i
  
  ClearLogUntil dispstat
  
End Sub

Sub ScanGlobalData()
Dim i As Long
Dim fBuff1%
Dim fBuff2%
Dim DispLvl%

6080
  DispStatus DispLvl, Scanning & C_GLOBAL & gc03FA
  curResID = 0
  
  ResSeek ResID_RCData + 1
  
  Base1 = CurResOffsetStart
  Get hVBFile, Base1 + &H24, gv0B68
  
  i = &H60
  Get hVBFile, Base1 + i, fBuff2: i = i + 2
  
  gv0B8A = Base1 + i
  i = i + Uint32(fBuff2)
  Get hVBFile, Base1 + i, fBuff1: i = i + 2
  Get hVBFile, Base1 + i, fBuff2: i = i + 2
  gv0B8E = Base1 + i
  gv09B6 = Uint32(fBuff2)
  gv09BA = (gv09B6 \ 2) - 1
  ReDim SomeDataBuff(0 To gv09BA), gv78AA(0 To gv09BA)
  ReDim gFormNames(0)
  ReDim gFrmNames2(0)
  StrTokens = String$(gv09BA, 0) & Chr$(&HE5)
  
  For fBuff1 = 0 To gv09BA
    Get hVBFile, , SomeDataBuff(fBuff1)
  Next
  
  i = i + gv09B6
  Get hVBFile, Base1 + i, fBuff2: i = i + 2
  
  i = i + Uint32(fBuff2)
  Get hVBFile, Base1 + i, fBuff2: i = i + 2
  gv0B92 = Base1 + i
  
  i = i + Uint32(fBuff2)
  gfrmOffset = i
  
  ClearLogUntil DispLvl
End Sub

Sub ScanModules()
Dim i As Long
Dim iBuff As Integer
Dim l017A As Integer
Dim iBuff2%
Dim FileName As String
Dim Disp_Lvl%

6090
' Scanning Local fixups
  DispStatus Disp_Lvl, Scanning & C_LOCAL & Fixups
  If gfrmOffset = 0 Then MakeErrBeep
  
  ResSeek ResID_RCData + 1
  If bPrjFilesCreated Then
    DispLog2 Disp_Lvl, Scanning & "Types"
    ScanTypes
    
    DispLog2 Disp_Lvl, Scanning & "Modules"
  End If
  
  i = gfrmOffset
  Do
    Get hVBFile, Base1 + i, iBuff2: i = i + 2
    If iBuff2 = 0 Or iBuff2 = -1 Then Exit Do
    
    l017A = RotateRight3(iBuff2)
    If NamesLookUpTbl(l017A) <> 1 Then MakeErrBeep
    
    iBuff2 = NamesLookUpTbl(l017A + 1)
    
    gFrmStruct(iBuff2).mfrmOffset = i
    FileName = PathAddExt("", Hex$(iBuff2))
    
   'Delete "Project.1", "Project.2"....

    While Len(Dir$(FileName, vbSystem Or vbHidden)): Kill FileName: Wend
    
    If bPrjFilesCreated Then ScanModuleToken1 iBuff2
    Get hVBFile, Base1 + i, iBuff: i = i + 2
    i = i + Uint32(iBuff)
    
    
    If gFrmStruct(iBuff2).M24D0 = &H46 Then '&H46 = 'F'
      Get hVBFile, Base1 + i, iBuff: i = i + 2
      i = i + iBuff
      
      Get hVBFile, Base1 + i, iBuff: i = i + 2
      i = i + iBuff
    Else
      Get hVBFile, Base1 + i, iBuff: i = i + 2
      i = i + iBuff
    End If
  Loop
  
  ClearLogUntil Disp_Lvl
  
End Sub

Sub ProcessControls()
6100
  curResID = 0
  ResSeek ResID_RCData
  Get hVBFile, CurResOffsetStart, FixedByteBuff
  Select Case VBver_11
  Case 1 To 3
    ScanRCFrmData
  Case Else
    ShowErrMsg mc0114
    VBver_11 = 0
  End Select
End Sub

Sub ScanTypes()
Dim l018E As Integer
Dim i

6110
  intTokens = -2 ' &HFFFE%
  
  If gv09B6 < 6 Then MakeErrBeep: Exit Sub
  
  l018E = SomeDataBuff(2)
  Do While l018E
    sub04AA l018E
  Loop
  
  For i = gv09BA - 2 To 0 Step -3
    If SomeDataBuff(i + 1) Or SomeDataBuff(i + 2) Then Exit For
    
    l018E = SomeDataBuff(i) Xor &H8000
    
    If l018E <= 0 Or l018E > gv0FD2 Then Exit For
    
    '' &HE0%
    sub04B5 i + 1, 2, 224, 9, 0
    SetFiltergFormNames gVBObjsDefs(l018E).Name_4, i + 1
  Next i
  
  ScanTypeSub1
  
End Sub

Sub sub0385(p019C As Integer)
Dim l019E As String
Dim l01A0 As String
Dim l01A2 As Integer
Dim l01A4 As Integer
Dim l01A8 As String * 2
Dim l01AA%
Dim l01AC%
Dim l01AE As Integer
Dim l01B0 As Integer
Dim EventName As String
Dim frmOffset As Integer
Dim mSize As Long
Dim Overlaps As Long
Dim size3 As Integer
Dim VBFrmName As String
Dim StringLen As Integer
Dim FrmName As String
Dim l01C6 As Long
Dim endPos As Long
 Const mc01DA = 1 ' &H1%
 Const mc01DC = 10 ' &HA%

6020
  Get hVBFile, , FrmStruct2
  'make pos
  mSize = FrmStruct2.mSize6 And &H80000000
  Overlaps = FrmStruct2.mSize6 Xor mSize
  Overlaps = Overlaps - Len(FrmStruct2) - 1
  endPos = Seek(hVBFile) + Overlaps + 1
  
  If mSize Then
    Get hVBFile, , l01A8
    Overlaps = Overlaps - 2
  End If
  
  Get hVBFile, , FixedByteBuff: size3 = Asc(FixedByteBuff)
  
  If size3 Then
    MakeErrBeep
    FrmName = Space$(size3)
    Get hVBFile, , FrmName
    Overlaps = Overlaps - size3
  
  Else
    l01A4 = Asc(FrmStruct2.M14E1)
    FrmName = gv0FF6(gFrmCount2 + l01A4).Name_4
    FrmStruct2.mSize6 = FrmStruct2.mSize6 + Len(FrmName)
  End If
  
  VBFrmName = Chr$(Len(FrmName)) & FrmName
  l01C6 = Seek(hVBFile)
  Get hVBFile, , FixedByteBuff
  StringLen = Asc(FixedByteBuff)
  
  If StringLen = &HFF Then
    
   'Read String
    Get hVBFile, , FixedByteBuff: l01A2 = Asc(FixedByteBuff)
    l01A0 = Space$(l01A2)
    Get hVBFile, , l01A0
    
    For StringLen = 1 To gKnowControlsCount
      
      If l01A0 = gKnowControlsData(StringLen).ControlTypeName Then Exit For
    Next StringLen
    
    If StringLen > gKnowControlsCount Then
      'Type like "SSFormFX" or "TextBox" not found
      std_p_e = l01A0
      ShowErrMsg Type_VBX
      StringLen = 0
    End If
  Else
    StringLen = StringLen + 1
    l01A0 = gKnowControlsData(StringLen).ControlTypeName
  End If
  
  If mSize Then VBFrmName = l01A8 & VBFrmName
  If p019C = 0 Then sub0394 gFrmCount2 + l01A4, StringLen, mSize <> 0
  
  Get hVBFile, l01C6 + Overlaps - 2, l01AC
  
  If l01AC <> &HFF Then
    l01AE = gKnowControlsData(StringLen).M2BE4 + 1
    Get hVBFile, l01C6 + Overlaps - 2 * l01AE - 2, l01AC
    If l01AC = (l01AE * &H100 Or &HFF) Then
      Overlaps = Overlaps - 2 * l01AE
    Else
      l01AE = &HFF
      Do
        Overlaps = Overlaps - 2
        l01AE = l01AE + &H100
        Get hVBFile, l01C6 + Overlaps - 2, l01AC
      Loop Until l01AC = l01AE
      l01AE = l01AC \ &H100
    End If
  End If
  
  FrmStruct2.mSize6 = FrmStruct2.mSize6 - 2 * l01AE - 1
  If p019C Then
    Put hFile_Mod, , FrmStruct2
    Put hFile_Mod, , VBFrmName
    Seek hVBFile, l01C6
    l019E = Space$(&H4000)
    While Overlaps > Len(l019E)
      Get hVBFile, , l019E: Put hFile_Mod, , l019E
      Overlaps = Overlaps - Len(l019E)
    Wend
    l019E = Space$(Overlaps - 1)
    Get hVBFile, , l019E: Put hFile_Mod, , l019E
    If Asc(Right$(l019E, 1)) <> &HFF Then MakeErrBeep
    Get hVBFile, , FixedByteBuff
    If Asc(FixedByteBuff) <> l01AE Then MakeErrBeep
    If l01AE Then Seek hVBFile, Seek(hVBFile) + 2 * l01AE
    If Seek(hVBFile) <> endPos Then MakeErrBeep
    
  ElseIf l01AE Then
    
    For l01AA = 0 To l01AE - 1
      Get hVBFile, , l01AC
      If l01AC Then
      
        If StringLen Then
          l01B0 = gKnowControlsData(StringLen).M2BDA + l01AA
          l01B0 = gv1510(l01B0)
          EventName = g_EventNames(l01B0).Name_4
        Else
          EventName = "event" & IntToHexStr(l01AA)
        End If
        
        If l01A4 Then
          EventName = FrmName & "_" & EventName
          
        ElseIf StringLen Then
          EventName = gKnowControlsData(StringLen).ControlTypeName & "_" & EventName
          
        Else
          EventName = "ctrl?_" & EventName
          
        End If
        
        If l01AC And 1 Then
          frmOffset = NamesLookUpTbl(RotateRight3(l01AC) + mc01DA)
          
          
          AddFunctionDeklaration frmOffset, EventName, mc01DC
          
          If StringLen Then sub03B4 frmOffset, l01B0, mSize <> 0
        Else
          MakeErrBeep
        End If
      
      End If
    Next l01AA
  
  End If
  
  Seek hVBFile, endPos

End Sub
'
'Public Function makeValidVBName(ByRef ControlName$) As Boolean
'Dim VALIDVBCHARS
'If Len(ControlName) = 40 Then exitsub
'Dim i
'
'
'End Function
