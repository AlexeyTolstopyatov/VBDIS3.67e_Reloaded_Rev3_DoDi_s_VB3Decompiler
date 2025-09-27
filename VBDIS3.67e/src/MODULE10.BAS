Attribute VB_Name = "MODULE10"
' Module10
Option Explicit
Const mc0062 = 1 ' &H1%
Const Type_GROUPCURSOR = 90 ' &H5A%
Const mc00AE = 86 ' &H56%
Const mc_Function = 7 ' &H7%
Const mc_Sub = 6 ' &H6%

Private Sub SetModuleNames()
Dim Moduls As Integer
Dim i As Integer
Dim ModNameString As String
Dim ModulBuff As T15C8

15000
  Moduls = ModulCount
  
  Do
    
    i = NamesLookUpTbl(Moduls + gc0E20)
    
    gFrmStruct(i).mSubroutinesNameLen = Moduls
    
    Get hVBFile, VBAbsOffs(Moduls), ModulBuff
    
    gFrmStruct(i).frmOffset = ModulBuff.frmOffset
    
    ModNameString = "Module" & Format$(i)
    
    SetFrmName i, ModNameString
    
    
    Moduls = NamesLookUpTbl(Moduls + gc0E22)
    
  Loop Until Moduls <= 0
  
End Sub

Private Sub Applyreloc1(Segm%)
Dim l0044%
Dim l0046 As Integer
Dim l0048 As Integer
Dim l004C As T1CBD
Dim l004E As Integer
Dim l0050%
Dim l0052 As Long
Dim l0054 As Integer
 Const mc0062 = 1 ' &H1%

15020
  ApplyReloc Segm
  
 'l0054 = (Segm - 3) * &H2000
 'l0054 = (Segm - VBStartSeg) * &H2000
  l0054 = 0
  
  
  For l0044 = 1 To gv0758
    l004C = gv0726(l0044)
    
    If l004C.iToken <> 0 Or l004C.M1CC7 <> 2 Then
      MakeErrBeep
    
    Else
      l0052 = Uint32(l004C.Size_M13F4)
      l0046 = l0054 + (l0052 \ 8) - 4
      
      If NamesLookUpTbl(l0046) <> gc0E1A Then MakeErrBeep
      l004E = NamesLookUpTbl(l0046 + gc0E26)
      l0048 = l004C.M1CB1
      
      If l004E = 0 Or NamesLookUpTbl(l004E) <> gc0E18 Then
        MakeErrBeep
      
      Else
        
        Do
          
          If NamesLookUpTbl(l0046) <> gc0E1A Then ShowErrMsg mc0062
          
          g_TKStruct(NamesLookUpTbl(l0046 + gc0E24)).segm0 = l0048
          Get hVBFile, curSegOffs + l0052, l0050
          
          l0052 = Uint32(l0050)
          l0046 = l0054 + (l0052 \ 8) - 4
          
        Loop Until l0050 = &HFFFF
      
      End If
    End If
  Next l0044
End Sub

Sub DoDecompilation()
Dim iSeg%
Dim l0070 As Long


15030
  curSegID = 0

  
 ' *** New *** ToDo Seek to Segment NE.Initial_CS points to
   
   
   For iSeg = 1 To Segments

    ' if Segment has RelocationFlag set and...
      If (Segs(iSeg).Flags And RELOCINFO) Then
       ' ... is not the Segment Entrypoint(CS:IP) points to
         If (iSeg <> NE.Initial_CS) Then Exit For
      End If

    Next
    
    SeekToSegment iSeg
    
    VBStartSeg = iSeg
  
    
   ScanGlobalData
  
   ReDim NamesLookUpTbl(curSegSize \ 8)

'*** old Code ***
'  l0070 = 0
'  SeekToSegment 3
'  If Segments <= 3 Then
'    Exit Sub
'  Else
'
'    For iSeg = 4 To 6
'      If (Segs(iSeg).Flags And RELOCINFO) = 0 Then Exit For
'      SeekToSegment iSeg
'    Next
'    l0070 = &H10000 * (iSeg - 4) + curSegSize '4224
'  End If
'  ScanGlobalData
'
'  ReDim NamesLookUpTbl(l0070 \ 8)
  ReDim gFormNames(0)
  FrmFilesCount = 0
  
  gSubroutines = 0
  curSegID = 0
  
  ScanSegmentData
  
  ReDim gFrmStruct(1 To FrmFilesCount)
  SetModuleNames
  sub0986
  ReDim g_TKStruct(gSubroutines)
  setPrjFilesCreatedTrue
  If bPrjFilesCreated Then
    sub0944
  End If
  
  ScanningFrms
     ProcessFrmRes
  ScanModules
End Sub


Sub ScanSegmentData()
Dim index As Integer
Dim l0090%
Dim l0094 As T15C8
Dim l0096%
 Const Type_GROUPCURSOR = 90 ' &H5A%

15040
  DispStatus l0096, Scanning & gc03E6
  Get hVBFile, CurResOffsetStart + Type_GROUPCURSOR, l0090 '136 '595c
  
  ModulCount = RotateRight3(l0090)
  Get hVBFile, VBAbsOffs(ModulCount), l0094
  
  l0090 = l0094.M15F5 '72
  ModulCount = RotateRight3(l0090)
  Do
    index = RotateRight3(l0090)
    Get hVBFile, VBAbsOffs(index), l0094
    
    FrmFilesCount = FrmFilesCount + 1
    
    NamesLookUpTbl(index) = gc0E18
    NamesLookUpTbl(index + gc0E20) = FrmFilesCount
    l0090 = l0094.M15F5
    
    If l0090 = &HFFFF Then Exit Do
    NamesLookUpTbl(index + gc0E22) = RotateRight3(l0090)
    
  Loop
  
  NamesLookUpTbl(index + gc0E22) = l0090 '-1 1810
  ClearLogUntil l0096
  
End Sub

Sub sub0986()
Dim l00A2 As Integer
Dim l00A4%
Dim l00A6 As Integer
Dim l00AC As VBModuleDef_struc
 Const mc00AE = 86 ' &H56%

15050
  Get hVBFile, CurResOffsetStart + mc00AE, l00A4
  gv0DE4 = RotateRight3(l00A4)
  Do While l00A4 <> &HFFFF
    l00A2 = RotateRight3(l00A4)
    Get hVBFile, VBAbsOffs(l00A2), l00AC
    gSubroutines = gSubroutines + 1
    NamesLookUpTbl(l00A2 + gc0E24) = gSubroutines
    NamesLookUpTbl(l00A2 + gc0E2A) = (l00AC.SubOrFunc \ &H100) And 7
    
   'Ext or local Function
    If l00AC.isExtOrLocal And 1 Then
      NamesLookUpTbl(l00A2) = gc0E1A
      NamesLookUpTbl(l00A2 + gc0E2C) = l00AC.Offs
    Else
      NamesLookUpTbl(l00A2) = gc0E1C
      NamesLookUpTbl(l00A2 + gc0E2C) = -1
    End If
    
    l00A6 = RotateRight3(l00AC.M_5_Size)
    If NamesLookUpTbl(l00A6) <> gc0E18 Then
      MakeErrBeep
    End If
    
    NamesLookUpTbl(l00A2 + gc0E26) = l00A6
    l00A4 = l00AC.M1551
    If l00A4 = &HFFFF Then Exit Do
    NamesLookUpTbl(l00A2 + gc0E28) = RotateRight3(l00A4)
  Loop
  NamesLookUpTbl(l00A2 + gc0E28) = l00A4
End Sub

Private Sub sub0944()
Dim l00B0 As Integer
Dim rol2 As Integer
Dim rol1 As Integer
Dim DisAsm As String
Dim i_frm%
Dim l00BA As Integer
Dim FuncType As Integer
Dim VBModuleDef As VBModuleDef_struc
 Const mc_Function = 7 ' &H7%
 Const mc_Sub = 6 ' &H6%

15010
  If gSubroutines = 0 Then Exit Sub
  rol1 = RotateLeft3(gv0DE4)
  
  Do While rol1 <> &HFFFF
    rol2 = RotateRight3(rol1)
    Get hVBFile, VBAbsOffs(rol2), VBModuleDef
    
    l00B0 = NamesLookUpTbl(rol2 + gc0E24)
    g_TKStruct(l00B0).TKFrmRes = rol2
    g_TKStruct(l00B0).frmOffset = VBModuleDef.frmOffset
    g_TKStruct(l00B0).mSubroutinesNameLen = NamesLookUpTbl(rol2 + gc0E26)
    
   'Sub or Function
    If VBModuleDef.SubOrFunc And 2 Then
      DisAsm = "fn": FuncType = mc_Function
      g_TKStruct(l00B0).M2515 = gv0EE2((VBModuleDef.SubOrFunc \ &H100) And 7)
      
    Else
      DisAsm = "sub": FuncType = mc_Sub
    End If
    
    If VBModuleDef.isExtOrLocal And 1 Then
      NamesLookUpTbl(rol2) = gc0E1A
      g_TKStruct(l00B0).Size_M13F4 = VBModuleDef.Offs
     
     'IsSubMain
      If rol1 = gSubMainIdx Then
        DisAsm = "Main": FuncType = 0
      Else
        DisAsm = DisAsm & Hex$(rol1)
      End If
      
    Else
      NamesLookUpTbl(rol2) = gc0E1C
      g_TKStruct(l00B0).Size_M13F4 = -1
      DisAsm = "ext" & DisAsm & Hex$(rol1)
      FuncType = FuncType + 2
    End If
    
    If FuncType Then DisAsm = ""
    
    AddFunctionDeklaration l00B0, DisAsm, FuncType
    rol1 = VBModuleDef.M1551
  Loop
  
'*** New***
Applyreloc1 VBStartSeg
'  For i_frm = 3 To 6
'    If (Segs(i_frm).Flags And RELOCINFO) = 0 Then Exit For
'    Applyreloc1 i_frm
'  Next
  
  For i_frm = 1 To FrmFilesCount
    rol2 = 0
    l00BA = gFrmStruct(i_frm).mSubroutinesNameLen
    
    For l00B0 = 1 To gSubroutines
      If g_TKStruct(l00B0).mSubroutinesNameLen = l00BA Then
        If g_TKStruct(l00B0).segm0 Then
          rol2 = rol2 + 1
          g_TKStruct(l00B0).M251F = rol2
        End If
      End If
    Next
  Next
End Sub
