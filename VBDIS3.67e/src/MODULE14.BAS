Attribute VB_Name = "MODULE14"
' Module14
Option Explicit
Const mc00E4 = 2048 ' &H800%
Const mc012E = 32 ' &H20%
Const mc0132 = 8192 ' &H2000%
Const mc0134 = 1024 ' &H400%
Const mc_TypeInteger = 8 ' &H8%
Const mc0150 = 9 ' &H9%
Const mc0152 = 10 ' &HA%
Const mc0154 = 11 ' &HB%
Const mc0156 = 12 ' &HC%
Const mc_TypeLong = 14 ' &HE%
Const mc015A = 10 ' &HA%
Const mc015C = 12 ' &HC%
Const mc015E = 512 ' &H200%
Const mc0160 = 256 ' &H100%
Const mc0162 = 8 ' &H8%
Const mc0164 = 9 ' &H9%
Const mc0166 = 13 ' &HD%
Const mc0168 = 11 ' &HB%
Const mc01F4 = 2 ' &H2%
Const mc01F6 = 1 ' &H1%
Const mc026A = 4 ' &H4%
Const mc026C = 8 ' &H8%
Const mc0284 = 2 ' &H2%
Const mc0288 = 1 ' &H1%
Const mc028E = 1 ' &H1%
Const mc0290 = 2 ' &H2%
Const mc02A2 = 3 ' &H3%
Const mc02A4 = 1 ' &H1%
Const mc02B6 = 2 ' &H2%

Sub sub0A71(p00DA As Integer)
Dim l00DC As Integer
 Const mc00E4 = 2048 ' &H800%

19040
  l00DC = SomeDataBuff(p00DA + 2)
  If (l00DC And &HFFF0) <> &HC100 Then
    ShowErrMsg mc00E4
    Exit Sub
  End If
  l00DC = l00DC And &HF
  If l00DC = 8 Then
  ElseIf l00DC = 0 Then
    l00DC = gc79CC
    If Not fn0CFD(SomeDataBuff(p00DA - 1)) Then
      ShowErrMsg mc00E4
      Exit Sub
    End If
  ElseIf l00DC <= 9 Then
  Else
    MakeErrBeep
    Exit Sub
  End If
  sub0CCF p00DA, gc79E2, l00DC, gc79D4
End Sub

Sub sub0AA2(p00E6 As Integer)
Dim l00E8 As Integer
Dim l00EA As Integer

19050
  l00E8 = gLocalVarsTypeArrExE(p00E6 + 2) And &HF
  If l00E8 = 0 Then l00E8 = gc79CC
  Select Case gLocalVarsTypeArrExE(p00E6 + 2) And &HF00
  Case 0: l00EA = gc79E2
    If gLocalVarsTypeArrExE(p00E6) = 0 Then MakeErrBeep
    sub0A71 UDiv(gLocalVarsTypeArrExE(p00E6), 2)
  Case &H100: l00EA = gc79DE
    If gLocalVarsTypeArrExE(p00E6) Then MakeErrBeep
  Case &H200: l00EA = gc79DE
    If gLocalVarsTypeArrExE(p00E6) Then MakeErrBeep
  Case &H500: l00EA = gc79DA
    If gLocalVarsTypeArrExE(p00E6) < 6 Then MakeErrBeep
  Case &H600: l00EA = gc_TypeInteger
    If gLocalVarsTypeArrExE(p00E6) >= 0 Then MakeErrBeep
  Case Else
    MakeErrBeep
  End Select
  ImproveTypeDiscovered p00E6, l00EA
  SetType2 p00E6, 0, l00EA, l00E8, gc79D4
End Sub

Sub sub0CCF(p00F4 As Integer, p00F6 As Integer, p00F8 As Integer, p00FA As Integer)
Dim l00FC As Integer
Dim l00FE As Integer
Dim l0100 As Integer

19060
  If p00F6 And &H1F Then
    MakeErrBeep
    p00F8 = p00F6 Or p00F8
    p00F6 = p00F6 And gc_0xE0
    p00F8 = p00F8 Xor p00F6
  End If
  If p00F6 = 0 Then
    MakeErrBeep
    Exit Sub
  End If
  l00FC = Asc(Mid$(StrTokens, p00F4, 1))
  If l00FC = 0 And SomeDataBuff(p00F4) Then
    p00F6 = gc79E6
  End If
  l00FE = p00F8 Or (p00FA And gc_0x10)
  If (p00F6 Or l00FE) = l00FC Then
    Exit Sub
  End If
  l0100 = l00FC And gc_0xE0
  l00FC = l00FC Xor l0100
  If p00F6 < l0100 Then p00F6 = l0100
  If l00FE = 0 Then
    l00FE = l00FC
  ElseIf l00FC And l00FC <> l00FE Then
    Exit Sub
  End If
  If (l00FE Or p00F6) = (l00FC Or l0100) Then
    Exit Sub
  End If
  sub04B5 p00F4, 0, p00F6, p00F8, p00FA
  If p00F8 And gv78AA(p00F4) = 0 Then
    sub0C63 p00F4
  End If
End Sub

Sub SetFiltergFormNames(ByRef FormName As String, p0104 As Integer)
Dim l0106 As Integer

19070
  If p0104 > gv09BA Then
    FormName = "???"
    Exit Sub
  End If
  
  l0106 = gv78AA(p0104)
  If Len(FormName) = 0 Then
    sub0C63 p0104
    FormName = fn0CDA(p0104)
    Exit Sub
  End If
    
    SetgFormNames FormName, gv78AA(p0104)

End Sub

Sub sub04B5(p010A As Integer, p010C As Integer, p010E As Integer, p0110 As Integer, p0112 As Integer)
Dim l0114 As Integer
Dim l0116 As Integer
Dim l0118%
Dim l011A%

19080
  If p010A < 3 Or p010A > gv09BA Then MakeErrBeep: Exit Sub
  l0114 = p010E Or p0110 Or (p0112 And gc_0x10)
  l0116 = Asc(Mid$(StrTokens, p010A, 1))
  If l0116 >= gc2_0xE0 Then
    If (l0114 Xor l0116) And &H1F Then
      p0110 = l0116 And &HF
      If l0116 And gc_0x10 Then p0112 = gc79D4 Else p0112 = 0
    End If
    Exit Sub
  End If
  
  If p010E = gc79E2 And p0112 = 0 Then
    Select Case p0110
    Case 1 To 6
      If p010C = 0 Then p010C = fn0D66(SomeDataBuff(), p010A, p0110, p0112)
      If p010C > 0 Then
        For l0116 = p010A To p010A + p010C - 1
          If SomeDataBuff(l0116) Then
            p010E = gc79E6
            Exit For
          End If
        Next
      End If
    End Select
  End If
  
  sub0C76 l0118, l011A, p010E%, p0110%, p010A%, True
  If p010E Then
    If p010C = 0 Then p010C = fn0D66(SomeDataBuff(), p010A, p0110, p0112)
    If p0110 And gc79C6 Then
      If intTokens > 0 Then
        If Asc(Mid$(StrTokens, p010A - 1, 1)) Then
          If Asc(Mid$(StrTokens, p010A - 1, 1)) <> p0110 Then MakeErrBeep
        End If
      End If
      Mid$(StrTokens, p010A - 1, 1) = Chr$(p0110)
    End If
    
    Mid$(StrTokens, p010A, 1) = Chr$(p010E Or p0110 Or (p0112 And gc_0x10))
    If p010E > gc_0x20 And p0110 Then
      If gv78AA(p010A) = 0 Then
        sub0C63 p010A
      End If
    End If
    
  Else
    Mid$(StrTokens, p010A, 1) = Chr$(p0110)
  End If
  
  If p010C > 1 Then
    Mid$(StrTokens, p010A + 1, p010C - 1) = String$(p010C - 1, gc79D2)
    If p0112 Then Mid$(StrTokens, p010A + 1, 1) = Chr$(p0112)
  End If
  
  bPrjFilesCreated = True
  
End Sub

Sub SetType1(VarIndex As Integer, pTypeDiscovered1 As Integer, pTypeDiscovered2 As Integer, pTypeDiscovered3%)
Dim curType As Integer
Dim pTypeDiscovered2_3 As Integer
Dim curType_Masked As Integer
 Const mc012E = 32 ' &H20%
 Const mc0132 = 8192 ' &H2000%
 Const mc0134 = 1024 ' &H400%

19090
  If (pTypeDiscovered1 And &H1F) Then
    ShowErrMsg mc012E
    pTypeDiscovered2 = (pTypeDiscovered2 Or pTypeDiscovered1) And &HF
    pTypeDiscovered1 = pTypeDiscovered1 And gc_0xE0
  End If
  
  If ScanOnly Then
    Select Case pTypeDiscovered1
    
    Case gv79C4, gc2_0xE0
      ShowErrMsg mc0132
      Exit Sub
    
    Case gc_0x20
    
    Case gc79DE
      If gLocalVarsTypeArrExE(VarIndex) Then
        pTypeDiscovered1 = gc79E6
      End If
    
    Case gc_TypeInteger, gc79DA, gc79E6
      If gLocalVarsTypeArrExE(VarIndex) = 0 Then
        ShowErrMsg mc0134
        pTypeDiscovered1 = gc79DE
      End If
    
    Case gc79E4
      If Uint32(gLocalVarsTypeArrExE(VarIndex)) > gv09B6 Or (gLocalVarsTypeArrExE(VarIndex) And 1) Then
      Else
        SetType2 VarIndex, 0, gc79E4, pTypeDiscovered2, pTypeDiscovered3
      End If
      Exit Sub
    
    Case Else
      ShowErrMsg mc012E
      Exit Sub
    End Select
  
  End If
  
  curType = Asc(Mid$(gLocalVarsTypeBuff, VarIndex, 1))
  pTypeDiscovered2_3 = pTypeDiscovered2 Or (pTypeDiscovered3 And gc_0x10)
  If (pTypeDiscovered1 Or pTypeDiscovered2_3) = curType Then
    Exit Sub
  End If
  
  curType_Masked = curType And gc_0xE0
  curType = curType Xor curType_Masked
  
  If pTypeDiscovered1 < curType_Masked Then pTypeDiscovered1 = curType_Masked
  If pTypeDiscovered1 = gc79E2 Then
    
    curType = Asc(Mid$(StrTokens, UDiv(gLocalVarsTypeArrExE(VarIndex), 2), 1))
    curType_Masked = curType And gc_0xE0
    curType = curType Xor curType_Masked
    
    If pTypeDiscovered1 < curType_Masked Then pTypeDiscovered1 = curType_Masked
  End If
  
  If pTypeDiscovered2_3 = 0 Then
    pTypeDiscovered2_3 = curType
  ElseIf curType Then
    If pTypeDiscovered2_3 <> curType Then
      ShowErrMsg mc00E4
    End If
    Exit Sub
  End If
  
  If (pTypeDiscovered2_3 Or pTypeDiscovered1) = (curType Or curType_Masked) Then
    Exit Sub
  End If
  
  SetType2 VarIndex, 0, pTypeDiscovered1, pTypeDiscovered2_3, pTypeDiscovered3

End Sub

Sub ProcessType(VarOffset%, pToken1 As Integer, pToken2 As Integer)
Dim CurVarType As Integer
Dim CurVarType_0xE0_Masked As Integer
Dim VarIndex As Integer
Dim Type_Discovered As Integer
Dim Type_Discovered2 As Integer
Dim l014A%
 Const mc_TypeInteger = 8 ' &H8%
 Const mc0150 = 9 ' &H9%
 Const mc0152 = 10 ' &HA%
 Const mc0154 = 11 ' &HB%
 Const mc0156 = 12 ' &HC%
 Const mc_TypeLong = 14 ' &HE%
 Const mc015A = 10 ' &HA%
 Const mc015C = 12 ' &HC%
 Const mc015E = 512 ' &H200%
 Const mc0160 = 256 ' &H100%
 Const mc0162 = 8 ' &H8%
 Const mc0164 = 9 ' &H9%
 Const mc0166 = 13 ' &HD%
 Const mc0168 = 11 ' &HB%

19100
  If pToken1 < 8 Then Exit Sub
  
 'Get Current VarType
  VarIndex = UDiv(VarOffset, 2)
  If VarIndex > gLocalVarsCount Then Exit Sub
  CurVarType = Asc(Mid$(gLocalVarsTypeBuff, VarIndex, 1))
  
  CurVarType_0xE0_Masked = CurVarType And gc_0xE0
  CurVarType = CurVarType Xor CurVarType_0xE0_Masked
    
    If gLocalVarsTypeArr(VarIndex) Then
      ImproveTypeDiscovered VarIndex, CurVarType_0xE0_Masked
      Exit Sub
    End If
    
  If CurVarType And gc_0x10 Then Exit Sub
  
  If CurVarType > gc2_0xE0 Then Exit Sub
  
 '0..32 Single Types
  If CurVarType_0xE0_Masked <= gc_0x20 Then
    Select Case pToken1
   
   'integer
    Case mc_TypeInteger
      Type_Discovered = gc_TypeInteger
      
    Case mc0150
      Type_Discovered = gc79DE
      
    Case mc0152
      Type_Discovered = gc79E4
      
    Case mc0154
      Type_Discovered = gc79DA
      
    Case mc0156
      If CurVarType_0xE0_Masked > gc_0x20 Then Exit Sub
      Type_Discovered = gc_0x20
      
    Case mc_TypeLong
      
      If pToken2 = mc015A Then
        sub0AA2 VarIndex
        Exit Sub
      ElseIf pToken2 = mc015C Then
        Type_Discovered = gc2_0xE0
      End If
      
      If gLocalVarsTypeArrExE(VarIndex) = 0 Then
        ShowErrMsg mc015E
        Type_Discovered = gc79DE
      ElseIf gLocalVarsTypeArrExE(VarIndex) And &H8001 Then
        Type_Discovered = gc_TypeInteger
      Else
        Type_Discovered = gc79DA
      End If
      
    Case Else
      MakeErrBeep
      If CurVarType_0xE0_Masked Then Exit Sub
      Type_Discovered = gc_0x20
    End Select
    
    ImproveTypeDiscovered VarIndex, Type_Discovered
  Else
    Type_Discovered = CurVarType_0xE0_Masked
  End If
  
  Select Case pToken2
  Case 0
    ShowErrMsg mc0160
    
  Case mc0162
    Type_Discovered2 = 0
    
  Case mc0164
    Type_Discovered2 = gc79CA
    
  Case mc015A
    sub0AA2 VarIndex
    Exit Sub
    
  Case mc0166: Type_Discovered2 = gTypeFixString
  
  Case Is < 8
    Type_Discovered2 = pToken2
  
  Case mc0168
    Type_Discovered2 = gc79CC
  
  Case mc015C
    SetType2 VarIndex, 1, gc2_0xE0, gc79EE, 0
    Exit Sub
  
  Case Else
    MakeErrBeep
    Exit Sub
  End Select
  
  SetType1 VarIndex, Type_Discovered, Type_Discovered2, l014A
  
End Sub

Sub SetType2(VarIndex%, pZero As Integer, pTypeDiscovered1 As Integer, pTypeDiscovered2_3 As Integer, pTypeDiscovered3 As Integer)
Dim i%
Dim l0176%
Dim l0178%

19110
  If pTypeDiscovered2_3 < 0 Then Exit Sub
  
  If pTypeDiscovered1 = gc79DE And pTypeDiscovered3 = 0 Then
    Select Case pTypeDiscovered2_3
    Case 1 To 6
      If pZero = 0 Then pZero = fn0D15(VarIndex, pTypeDiscovered1 Or pTypeDiscovered2_3, pTypeDiscovered3)
      If pZero > 0 Then
        
        For i = VarIndex To VarIndex + pZero - 1
          
          If gLocalVarsTypeArrExE(i) Then
            pTypeDiscovered1 = gc79E6
            Exit For
          End If
        
        Next
        
      End If
    End Select
    
  End If
  
  sub0C76 l0176, l0178, pTypeDiscovered1, pTypeDiscovered2_3, VarIndex, False
  If pTypeDiscovered1 Then
    If pZero = 0 Then pZero = fn0D15(VarIndex, pTypeDiscovered1 Or pTypeDiscovered2_3, pTypeDiscovered3)
    
    If pTypeDiscovered2_3 And gc79C6 Then
      If Asc(Mid$(gLocalVarsTypeBuff, VarIndex - 1, 1)) Then
        If Asc(Mid$(gLocalVarsTypeBuff, VarIndex - 1, 1)) <> pTypeDiscovered2_3 Then
          If Asc(Mid$(gLocalVarsTypeBuff, VarIndex - 1, 1)) <> gc79D2 Then
            If intTokens > 0 Then
              MakeErrBeep
              Exit Sub
            End If
          End If
        End If
      End If
      Mid$(gLocalVarsTypeBuff, VarIndex - 1, 1) = Chr$(pTypeDiscovered2_3)
    End If
    If pTypeDiscovered1 = gc79E4 Then sub04B5 UDiv(gLocalVarsTypeArrExE(VarIndex), 2), 0, gc79E2, pTypeDiscovered2_3, pTypeDiscovered3
    Mid$(gLocalVarsTypeBuff, VarIndex, 1) = Chr$(pTypeDiscovered1 Or pTypeDiscovered2_3 Or (pTypeDiscovered3 And gc_0x10))
  Else
    Mid$(gLocalVarsTypeBuff, VarIndex, 1) = Chr$(pTypeDiscovered2_3)
  End If
  
  gLocalVarsTypeArr(VarIndex) = fn0D0A(VarIndex)
  
  If pZero > 1 Then
    Mid$(gLocalVarsTypeBuff, VarIndex + 1, pZero - 1) = String$(pZero - 1, gc79D2)
    If pTypeDiscovered3 Then
      Mid$(gLocalVarsTypeBuff, VarIndex + 1, 1) = Chr$(pTypeDiscovered3)
    End If
  End If
  
  gWriteOutPut = True
End Sub

Sub SetFrmName(idx As Integer, FrmName As String)
Dim Offset As Integer

19120
  Offset = gFrmStruct(idx).frmOffset
  SetgFormNamesByName Offset, FrmName
End Sub

Sub SetgFormNamesByName(Offset As Integer, ByVal FrmName As String)
Dim Offset4 As Integer
Dim l0190 As Integer

19130
  Offset4 = UDiv(Offset, 4)
  If Offset4 > UBound(gFrmNames2) Then ReDim Preserve gFrmNames2(Offset4)
  
  l0190 = gFrmNames2(Offset4)
  If l0190 <= 0 Then
    SetgFormNames FrmName, gFrmNames2(Offset4)
  
  ElseIf gFormNames(l0190) <> FrmName Then
    bPrjFilesCreated = True
    gFormNames(l0190) = FrmName
  
  End If

End Sub

Sub AddFunctionDeklaration(ModuleNum As Integer, ModuleName As String, p0196 As Integer)
Dim frmOffset As Integer
Dim frmIndex As Integer

19140
  frmOffset = g_TKStruct(ModuleNum).frmOffset
  
'Enlage Arr if needed
 frmIndex = UDiv(frmOffset, 4)
 If frmIndex > UBound(gFrmNames2) Then ReDim Preserve gFrmNames2(frmIndex)
  
  
  If Len(ModuleName) Then
    If ModuleName = getFrmNamebyOffbyIdx(ModuleNum) Then Exit Sub
    SetgFormNamesByName frmOffset, ModuleName
  
  Else
    
    gFrmNames2(frmIndex) = -p0196
    ModuleName = getFrmNamebyOff(frmOffset)
    bPrjFilesCreated = True
  
  End If
  
End Sub

Private Sub sub0C76(p01A4 As Integer, p01A6 As Integer, p01A8 As Integer, p01AA As Integer, p01AC As Integer, p01AE As Integer)
Dim l01B0 As Integer
Dim l01B2 As Integer

19150
  If intTokens = 0 Then
    If gv0BC6 Then MakeErrBeep
    Exit Sub
  ElseIf intTokens = gc79BC Then
    If p01AE Then p01A4 = intTokens
    p01A6 = intTokens
    Exit Sub
  End If
  
  If p01AE Then
    l01B0 = Asc(Mid$(StrTokens, p01AC, 1))
  Else
    l01B0 = Asc(Mid$(gLocalVarsTypeBuff, p01AC, 1))
  End If
  
  l01B2 = l01B0 And &H1F
  l01B0 = l01B0 Xor l01B2
  l01B2 = l01B2 And &HF
  
  Select Case l01B0
  Case gv79C4
    p01A4 = intTokens
  Case gc_0x20
    If p01A8 > gc_0x20 Then p01A4 = intTokens
  Case gc79E6
    If p01A8 <> l01B0 Then
      If p01A4 <> gc79BE And p01A8 <> gc79DE And p01A8 <> gc79E2 Then
        GoSub L17CF6
      End If
      p01A8 = l01B0
    End If
  Case gc2_0xE0
      p01A8 = l01B0
      Exit Sub
    If p01A8 <> l01B0 Then
      GoSub L17CF6
      p01A8 = l01B0
      Exit Sub
    End If
  Case Else
    If p01A8 <> l01B0 Then
      If p01A8 <> gc79E6 Then
        GoSub L17CF6
        p01A8 = l01B0
      End If
    End If
  End Select
  
  If p01AA Then
    If l01B2 = 0 Then
      p01A6 = intTokens
    ElseIf p01AA <> l01B2 Then
      GoSub L17CF6
    End If
  Else
    p01AA = l01B2
  End If
Exit Sub

L17CF6:
  ShowErrMsg mc00E4
Return
End Sub

Sub sub04AA(p01B6%)
Dim l01B8 As Integer
Dim l01BA As Integer
Dim l01BC As Integer
Dim l01BE As Integer
Dim l01C0 As Integer
Dim l01C2 As Integer
Dim l01C4 As Integer
Dim l01C8 As T1379
Dim l01CC As T13D2
Dim l01D0 As T20E6

19160
  l01BA = UDiv(p01B6, 2)
  If Asc(Mid$(StrTokens, l01BA, 1)) Then
    p01B6 = 0
    Exit Sub
  End If
  Get hVBFile, gv0B8E + Uint32(p01B6), l01C8
  sub04B5 l01BA, 4, gc7A0C, 0, 0
  sub0C96 gc78E8, l01BA
  l01B8 = l01C8.M13A7
  If l01C8.M1390 And 1 Then
    l01BC = UDiv(l01B8, 2)
    l01B8 = SomeDataBuff(l01BC - 1)
  End If
  Do While l01B8
    l01BC = UDiv(l01B8, 2)
    Get hVBFile, gv0B8E + Uint32(l01B8), l01CC
    l01C4 = l01BC + 4
    l01BE = l01CC.Type_3
    If l01BE > 9 Or l01BE < 0 Then
      l01BE = gc7A0E Or gc79CE
    Else
      l01BE = gc7A0E Or l01CC.Type_3
    End If
    If l01CC.M13DF And 1 Then
      l01BE = l01BE Or gc_0x10
      l01C0 = 6 + (SomeDataBuff(l01BC + 6) And &HFF) * 2
      sub04B5 l01C4, l01C0, 0, gc79D4, 0
      l01C4 = l01C4 + l01C0
    End If
    If Asc(Mid$(StrTokens, l01BC - 1, 1)) = 0 Then
      l01BE = l01BE Or gc79C6
      If l01CC.Type_3 = 8 Then
        l01C2 = gTypeFixString
      ElseIf l01CC.Type_3 = 0 Then
        MakeErrBeep
      Else
        MakeErrBeep
      End If
    Else
      l01C2 = 0
    End If
    sub04B5 l01BC, 4, l01BE, l01C2, 0
    sub0C96 gc78EA, l01BC
    l01B8 = l01CC.M13DF And -2
  Loop
  If l01C8.M1390 And 1 Then
    l01B8 = l01C8.M13A7
    l01BC = UDiv(l01B8, 2) - 1
    If l01BC <> l01C4 Then MakeErrBeep
    Get hVBFile, gv0B8E + Uint32(l01B8), l01D0
    l01C0 = (6 + l01D0.M1EFD * 3) \ 2
    sub04B5 l01BC, l01C0, 0, gc79D2, 0
    l01C4 = l01BC + l01C0
  End If
  Mid$(StrTokens, l01C4 - 1, 1) = Chr$(gc7A10)
  p01B6 = l01C8.M1390 And -2
End Sub

Sub sub0C63(p01D6 As Integer)
Dim l01D8 As Integer
19170
  l01D8 = gv78AA(p01D6)
  If l01D8 Then Exit Sub
  gv78AA(p01D6) = fn0CE6(p01D6)
End Sub

Sub sub0C96(p01DA As Integer, p01DC As Integer)
Dim l01DE As Integer
19180
  l01DE = gv78AA(p01DC)
  If l01DE > 0 Then
    Exit Sub
  ElseIf l01DE < 0 Then
    If l01DE = -p01DA Then Exit Sub
  End If
  gv78AA(p01DC) = -p01DA
  bPrjFilesCreated = True
End Sub

Sub sub0CAC(p01E0 As Integer, p01E2 As Integer)
19190
  gLocalVarsTypeArr(p01E0) = p01E2
End Sub

Sub sub0A90(p01E4 As String, p01E6 As Integer)
Dim l01E8 As Integer
Dim l01EA As Integer
 Const mc01F4 = 2 ' &H2%
 Const mc01F6 = 1 ' &H1%

19200
  l01EA = Asc(Mid$(gLocalVarsTypeBuff, p01E6, 1))
  Select Case l01EA And gc_0xE0
  Case gv79C4, gc2_0xE0
    Select Case l01EA
    Case gc7A1C
      l01E8 = RotateRight3(gLocalVarsTypeArrExE(p01E6))
      If NamesLookUpTbl(l01E8) = mc01F4 Then
        If p01E4 = getFrmNamebyOffbyIdx(NamesLookUpTbl(l01E8 + mc01F6)) Then Exit Sub
        Stop
        AddFunctionDeklaration NamesLookUpTbl(l01E8 + mc01F6), p01E4, 0
        Exit Sub
      End If
    Case gc7A16
      If intTokens <> gc79BE Then
        MakeErrBeep
      Else
        GoTo L183C2
      End If
    Case Else
      MakeErrBeep
    End Select
  Case gc79E4
    If gLocalVarsTypeArrExE(p01E6) Then
      SetFiltergFormNames p01E4, UDiv(gLocalVarsTypeArrExE(p01E6), 2)
    Else
      MakeErrBeep
    End If
  Case gc79E6, gc79DE, gc_TypeInteger, gc79DA
L183C2:
    l01EA = UBound(gVBCodeTokenAddData) + 1
    ReDim Preserve gVBCodeTokenAddData(l01EA)
    gVBCodeTokenAddData(l01EA) = p01E4
    gLocalVarsTypeArr(p01E6) = l01EA
  Case Else
    MakeErrBeep
  End Select
End Sub

Function fn0CBF() As String
19210
End Function

Sub sub0C3D(p01FC As Integer, p0200 As ComboBox, p0204 As ComboBox, p0208 As ComboBox)
19020
  p0200.AddItem "<special>"
  p0200.AddItem "<refd>"
  If p01FC Then
    p0200.AddItem "?"
    p0200.AddItem "?"
    p0200.AddItem "Member"
  Else
    p0200.AddItem "Param"
    p0200.AddItem "Modul"
    p0200.AddItem "Local"
  End If
  p0200.AddItem "Global"
  p0200.AddItem "Const"
  p0200.AddItem "<fix>"
  p0204.AddItem ""
  p0204.AddItem "%"
  p0204.AddItem "&"
  p0204.AddItem "!"
  p0204.AddItem "#"
  p0204.AddItem "@"
  p0204.AddItem "Variant"
  p0204.AddItem "$"
  p0204.AddItem "-$*"
  p0204.AddItem "-Object"
  p0204.AddItem "-Type"
  p0204.AddItem "+MType"
  p0208.AddItem ""
  p0208.AddItem "+"
  p0208.AddItem "+()"
End Sub

Function fn0C4A(p020C As Integer, p020E As Integer) As String
Dim l0210 As Integer

19030
  l0210 = gFrmStruct(p020C).M_2_curr
  l0210 = gVBObjsDefs(l0210).mFrmCount
  If l0210 = 0 Then MakeErrBeep
  fn0C4A = gv0FF6(l0210 + p020E).Name_4
End Function

Function fn0CDA(p0224 As Integer) As String
Dim l0226 As Integer
Dim l0228 As String

19220
  l0226 = fn0CE6(p0224)
  If l0226 > 0 Then
    fn0CDA = gFormNames(l0226)
    Exit Function
  Else
    l0228 = fnPreFix(-l0226)
  End If
  If Left$(l0228, 1) = "g" Then
    fn0CDA = l0228 & IntToHexStr(Mul_UInt(p0224, 2))
  Else
    fn0CDA = l0228 & IntToHexStr(SomeDataBuff(p0224))
  End If
End Function

Function fn0CE6(p0230 As Integer) As Integer
Dim l0232 As Integer

19230
  If p0230 <= 0 Or p0230 > gv09BA Then
    Exit Function
  End If
  l0232 = gv78AA(p0230)
  If l0232 Then fn0CE6 = l0232: Exit Function
  l0232 = Asc(Mid$(StrTokens, p0230, 1))
  If l0232 = 0 Or l0232 = gc7A1A Then
    If SomeDataBuff(p0230) Then
      l0232 = gc79E6
    Else
      l0232 = gc79E2
    End If
  End If
  Select Case l0232 And gc_0xE0
  Case gc79E2
    fn0CE6 = -gc78E4
  Case gc79E6, gc2_0xE0
    fn0CE6 = -gc78E6
  Case Else
  End Select
End Function

Function fn0CF1(p0236 As Integer) As Integer
Dim l0238 As Integer
Dim l023A As Integer

19240
On Error GoTo L18CF0
  l023A = UDiv(gLocalVarsTypeArrExE(p0236), 2)
  If l023A = 0 Or l023A > gv09BA Then Exit Function
  l0238 = Asc(Mid$(gLocalVarsTypeBuff, p0236, 1))
  Select Case l0238 And gc_0xE0
  Case gc79E4
19241
    fn0CF1 = l023A
  Case gv79C4, gc_0x20
19242
    l0238 = Asc(Mid$(StrTokens, l023A, 1))
    If l0238 = 0 Then
19243
      If Asc(Mid$(StrTokens, l023A - 1, 1)) = 0 Then
        Exit Function
      End If
      fn0CF1 = l023A
    Else
19244
      l0238 = l0238 And gc_0xE0
      If l0238 = gc79E6 Or l0238 = gc79E2 Then
        fn0CF1 = l023A
      End If
    End If
  End Select
L18CF0:
19249
Exit Function
End Function

Function fn0CFD(ByVal pv023E%) As Integer
19250
  pv023E = UDiv(pv023E, 2)
  If pv023E < gv09BA Then fn0CFD = Asc(Mid$(StrTokens, pv023E, 1)) = gc7A0C
End Function

Function fn0D0A(VarIndex As Integer) As Integer
Dim l0244 As Integer
Dim curType As Integer
Dim VarOffset As Integer
Dim l024A As Integer
Dim curArrType As Integer

19260
  curArrType = gLocalVarsTypeArr(VarIndex)
  If curArrType Then fn0D0A = curArrType: Exit Function
  
  VarOffset = Mul_UInt(VarIndex, 2)
  l024A = gLocalVarsTypeArrExE(VarIndex)
  
  curType = Asc(Mid$(gLocalVarsTypeBuff, VarIndex, 1))
  Select Case curType
  Case 0, gc7A1A
    fn0D0A = -gc78F8
    Exit Function
  
  Case gc7A1C
    fn0D0A = -gc78F6
    l0244 = RotateRight3(l024A)
    
    If l0244 > UBound(NamesLookUpTbl) Then
      Exit Function
    End If
    
    If NamesLookUpTbl(l0244) < mc01F4 Then
      Exit Function
    End If
    
    fn0D0A = -gc78FE
    Exit Function
  
  Case gc7A16
    fn0D0A = -gc7908
    Exit Function
  
  Case gc7A18
    fn0D0A = -gc78FC
    Exit Function
  
  End Select
  
  Select Case curType And gc_0xE0
  Case gc79DE
    If l024A Then MakeErrBeep
    fn0D0A = -gc78FA
  
  Case gc79E4
    If l024A = 0 Then MakeErrBeep
    fn0D0A = -gc78FC
  
  Case gc79DA
    fn0D0A = -gc7902
  
  Case gc_TypeInteger
    If l024A <= 0 Or (l024A And 1) > 0 Then
      fn0D0A = -gc7900
    Else
      fn0D0A = -gc7904
    End If
  Case gc79E6
    fn0D0A = -gc7906
  Case gc_0x20
    fn0D0A = -gc78F8
  Case Else
    fn0D0A = -gc78F6
  End Select
End Function

Function fn0D15(p0250%, p0252 As Integer, p0254 As Integer) As Integer
19270
  Select Case p0254
  Case 0
  Case gc79D2
    MakeErrBeep
  Case gc79D4
    fn0D15 = fn0C30(gLocalVarsTypeArrExE(), p0250)
    Exit Function
  Case Else
    MakeErrBeep
    fn0D15 = 2
    Exit Function
  End Select
  Select Case p0252 And gc_0xE0
  Case gv79C4: MakeErrBeep
  Case gc_0x20
  Case gc79E4
    fn0D15 = 1
  Case gc79DE, gc79E6
    fn0D15 = fn0D66(gLocalVarsTypeArrExE(), p0250, p0252 And &H1F, p0254)
  Case gc79DA
    Select Case p0252 And &HF
    Case 6
      fn0D15 = 2
    Case Else
      fn0D15 = 1
    End Select
  Case gc_TypeInteger
    Select Case p0252 And &HF
    Case 6
      fn0D15 = 2
    Case 0 To gc79CA
      fn0D15 = 1
    Case gc79CC
      fn0D15 = 1
    Case Else
      MakeErrBeep
      fn0D15 = 1
    End Select
  Case Else
    MakeErrBeep
  End Select
End Function

Function ModulFileNames(p0258 As Integer) As String
On Error Resume Next
19280
  ModulFileNames = getFrmNamebyOff(gFrmStruct(p0258).frmOffset)
End Function

Function getFrmNamebyOff(FrmNameOff%) As String
Dim FrmNameIdx As Integer
Dim VarType As Integer
19290
  FrmNameIdx = UDiv(FrmNameOff, 4)
  VarType = gFrmNames2(FrmNameIdx)
  
  If VarType = 0 Then
    VarType = -1
    gFrmNames2(FrmNameIdx) = VarType
  End If
  
  
  If VarType > 0 Then
   'Get Name
    getFrmNamebyOff = gFormNames(VarType)
    
  ElseIf VarType < 0 Then
   'Generate a Name
    getFrmNamebyOff = fnPreFix(-VarType) & IntToHexStr(FrmNameOff)
  Else
    gFrmNames2(FrmNameIdx) = -1
  End If

End Function

Sub ScanExe()
Dim l0262%
 Const mc026A = 4 ' &H4%
 Const mc026C = 8 ' &H8%

19300
  If Segments <= 3 Then
    ShowErrMsg "Program without code - abort"
    gv7B32 = gv7B32 Or mc026A
    UnloadReason = mc026C
    Exit Sub
  End If
  
  ProcessControls
  gv0B9E = ""
  
  If VBver_11 = 3 Then
    SeekToRCData &H8003
    
    If curResID Then
      Get hVBFile, , l0262
      gv0B9E = Space$(l0262)
      Get hVBFile, , gv0B9E
    
    End If
  ElseIf VBver_11 <> 2 Then
    ShowErrMsg mc026A
    Exit Sub
  End If
  
  ScanOnly = True
  
  DoDecompilation
  Save1File
  ScanOnly = False
  NewFileExt = 0
End Sub

Sub sub07AD(p0278() As Integer, p027E As Integer)
Dim l0280 As Integer
Dim i As String
 Const mc0284 = 2 ' &H2%
Static m0286 As Integer
 Const mc0288 = 1 ' &H1%

19310
  Select Case p0278(p027E - 1) And &HFFC0
  Case &HC000, &HC500, &HC600
    l0280 = 0
  Case &HC100, &HC200
    l0280 = p0278(p027E - 2)
    If l0280 And &H4000 Then
      l0280 = l0280 And &HFF
    Else
      l0280 = 0
    End If
  Case &H4600
    l0280 = 0
  Case Else
    l0280 = p0278(p027E + mc0284)
    If (l0280 And &H7FC0) = &H5000 Then
      l0280 = l0280 And &HFF
    Else
      ShowErrMsg mc012E
      Exit Sub
    End If
  End Select
  i = ")"
  p027E = p027E + 6
  While l0280 > 0
    i = Format$(p0278(p027E + m0286) + p0278(p027E + mc0288) - 1) & i
    If p0278(p027E + mc0288) Then
      i = Format$(p0278(p027E + mc0288)) & " To " & i
    End If
    l0280 = l0280 - 1
    If l0280 Then i = ", " & i
    p027E = p027E + 2
  Wend
  DisOutput "(" & i
End Sub

Sub sub0D47(p028A As Long)
Dim l028C As Integer
 Const mc028E = 1 ' &H1%
 Const mc0290 = 2 ' &H2%

19320
  l028C = p028A \ 2
  DisOutput "  " & fn0CDA(l028C)
  If SomeDataBuff(l028C + mc028E) And 1 Then
    sub07AD SomeDataBuff(), l028C + 4
  End If
  If SomeDataBuff(l028C + mc0290) >= gc79CA Or SomeDataBuff(l028C + mc0290) < 0 Then
    DetermDeclareType gc79CC, SomeDataBuff(l028C + mc0290), False
  Else
    DetermDeclareType SomeDataBuff(l028C + mc0290), SomeDataBuff(l028C - 1), False
  End If
  DisOutputWrite
  p028A = SomeDataBuff(l028C + mc028E) And &HFFFE&
End Sub

Sub DetermDeclareType(TokenByte As Integer, pVBCodeTokenWordsExE%, pUseNew As Integer)
19330
  Select Case TokenByte And &HF
  Case 0
  
  Case gc79CC
    DisOutput " As " & fn0CDA(UDiv(pVBCodeTokenWordsExE, 2))
  
  Case gTypeFixString
    DisOutput " As String *" & Str$(Uint32(pVBCodeTokenWordsExE))
    
  Case Is < 8
    DisOutput " As " & gVB_a_DataTypes(TokenByte And &HF)
  
  Case gc79CA:
    If gVBObjsDefs(pVBCodeTokenWordsExE And &HFF).id_5 < 0 Then DisOutput "()"
    DisOutput " As "
    If pUseNew And pVBCodeTokenWordsExE And &H8000 Then DisOutput "New "
    DisOutput gVBObjsDefs(pVBCodeTokenWordsExE And &HFF).Name_4
  
  Case Else
    If fn0CFD(TokenByte) Then
      DisOutput " As " & fn0CDA(UDiv(TokenByte, 2))
    Else
      MakeErrBeep
      DisOutput " As " & Hex$(TokenByte)
    End If
    
  End Select
  
DisOutput " ' " & Hex$(TokenByte)

End Sub

Sub sub0D56(p029E%)
Dim l02A0 As Long
 Const mc02A2 = 3 ' &H3%
 Const mc02A4 = 1 ' &H1%

19340
  DisOutput "Type " & fn0CDA(p029E): DisOutputWrite
  l02A0 = Uint32(SomeDataBuff(p029E + mc02A2))
  If SomeDataBuff(p029E + mc02A4) And 1 Then
    l02A0 = Uint32(SomeDataBuff((l02A0& \ 2) - 1))
  End If
  While l02A0
    sub0D47 l02A0
  Wend
  DisOutput "End Type": DisOutputWrite: DisOutputWrite
End Sub

Function fn0D66(p02A8() As Integer, p02AE%, p02B0 As Integer, p02B2 As Integer) As Integer
Dim l02B4 As Integer
 Const mc02B6 = 2 ' &H2%

19350
  Select Case p02B2
  Case gc79D4
    fn0D66 = fn0C30(p02A8(), p02AE)
    Exit Function
  End Select
  Select Case p02B0
  Case 0
    fn0D66 = 0
  Case 1
    fn0D66 = 1
  Case 2, 3
    fn0D66 = 2
  Case 4
    fn0D66 = 4
  Case 5
    fn0D66 = 4
  Case 6
    fn0D66 = 8
  Case 7
    fn0D66 = 2
  Case gTypeFixString
    fn0D66 = UDiv(p02A8(p02AE - 1) + 1, 2)
  Case gc79CC
    l02B4 = p02A8(p02AE - 1)
    If Not fn0CFD(l02B4) Then
      ShowErrMsg mc00E4
      Exit Function
    End If
    l02B4 = UDiv(l02B4, 2)
    fn0D66 = UDiv(SomeDataBuff(l02B4 + mc02B6) + 1, 2)
  Case gc79CA
    fn0D66 = 2
  Case Else
    MakeErrBeep
  End Select
End Function

Function getFrmNamebyOffbyIdx(p02BA As Integer) As String
19360
  getFrmNamebyOffbyIdx = getFrmNamebyOff(g_TKStruct(p02BA).frmOffset)
End Function

Sub SetgFormNames(ByVal p02BC As String, p02BE As Integer)
19000
  If p02BE <= 0 Then
    p02BE = UBound(gFormNames) + 1
    ReDim Preserve gFormNames(p02BE)
  
  Else
    If p02BC = gFormNames(p02BE) Then Exit Sub
  End If
  
  gFormNames(p02BE) = p02BC
  bPrjFilesCreated = True

End Sub

Function fn0C30(p02C2() As Integer, p02C8 As Integer) As Integer
Dim l02CA As Integer
 Const mc02CC = 1 ' &H1%

19010
  l02CA = p02C2(p02C8 + 2)
  Select Case l02CA And &HFFF0
  Case &HC000, &HC500, &HC600
    fn0C30 = 3
  Case &HC100, &HC200
    l02CA = p02C2(p02C8 + 1)
    If l02CA And &H4000 Then
      l02CA = l02CA And &HF
    Else
      l02CA = 8
    End If
    fn0C30 = l02CA * 2 + 9
  Case &H4600
    fn0C30 = 3
  Case Else
    ShowErrMsg mc02CC
    Exit Function
  End Select
End Function
