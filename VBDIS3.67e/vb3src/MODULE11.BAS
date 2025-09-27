'Attribute VB_Name = "Decompile_MODULE11"
' Module11
Option Explicit
Const mc010C = 4130815  ' &H3F07FF&
Const mc0110 = 4130752  ' &H3F07C0&
Const mc0114 = 4130754  ' &H3F07C2&
Const mc0118 = 1026 ' &H402%
Const mc01F4 = 32 ' &H20%
Const mc0204 = 7 ' &H7%
Const mc0212 = 192 ' &HC0%
Const mc0214 = 16 ' &H10%
Const mc021E = 9 ' &H9%

Sub ReadIntToken()
16000
  CurOffset = VBCodeStart
  Get hVBFile, VBCodeStart, intTokens
  
  ConvertToken intTokens
  
End Sub

Sub ConvertToken(pToken As Integer)
16010
  
  intTokens = pToken
  
  If intTokens <= 0 Then
    tkMain.iToken = 0
    tkMain.AltToken = 0
  Else
    tkMain.iToken = Asc(Mid$(VBDIS_ControlToken, intTokens, 1))
    tkMain.iToken1 = Asc(Mid$(VBDIS_ControlToken, intTokens + 1, 1))
    tkMain.iToken2 = Asc(Mid$(VBDIS_ControlToken, intTokens + 2, 1))
   ' \ 3 cut away 1,5 bit from intTokens 7FFF->3FFF->2AAA
    tkMain.AltToken = VBDIS_FlagToken.FileData(intTokens \ 3)
  End If
  
  AltToken_9Bit = tkMain.AltToken And &H1FF
  
 '                                Array size 511->&h1FF
  LSet ATK9B_VBDat_1 = VBdis_Struct.Token9Bit(AltToken_9Bit)
  
  ATK9B_VBDat_1.ForwardWords = (ATK9B_VBDat_1.TK_Case1_8bit And &HF000) \ &H1000
  
  ATK9B_VBDat_1.M2CF2 = (ATK9B_VBDat_1.mTokenFlags And &HF000) \ &H1000
  
End Sub

Function fn09D7() As String
Dim l0166 As String
Dim l0168 As Integer

16040
  l0168 = ScanTokens()
  If tkMain.AltToken Then l0166 = fn0BD9(AltToken_9Bit)
  If Len(l0166) = 0 Then
    l0166 = "****"
  End If
  l0166 = Chr$(9) & l0166
  If ScanedTokensCount >= 0 Then l0166 = " " & ScanedTokens(ScanedTokensCount).M26A6 & l0166
  If ATK9B_VBDat_1.ForwardWords < 0 Then
    Seek hVBFile, VBCodeStart
  Else
    CurOffset = CurOffset + 2
    Seek hVBFile, CurOffset
    While CurOffset < VBCodeStart
      Get hVBFile, , Tk_ForwardBytes
      l0166 = l0166 & " +" & IntToHexStr(Tk_ForwardBytes)
      CurOffset = CurOffset + 2
    Wend
  End If
  fn09D7 = l0166
End Function

Sub ScanSubroutineCalls()
Dim j%
Dim i%
Dim Displvl1%
Dim Displvl2%

16050
  ScanOnly = 1
  If gv0BC4 Then ScanOnly = ScanOnly Or gMASK_0000_0010
  FrmMain.dum800.Value = True
  DispStatus Displvl1, CurFrmName
  
'Decompile modules
  For i = 1 To FrmFilesCount
    DispLog2 Displvl1, CurFrmName & " " & ModulFileNames(i)
    g_CurSub = 0
    ScanModuleToken1 i
    
    For j = 1 To gSubroutines
      If g_TKStruct(j).mSubroutinesNameLen = gSubroutinesNameLen Then
        If g_TKStruct(j).Size_M13F4 <> -1 Then
          DispStatus Displvl2, getFrmNamebyOffbyIdx(j)
          DecompileSubroutine j
          
          ClearLogUntil Displvl2
          
        End If
      End If
    Next
  Next
  
  ScanOnly = 0
  FrmMain.dum700.Value = True
  
  ClearLogUntil Displvl1
  
  NewFileExt = 0
End Sub

Sub DecompSubroutine()
16060
  
  SetStartAndEndOffset VBCodeStart, VBCodeEnd
  
  If VBCodeStart <= 1 Then
    MakeErrBeep
    Exit Sub
  End If
  
' Seek to start
  Seek hVBFile, VBCodeStart

If ScanOnly Then
    TKInit1
    Do While VBCodeStart < VBCodeEnd
      If ScanTokens() = &H30 Then
        Exit Do
      End If
    Loop
    sub0A0A
Else
    gv0980 = 0
    ReDim gv094E(TokenStruct1.M15B4 + 20)
    sub07EA
    gv094E(gv0980 + 1).Offs = VBCodeStart
    sub05F4 VBCodeEnd - VBCodeStart, 1
End If

End Sub

Function ScanTokens() As Integer
Dim ContiuneExec As Integer
Dim l01AE%
Dim l01B2 As Integer
Dim l01B4 As Integer
Dim l01B6 As Integer
Dim l01B8 As Integer
Dim l01BA As Integer
Dim l01BC As Integer
Dim l01BE As String
Dim l01C0 As String
Dim l01C2 As Integer
Dim lToken As T2679
Dim l01CA As Integer
Dim l01CC As Long
Dim l01CE As Long
Dim l01D0 As Long
  Const gMASK_0000_0101 = 5 ' &H5%
 Const SCAN_STARTED = 32 ' &H20%
 Const gMASK_0000_0100 = 4 ' &H4%
 Const mc01DA = 8 ' &H8%
 Const MASK_0100_0000 = 64 ' &H40%
 Const mc01F0 = 2 ' &H2%
 Const mc01F2 = 3 ' &H3%
 Const mc01F4 = 32 ' &H20%
 Const mc01F6 = 5 ' &H5%
 Const mc01F8 = 1 ' &H1%
 Const FoundIncompatibleTypes = 2048 ' &H800%
 Const mc0202 = 2 ' &H2%
 Const mc0204 = 7 ' &H7%
 Const mc0206 = 9 ' &H9%
 Const mc0208 = 10 ' &HA%
 Const mc020A = 224 ' &HE0%
 Const mc020C = 32 ' &H20%
 Const mc020E = 64 ' &H40%
 Const mc0210 = 128 ' &H80%
 Const mc0212 = 192 ' &HC0%
 Const mc0214 = 16 ' &H10%
 Const mc0216 = 2 ' &H2%
 Const mc021E = 9 ' &H9%
 Const mc0220 = 1 ' &H1%
 Const mc0224 = 512 ' &H200%
 Const mc0226 = 12 ' &HC%
 Const mc0228 = 9 ' &H9%
 Const mc022A = 10 ' &HA%
 Const mc022C = 8192 ' &H2000%

16070
'  Stop
  ReadIntToken
  lToken.Offs = CurOffset
  lToken.iTokens = intTokens
  lToken.TokenIndex = ScanedTokensCount
  
  ContiuneExec = tkMain.AltToken
  
' Forward to next token
  Tk_ForwardBytes = ATK9B_VBDat_1.ForwardWords * 2
  
  If Tk_ForwardBytes < 0 Then
  ' Forward a couple of bytes
    Get hVBFile, , Tk_ForwardBytes: VBCodeStart = CurOffset + 4 + Tk_ForwardBytes
  Else
  ' Normal Forward
    VBCodeStart = CurOffset + 2 + Tk_ForwardBytes
  End If
  
  
  If ContiuneExec = 0 Then
    ShowErrMsg &H100
    ScanTokens = &H30
    ScanedTokensCount = ScanedTokensCount + 1
    ScanedTokens(ScanedTokensCount) = lToken
    Exit Function
  End If
  
  Select Case ATK9B_VBDat_1.TK_Case1_8bit And &HF 'MASK_0000_1111
  
  Case gMASK_0000_0101  '5 -Start
    ScanTokens = SCAN_STARTED
    ScanedTokensCount = -1
    lToken.TokenIndex = -1
    ReDim ScanedTokens(64)
    

  Case gMASK_0000_0100    '4 - Finish
    ScanTokens = &H30

  Case mc01DA    '8
    lToken.M26AC = 12
    
    If ScanOnly > 0 Then   'Means not True(-1) or 0
      If ScanOnly And gMASK_0000_0010 Then
        lToken.M26B2 = "$"
        Get hVBFile, CurOffset + 2, l01AE
        Get hVBFile, , ContiuneExec
        
        If l01AE > 2 Then
          Get hVBFile, , ContiuneExec
          If ContiuneExec Then
            lToken.M26A6 = Space$(ContiuneExec): Get hVBFile, , lToken.M26A6
          End If
        End If
        
        lToken.M26A6 = """" + lToken.M26A6 + """"
      End If
    End If
  
  Case Else
  
    If ScanOnly > 0 Then   'Means not True(-1) or 0
      l01B2 = Asc(Mid$(VBDIS_FlagToken2, intTokens + 1, 1))
      l01BC = Asc(Mid$(VBDIS_FlagToken2, intTokens + 2, 1))
      
      If l01BC And gc11D6 Then
        l01BC = l01BC Xor gc11D6
      End If
      
      If l01B2 Then
        l01B2 = intTokens - l01B2
        l01B4 = Asc(Mid$(VBDIS_FlagToken2, l01B2, 1))
        l01BA = Asc(Mid$(VBDIS_FlagToken2, l01B2 + 1, 1))
        If l01B4 > gc114E Then
          Get hVBFile, VBCodeStart - 2, l01AE: l01AE = UDiv(l01AE, 2)
          If ATK9B_VBDat_1.mTokenFlags And MASK_0100_0000 Then
            lToken.M268E = l01AE
          End If
          Select Case l01BC
          Case 0
          Case gc1172
            Select Case l01B4
            Case gc1156
              l01AE = fn033D(l01AE, 2)
              l01AE = RotateRight3(l01AE)
            Case gc1152
              l01AE = RotateRight3(gVBCodeTokenWordsExE(l01AE))
            Case Else
              ShowErrMsg &H100
            End Select
            l01BA = NamesLookUpTbl(l01AE)
            If l01BA <> mc01F0 And l01BA <> mc01F2 Then
              ShowErrMsg mc01F4
            End If
            l01BA = NamesLookUpTbl(l01AE + mc01F6)
            l01BA = gvConv_bw(l01BA)
            l01AE = NamesLookUpTbl(l01AE + mc01F8)
            l01B4 = gc115A
            lToken.M26A0 = l01AE
            If l01BA = 4 Then
              l01B8 = 5
            End If
          Case gc1168
            l01BA = Asc(Mid$(gIntTokenStrBuff, l01AE, 1))
            Select Case l01B4
            Case gc1154, gc1152, gc1150
            Case gc1156
              l01CA = Asc(Mid$(StrTokens, UDiv(gVBCodeTokenWordsExE(l01AE), 2), 1))
              If l01CA < l01BA Then
                Mid$(StrTokens, UDiv(gVBCodeTokenWordsExE(l01AE), 2), 1) = Chr$(l01BA)
              ElseIf l01CA > l01BA Then
                l01BA = l01CA
                Mid$(gIntTokenStrBuff, l01AE, 1) = Chr$(l01BA)
              End If
            Case Else
              ShowErrMsg mc01F4
              l01BA = 0
            End Select
            If (l01BA And &HF) Then
              l01BA = gvConv_bw(l01BA And &HF)
            Else
              l01BA = gc119C
            End If
            If l01BA = 12 Then l01B6 = gc11C8
          Case gc1170
            If l01BA <> 8 Then ShowErrMsg FoundIncompatibleTypes
            If gVBCodeTokenWordsExE(l01AE) < 0 Then
              l01B4 = gc1154
            Else
              l01B4 = gc1150
            End If
          Case gc1178
            Get hVBFile, VBCodeStart, ContiuneExec
            If VBver_11 = 3 And ContiuneExec = &H6A63 Then
              l01BA = 12
            ElseIf VBver_11 = 2 And ContiuneExec = &H5D63 Then
              l01BA = 12
            Else
              ContiuneExec = Asc(Mid$(gIntTokenStrBuff, l01AE, 1))
              Select Case ContiuneExec And &H1F
              Case 0
                  l01BA = gc11A6
              Case mc0202
                l01BA = 3
              Case mc0204
                l01BA = 12
              Case mc0206
                l01BA = gvConv_bw(ContiuneExec And &H1F)
              Case mc0208
                MakeErrBeep
                l01BA = gvConv_bw(ContiuneExec And &H1F)
              Case Else
                l01B8 = gc11A6
              End Select
            End If
          Case gc1176
            ContiuneExec = Asc(Mid$(gIntTokenStrBuff, l01AE, 1))
            Select Case ContiuneExec And mc020A
            Case Is <= mc020C
              If gVBCodeTokenWordsExE(l01AE) < 0 Then
                l01B4 = gc1154
              End If
            Case mc020E
              l01B4 = gc1150
            Case mc0210
              l01B4 = gc1154
            Case mc0212
              l01B4 = gc1158
            Case Else
              ShowErrMsg mc01F4
              ContiuneExec = 0
            End Select
            If ContiuneExec And &HF Then
              l01BA = gvConv_bw(ContiuneExec And &H1F)
              If l01BA = gc118C Then
                l01BA = gc11A4
              ElseIf l01BA And 8 Then
                l01B6 = gc11C8
              ElseIf l01B4 = gc1150 Then
                l01B6 = gc11CA
              End If
            Else
              l01BA = gc11A4
            End If
          Case gc116E
            l01B4 = gVBCodeTokenWordsExE(l01AE + 2)
            l01BA = l01B4 And &HFF
            Select Case l01B4 And &HF00
            Case 0
              l01B4 = gc1156
            Case &H100, &H200
              l01B4 = gc1152
            Case &H500
              l01B4 = gc1150
            Case &H600
              l01B4 = gc1154
            Case Else
              ShowErrMsg FoundIncompatibleTypes
            End Select
            Get hVBFile, CurOffset + 2, ContiuneExec
            If ContiuneExec < 0 Then
              l01BA = gc11A0
            Else
              If l01BA = 0 Then
                l01BA = gc119C
              Else
                l01BA = gvConv_bw(l01BA)
                If l01BA And 8 Then
                  l01B6 = gc11C8
                Else
                  l01B6 = gc11CA
                End If
              End If
              l01BA = l01BA Or mc0214
            End If
          Case gc116A, gc116C
            If l01B4 <> gc115E Then ShowErrMsg FoundIncompatibleTypes
            ContiuneExec = SomeDataBuff(l01AE + mc0216)
            Select Case ContiuneExec
            Case Is < 0, Is >= mc021E
              ContiuneExec = mc0208
            Case mc0204
              l01B6 = gc11C8
            Case Is < mc0204
            End Select
            l01BA = gvConv_bw(ContiuneExec)
            If SomeDataBuff(l01AE + mc0220) And 1 Then
              l01BA = l01BA Or mc0214
              If l01BC <> gc116C Then ShowErrMsg FoundIncompatibleTypes
            ElseIf l01BC = gc116C Then
              ShowErrMsg FoundIncompatibleTypes
            End If
          End Select
          lToken.M2694 = l01B4
          lToken.M269A = l01BA
          If ScanOnly And gMASK_0000_0010 Then
            lToken.M26A6 = Mid$(gc114A, l01B4, 1)
            If l01BA = gc11A0 Then
              lToken.M26A6 = lToken.M26A6 + "a" + IntToHexStr(lToken.M268E * 2)
            Else
              If l01BA And mc0214 Then
                lToken.M26A6 = lToken.M26A6 + "a"
                l01BA = l01BA And &HF
              ElseIf l01B4 <> gc115A Then
                lToken.M26A6 = lToken.M26A6 + "v"
              End If
              lToken.M26A6 = lToken.M26A6 + IntToHexStr(Uint32(lToken.M268E) * 2)
              If l01BA > 1 Then
                lToken.M26A6 = lToken.M26A6 + Mid$(gc117C, l01BA, 1)
              End If
            End If
          Else
            l01BA = l01BA And &HF
          End If
          If l01B8 = 0 Then l01B8 = l01BA And &H3F
          If l01B8 >= gc11A8 Then ShowErrMsg FoundIncompatibleTypes
          l01B8 = l01B8 Or l01B6
          ContiuneExec = Asc(Mid$(VBDIS_FlagToken2, l01B2 + 2, 1))
          l01BE = Mid$(VBDIS_FlagToken2, l01B2 + 3, ContiuneExec)
        Else
          
          l01B8 = gc118C
          ContiuneExec = 0
          Select Case l01BC
          Case 0
            ContiuneExec = Asc(Mid$(VBDIS_FlagToken2, l01B2 + 2, 1))
            l01BE = Mid$(VBDIS_FlagToken2, l01B2 + 3, ContiuneExec)
          
          Case gc1174
            Get hVBFile, VBCodeStart - 4, l01B2
            l01B2 = VerDepSig + (l01B2 And &HFF) * &H18
            l01B2 = l01B2 + Asc(Mid$(VBDIS_FlagToken2, l01B2, 1)) + 1
            Get hVBFile, , ContiuneExec
            ContiuneExec = ContiuneExec + 1
            l01BE = Mid$(VBDIS_FlagToken2, l01B2, ContiuneExec)
          
          Case gc117A
            ContiuneExec = Asc(Mid$(VBDIS_FlagToken2, l01B2 + 2, 1))
            l01BE = Mid$(VBDIS_FlagToken2, l01B2 + 3, ContiuneExec)
          
          Case Else
            ShowErrMsg &H100
          
          End Select
        End If
        If ContiuneExec Then
          lToken.M26AC = l01B8 Or l01B6
          l01B8 = l01B8 And gc11D2
          
          For l01B2 = Len(l01BE) To 2 Step -1
            
            ContiuneExec = Asc(Mid$(l01BE, l01B2, 1))
            
            If ContiuneExec And gc11C4 Then
              Select Case ContiuneExec And &H1F
              Case gc11A8
                Get hVBFile, CurOffset + 2, ContiuneExec
                If lToken.M26A0 Then
                  sub0A39 lToken, ContiuneExec, l01C0
                Else
                  If ContiuneExec > 0 Then GoSub L12AD2
                End If
              Case 16
                l01B6 = 0
                ContiuneExec = 3: l01BA = ContiuneExec
                Get hVBFile, CurOffset + 2, l01AE
                If l01AE >= 0 Then GoSub L12C40
                GoSub L12C40
              Case Else
                ShowErrMsg mc0224
              End Select
            
            ElseIf ContiuneExec Then
              If (ContiuneExec And &HF) = gc118C Then
                l01B6 = ContiuneExec And gc11CA
                l01BA = l01B8 Or l01B6
                If l01BA > &H17 Then
                  ShowErrMsg FoundIncompatibleTypes
                  l01BA = l01BA Xor gc11CA
                End If
              Else
                l01BA = ContiuneExec And gc11D2
              End If
              
              GoSub L12C40
            
            End If
          Next
          
          ContiuneExec = Asc(Left$(l01BE, 1))
          
          If ContiuneExec And gc11C4 Then
          ' Found incompatible Types
            ShowErrMsg FoundIncompatibleTypes
          ElseIf ContiuneExec Then
            l01BA = ContiuneExec And gc11D0
            
            Select Case l01BA
            
            Case gc11A4, gc11A6
              If l01B8 > gc118C Then
                ContiuneExec = (ContiuneExec Xor l01BA) Or l01B8
              End If
            
            Case Is >= gc11A8
              ShowErrMsg FoundIncompatibleTypes
            End Select
          End If
          
          GoSub L12FE6
          If ScanOnly And gMASK_0000_0010 Then
            lToken.M26A6 = lToken.M26A6 + "(" + Mid$(l01C0, 2) + ")"
          End If
        End If
      ElseIf l01BC Then
        ShowErrMsg mc01F4
      End If
      
    ElseIf ScanOnly = True Then
    '&H40%
      If ATK9B_VBDat_1.mTokenFlags And MASK_0100_0000 Then
      
        Get hVBFile, VBCodeStart - 2, l01AE
        
        If tkMain.iToken2 = mc0226 Then
          
          Get hVBFile, CurOffset + 2, ContiuneExec
          
          If tkMain.iToken1 = mc0228 Then
            sub07DC l01AE, tkMain.iToken1, tkMain.iToken2
            l01AE = gVBCodeTokenWordsExE(UDiv(l01AE, 2))
          ElseIf tkMain.iToken1 <> mc022A Then
          End If
          
          sub0A47 l01AE, ContiuneExec
        
        Else
          sub07DC l01AE, tkMain.iToken1, tkMain.iToken2
        End If
      ElseIf VBver_11 = 3 Then
      
        Select Case intTokens
        Case &H4FA6
          Get hVBFile, VBCodeStart - 2, ContiuneExec
          sub0A57 ContiuneExec
        Case &H6F79
        End Select
        
        
      ElseIf VBver_11 = 2 Then
        Select Case intTokens
        Case &H4D89
          Get hVBFile, VBCodeStart - 2, ContiuneExec
          sub0A57 ContiuneExec
        Case &H6231
        End Select
      End If
      
    End If
    
  End Select
  
  If ScanOnly > 0 Then
    ScanedTokensCount = ScanedTokensCount + 1
    
'    Stop
'    On Error Resume Next
'    '^^ Added Line^^ !
   
  ' Store ScanedToken
    If ScanedTokensCount > UBound(ScanedTokens) Then ReDim Preserve ScanedTokens(ScanedTokensCount + 10)
    ScanedTokens(ScanedTokensCount) = lToken
  End If
  
Exit Function







'================================================
' GOSUB's

L12AD2:
  If ContiuneExec <= 0 Then
    If ContiuneExec = &H8000 Then
      If lToken.M26AC <> gc11A0 Then
        lToken.M26AC = gc11A0
      End If
    End If
    Return
  End If
  l01B4 = lToken.TokenIndex
  While ContiuneExec > 0
    If l01B4 < 0 Then
      If ScanOnly And gMASK_0000_0010 Then
        l01C0 = ",?" + l01C0
      End If
      ContiuneExec = ContiuneExec - 1
    Else
      l01C2 = ScanedTokens(l01B4).M26AC
      If l01C2 Then
        l01BE = l01BE + Chr$(l01C2)
        If ScanOnly And gMASK_0000_0010 Then
          l01C0 = "," + ScanedTokens(l01B4).M26B2 + l01C0
        End If
        ContiuneExec = ContiuneExec - 1
      End If
      l01B4 = ScanedTokens(l01B4).TokenIndex
    End If
  Wend
  lToken.TokenIndex = l01B4
Return

L12C40:
  If l01BA And gc11C6 Then ShowErrMsg FoundIncompatibleTypes: l01BA = l01BA Xor gc11C6
  l01B4 = lToken.TokenIndex
  Do While l01B4 > 0
    l01CA = ScanedTokens(l01B4).M26AC And gc11D2
    If l01CA Then
      If l01CA <> l01BA Then
        l01CC = gvBase2(l01BA)
        l01CE = gvBase2(l01CA)
        l01D0 = l01CC And l01CE
        If l01CC = 0 Or l01CE = 0 Then
          ShowErrMsg FoundIncompatibleTypes
        ElseIf l01D0 = 0 Then
        ElseIf l01D0 = l01CC Then
          If ScanedTokens(l01B4).M26A0 Then
            ShowErrMsg FoundIncompatibleTypes
          ElseIf ScanedTokens(l01B4).M268E > 0 Then
            sub0A66 ScanedTokens(l01B4).M268E, (ScanedTokens(l01B4).M2694 * &H20) And &HFF, (gvConv_ff(l01BA And gc11D0)), 0
          Else
            ShowErrMsg FoundIncompatibleTypes
          End If
          l01CA = ScanedTokens(l01B4).M26AC Xor l01CA
          ScanedTokens(l01B4).M26AC = l01BA Or l01CA
        ElseIf l01D0 = l01CE Then
          l01BA = l01CA
          If (ContiuneExec And &HF) = gc118C Then
            l01B8 = l01BA
            If lToken.M268E Then
              If lToken.M26A0 Then
                ShowErrMsg mc022C
              Else
                ShowErrMsg mc022C
              End If
            End If
          End If
        Else
          ShowErrMsg FoundIncompatibleTypes
        End If
      End If
      If ScanOnly And gMASK_0000_0010 Then
        l01C0 = "," + ScanedTokens(l01B4).M26B2 + l01C0
      End If
      lToken.TokenIndex = ScanedTokens(l01B4).TokenIndex
      Return
    End If
    l01B4 = ScanedTokens(l01B4).TokenIndex
  Loop
  ShowErrMsg FoundIncompatibleTypes
  If ScanOnly And gMASK_0000_0010 Then
    l01C0 = l01C0 + ",?"
  End If
  lToken.TokenIndex = 0
Return

L12FE6:
  If (ContiuneExec And gc11D0) = gc118C Then
    l01B6 = ContiuneExec And Not gc11D0
    ContiuneExec = l01B8 Or l01B6
  End If
  If ContiuneExec And 8 Then ContiuneExec = ContiuneExec And Not gc11CA
  If ContiuneExec = gc118C Then ShowErrMsg FoundIncompatibleTypes
  lToken.M26AC = ContiuneExec
  If ContiuneExec = 0 Then Return
  If lToken.M268E And lToken.M26A0 = 0 Then
    If (lToken.M269A And gc11D0) > 12 Then
      If lToken.M269A <> (ContiuneExec And gc11D0) Then
        sub0A66 lToken.M268E, lToken.M2694 * &H20, (gvConv_ff(ContiuneExec And &HF)), 0
      End If
    End If
  End If
  If ScanOnly And gMASK_0000_0010 Then
    If ContiuneExec And &HF Then
      lToken.M26B2 = Mid$(gc1184, ContiuneExec And &HF, 1)
    Else
      lToken.M26B2 = "?"
    End If
    If ContiuneExec And gc11C8 Then
      lToken.M26B2 = "R" + lToken.M26B2
    ElseIf ContiuneExec And gc11CA Then
      lToken.M26B2 = "^" + lToken.M26B2
    End If
    If Len(lToken.M26A6) = 0 Then lToken.M26A6 = lToken.M26B2
  End If
Return
End Function

Sub ScanTypeSub1()
Dim i1_FromFile%
Dim i3_FromFile%
Dim i2_FromFile%
Dim i_ToAssert30%
Dim l0236 As Integer
 Const mc023E = 224 ' &HE0%
 Const mc0240 = 231 ' &HE7%

16080
  Get hVBFile, gv0B8E + gv09B6, i1_FromFile
  Get hVBFile, , i2_FromFile
  
  Get hVBFile, , i_ToAssert30
  If i_ToAssert30 <> &H1E Then MakeErrBeep
  
  If i1_FromFile \ 2 <> i2_FromFile + 2 Then MakeErrBeep
  
  While i2_FromFile > 0
    Get hVBFile, , i3_FromFile
    i1_FromFile = UDiv(i3_FromFile, 2)
    
    l0236 = Asc(Mid$(StrTokens, i1_FromFile, 1))
    
    If l0236 = 0 Then
      If i3_FromFile And 1 Then
        If SomeDataBuff(i1_FromFile + 1) <> gv0B68 Then MakeErrBeep
        
        sub04B5 i1_FromFile, 2, mc023E, mc0204, 0
        
        If Asc(Mid$(StrTokens, i1_FromFile, 1)) <> mc0240 Then MakeErrBeep
        
      Else
        sub0A71 i1_FromFile
      End If
    End If
    i2_FromFile = i2_FromFile - 1
  Wend
  
End Sub

Sub ScanFixupsAndTokens()
Dim i%
Dim l0244%
Dim Value2%
Dim Value3%
Dim l024A As Integer
Dim l024E As T20E6
Dim DispLvl%
 Const mc0258 = -2 ' &HFFFE%
 Const mc025A = 96 ' &H60%
 Const mc025C = 224 ' &HE0%
 Const mc0268 = 10 ' &HA%
 Const mc026A = 16 ' &H10%
 Const mc_4 = 4 ' &H4%

16090

' ScanFixups
  DispStatus DispLvl, Scanning & C_LOCAL & Fixups
  intTokens = mc0258
  g_CurSub = 0
  
  Do
    Get hVBFile, , i
    Get hVBFile, , Value2
    Get hVBFile, , Value3
    
    If Value3 <> &H1E Then Exit Do
    
    If Value2 = 0 Then
      Value2 = (i - 4) \ 3
      
      While Value2 > 0
        
        Get hVBFile, , l024E
        i = UDiv(l024E.M1EFD, 2)
        
        Select Case Asc(l024E.M12C8)
        Case 6
          MsgBox "patch 6"
        
        Case 7
          sub0A82 i, 2, mc025A, mc0204, 0
        
        Case 9
          
          Select Case gVBCodeTokenWordsExE(i) And &HF000
          
          Case &H8000
            If (gVBCodeTokenWordsExE(i - 1) And &HFF00) <> &H4000 Then MakeErrBeep
            sub0A82 i, 2, mc025C, mc021E, 0
            sub0A90 fn0C4A(NewFileExt, gVBCodeTokenWordsExE(i) And &HFF), i
          
          Case &HC000
            sub0A82 i, 2, mc025C, mc021E, 0
            sub068D curFrmStruct.M_2_curr
            sub0A90 fn0B6F(gVBCodeTokenWordsExE(i)), i
          
          Case 0
            If gVBCodeTokenWordsExE(i - 1) > gv0FD2 Then
              MakeErrBeep
            Else
              sub0A66 i, mc025A, mc021E, 0
            End If
          
          Case Else
            MakeErrBeep
          End Select
        
        Case &H20
          If Not fn0CFD(gVBCodeTokenWordsExE(i - 1)) Then MakeErrBeep
          sub0A66 i, mc025A, mc0268, 0
        
        Case &H80
          sub0AA2 i
        
        Case &HC0
          If gVBCodeTokenWordsExE(i) Then MakeErrBeep
          sub0AA2 i
        
        Case Else
          ShowErrMsg mc026A
        End Select
        
        Value2 = Value2 - 1
      
      Wend
    
    Else
      
      If i \ 2 <> Value2 + 2 Then ShowErrMsg mc01F4
      
      While Value2
        
        Get hVBFile, , l0244
        i = UDiv(l0244, 2)
        l024A = Asc(Mid$(gIntTokenStrBuff, i, 1))
        
        If l024A = 0 Then
          If l0244 And 1 Then
            If gVBCodeTokenWordsExE(i + 1) <> gv0B68 Then MakeErrBeep
            sub0A82 i, 2, mc0212 Or 7, 0, 0
          
          Else
            sub0AA2 i
          End If
        
        End If
        Value2 = Value2 - 1
      Wend
    
    End If
  Loop
  
'Scanning Tokens
  DispLog2 DispLvl, Scanning & gc03F2
  CreateDeclarations NewFileExt
  
  For i = 1 To gSubroutines
  
    If g_TKStruct(i).mSubroutinesNameLen = gSubroutinesNameLen Then
    
      If (gDecompileStatus And mc_4) = 0 Then
      
        FrmMain.duma00.Text = getFrmNamebyOffbyIdx(i)
        
      End If
      
      DoEvents
      
      If g_TKStruct(i).Size_M13F4 <> -1 Then
        DecompileSubroutine i
      End If
    
    End If
  
  Next i
  
  DispLog2 DispLvl, "Creating " & C_LOCAL & gc03EA
  
  ScanSth
  
  ClearLogUntil DispLvl
End Sub

Sub ProcessFrmRes()
16100
  DoDecompile False
End Sub

Function fn0AD6(p027A As T2679) As String
Dim l027C As String
Dim l027E As Integer
Dim l0280 As Integer

16110
  If p027A.M2694 <= 1 Then Exit Function
  l027E = p027A.M2694
  l027C = Mid$(gc114A, l027E, 1)
  l0280 = p027A.M269A
  If l0280 = gc11A0 Then
    l027C = l027C + "a" + IntToHexStr(p027A.M268E * 2)
  Else
    If l0280 And mc0214 Then
      l027C = l027C + "a"
      l0280 = l0280 And &HF
    ElseIf l027E <> gc115A Then
      l027C = l027C + "v"
    End If
    l027C = l027C + IntToHexStr(p027A.M268E * 2)
    If l0280 > 1 Then
      l027C = l027C + Mid$(gc117C, l0280, 1)
    End If
  End If
  fn0AD6 = l027C
End Function

Sub sub0AE1()
Dim i
Dim Start As Long

16120
  Start = 1
  For i = 2 To 12
    gvBase2(i) = Start
    Start = Start * 2
  Next
  
  Start = &H10000
  For i = 2 To 7
    gvBase2(i + 16) = Start
    Start = Start * 2
  Next
  
  Start = &H40
  For i = 8 To 12
    gvBase2(i + 16) = Start
    Start = Start * 2
  Next
  
  gvBase2(gc118C) = mc010C
  gvBase2(gc11A4) = mc0114
  gvBase2(gc11A6) = mc0118
  gvBase2(gc11DA) = mc0110
End Sub

Sub sub09C7()
Dim l0284 As Long
16020
  LSet ATK9B_VBDat_1 = VBdis_Struct.Token9Bit(AltToken_9Bit)
  ATK9B_VBDat_1.ForwardWords = (ATK9B_VBDat_1.TK_Case1_8bit And &HF000) \ &H1000
  ATK9B_VBDat_1.M2CF2 = (ATK9B_VBDat_1.mTokenFlags And &HF000) \ &H1000
End Sub

Sub Init_gVBObjsDefs()

16030
  Erase gVBObjsDefs
  
  ReDim gv0FF6(0)
  
  gControlCount1 = 0
  gVBObjsDefs(1).Name_4 = "Form"
    gVBObjsDefs(1).id_5 = &HD + 1
    gVBObjsDefs(1).id_add_5 = 5
    
  gVBObjsDefs(2).Name_4 = "MDIForm"
    gVBObjsDefs(2).id_5 = &H14 + 1
  
  gVBObjsDefs(4).Name_4 = "Control"
    gVBObjsDefs(4).id_5 = 0
  
  gVBObjsDefs(5).Name_4 = "Controls"
    gVBObjsDefs(5).id_add_5 = -4
  
  gVBObjsDefs(8).Name_4 = "Forms"
    gVBObjsDefs(8).id_add_5 = -1
    
    gVBObjsDefs(&HB).Name_4 = "Database"
      gVBObjsDefs(&HB).id_5 = &H20 + 1
      gVBObjsDefs(&HB).id_add_5 = &HC
    
    gVBObjsDefs(&HC).Name_4 = "TableDefs"
      gVBObjsDefs(&HC).id_5 = &H1D + 1
      gVBObjsDefs(&HC).id_add_5 = -&HD
    
    gVBObjsDefs(&HD).Name_4 = "TableDef"
      gVBObjsDefs(&HD).id_5 = &H1A + 1
      gVBObjsDefs(&HD).id_add_5 = &HF
    
    gVBObjsDefs(&HE).Name_4 = "Dynaset"
      gVBObjsDefs(&HE).id_5 = &H1F + 1
      gVBObjsDefs(&HE).id_add_5 = &HF
    
    gVBObjsDefs(&HF).Name_4 = "Fields"
      gVBObjsDefs(&HF).id_5 = &H1C + 1
      gVBObjsDefs(&HF).id_add_5 = -&H11
    
    gVBObjsDefs(&H10).Name_4 = "Indexes"
      gVBObjsDefs(&H10).id_5 = &H1E + 1
      gVBObjsDefs(&H10).id_add_5 = -&H12
    
    gVBObjsDefs(&H11).Name_4 = "Field"
      gVBObjsDefs(&H11).id_5 = &H19 + 1
    
    gVBObjsDefs(&H12).Name_4 = "Index"
      gVBObjsDefs(&H12).id_5 = &H1B + 1
  
  If VBver_11 = 3 Then
  
    gVBObjsDefs(&H14).Name_4 = "Object"
      gVBObjsDefs(&H14).id_5 = &H24 + 1
    
    gVBObjsDefs(&H15).Name_4 = "QueryDef"
      gVBObjsDefs(&H15).id_5 = &H21 + 1
    
    gVBObjsDefs(&H16).Name_4 = "Table"
      gVBObjsDefs(&H16).id_5 = &H22 + 1
      gVBObjsDefs(&H16).id_add_5 = &HF
    
    gVBObjsDefs(&H17).Name_4 = "Snapshot"
      gVBObjsDefs(&H17).id_5 = &H23 + 1
      gVBObjsDefs(&H17).id_add_5 = &HF
    
    gVBObjsDefs(&H32).Name_4 = "Printer"
      gVBObjsDefs(&H32).id_5 = &HC + 1
    
    gVBObjsDefs(&H33).Name_4 = "Screen"
      gVBObjsDefs(&H33).id_5 = &HE + 1
    
    gVBObjsDefs(&H34).Name_4 = "Clipboard"
      gVBObjsDefs(&H34).id_5 = &HF + 1
    
    gVBObjsDefs(&H3D).Name_4 = "App"
      gVBObjsDefs(&H3D).id_5 = &H15 + 1
  
  ElseIf VBver_11 = 2 Then
    gVBObjsDefs(&H2C).Name_4 = "Printer"
      gVBObjsDefs(&H2C).id_5 = &HC + 1
    
    gVBObjsDefs(&H2D).Name_4 = "Screen"
      gVBObjsDefs(&H2D).id_5 = &HE + 1
    
    gVBObjsDefs(&H2E).Name_4 = "Clipboard"
      gVBObjsDefs(&H2E).id_5 = &HF + 1
    
    gVBObjsDefs(&H37).Name_4 = "App"
      gVBObjsDefs(&H37).id_5 = &H15 + 1
  Else
    MakeErrBeep
  End If
End Sub
