'Attribute VB_Name = "MODULE16"
' Module16
Option Explicit
Const mc000C = -1 ' &HFFFF%
Dim m000E() As Integer
Dim m0040 As Integer
Dim mMaxPos As Integer
Dim mCurPos As Integer
Const mc0076 = 8 ' &H8%
Const mc0078 = 9 ' &H9%
Const mc007A = 64 ' &H40%
Const mc008C = 5 ' &H5%
Const mc009E = 227 ' &HE3%
Const mc00BA = 15 ' &HF%
Const mc00BE = 1 ' &H1%
Const mc00C2 = 21 ' &H15%
Const mc00CA = 14 ' &HE%
Const mc0108 = 2048 ' &H800%
Const mc010C = 2 ' &H2%
Const mc0116 = 15 ' &HF%
Const mc0118 = 13 ' &HD%
Const mc011A = 48 ' &H30%
Const mc012C = 8 ' &H8%
Const mc012E = 32 ' &H20%
Const mc0130 = 8192 ' &H2000%
Const mc0132 = 128 ' &H80%
Const mask_1110_0000 = 224 ' &HE0%
Const mc013E = 96 ' &H60%
Const mc0142 = 192 ' &HC0%
Const mc0144 = 160 ' &HA0%
Const mc0146 = 32 ' &H20%
Const mc014E = 2 ' &H2%
Const mc0150 = 3 ' &H3%
'Const m0160 = 0 ' &H0%
Const mc0162 = 225 ' &HE1%
Const mc0164 = 226 ' &HE2%
Const mc0166 = 233 ' &HE9%
Const mc016E = 231 ' &HE7%
Const mc0184 = 160 ' &HA0%
Const mc01B0 = 12 ' &HC%
Const mc01B2 = 32 ' &H20%
Const mc01B4 = 2 ' &H2%
Const mc01B6 = 6 ' &H6%
Const mc01BA = 7 ' &H7%
Const mc01BC = 8 ' &H8%
Const mc01BE = 9 ' &H9%
Const mc01C0 = 11 ' &HB%
Const mc01DC = 16 ' &H10%

Sub sub03B4(p0046 As Integer, p0048 As Integer, p004A As Integer)
24000
  g_TKStruct(p0046).M2503 = p0048
  g_TKStruct(p0046).M250C = p004A
End Sub

Sub sub0DE6(p0054 As Integer, p0056 As Integer)
Dim l0058 As Integer
Dim l005A As String
Dim l005C As String
Dim l005E As Integer
Dim l0060 As Integer
Dim l0062 As Integer
Dim l0064%
 Const mc0076 = 8 ' &H8%
 Const mc0078 = 9 ' &H9%
 Const mc007A = 64 ' &H40%

24010
  l0058 = g_TKStruct(p0056).M2503
  l005A = g_EventNames(l0058).M2C33
  l005C = g_EventNames(l0058).M1CC7
  If g_TKStruct(p0056).M250C Then
    If Len(l005A) Then
      l005A = "Index%, " & l005A
    Else
      l005A = "Index%"
    End If
    l005C = Chr$(1) & l005C
  End If
  l0060 = gVBCodeTokenWordsExE(p0054) - 6
  If l0060 Mod 4 Then MakeErrBeep
  l0060 = l0060 \ 4
  If l0060 >= Len(l005C) Then
    MakeErrBeep
    Exit Sub
  End If
  l0060 = Len(l005C) - 1 - l0060
  l0062 = Asc(Mid$(l005C, l0060 + 1, 1))
  l005E = 0
  While l0060 > 0
    l005E = InStr(l005E + 1, l005A, ",")
    If l005E = 0 Then MakeErrBeep
    l0060 = l0060 - 1
  Wend
  
  l005A = Trim$(Mid$(l005A, l005E + 1))
  l005E = InStr(l005A, ",")
  If l005E Then l005A = Left$(l005A, l005E - 1)
  l005E = InStr(l005A, " ")
  
  If l005E Then
    l005A = RTrim$(Left$(l005A, l005E))
  Else
    l005E = InStr("%&!#@?$", Right$(l005A, 1))
    If l005E Then l005A = Left$(l005A, Len(l005A) - 1)
  End If
  
  Select Case l0062
  Case 1 To 5
  Case 6: l0062 = 7
  Case 7: l0062 = mc0076
  Case 8: l0062 = mc0078
  Case Else: MakeErrBeep
  End Select
  
  sub0A66 p0054, mc007A, l0062, l0064
  sub0A90 l005A, p0054
End Sub

Sub sub0A0A()
24040
End Sub

Private Sub sub0E0B(p007C As Integer)
Dim l007E
Dim l0080 As Integer
Dim l0082 As Integer
Dim l0084 As Integer

 Const mc008C = 5 ' &H5%
 Const mc009E = 227 ' &HE3%

24050
  intTokens = &HFFFE
  ReDim m000E%(gSubroutines)
  
  m0040 = 0
  
  For l007E = 1 To gSubroutines
    If g_TKStruct(l007E).mSubroutinesNameLen = p007C Then
      l0080 = g_TKStruct(l007E).M251F
      l0082 = g_TKStruct(l007E).TKFrmRes
      If l0080 Then
        m000E(l0080) = l007E
        m0040 = m0040 + 1
      End If
      
      If NamesLookUpTbl(l0082 + mc008C) Then l0084 = l0084 + 1
    End If
  Next
  If l0084 Then
    If gFrmStruct(NewFileExt).M_2_curr Then l0082 = &H1A \ 2 Else l0082 = 6 \ 2
    Mid$(gIntTokenStrBuff, l0082, l0084) = String$(l0084, mc009E)
  End If
End Sub

Function fn0DDA(p00A4%) As String
Dim l00A6 As Integer
Dim l00A8 As Integer
Dim l00AA As Integer
Dim l00AC As Integer
 Const mc00BA = 15 ' &HF%
 Const mc00BE = 1 ' &H1%
 Const mc00C2 = 21 ' &H15%
 Const mc00CA = 14 ' &HE%

24060
  l00AC = fn0D0A(p00A4)
  If l00AC > 0 Then fn0DDA = gVBCodeTokenAddData(l00AC): Exit Function
  l00A6 = fn033D(p00A4, 2)
  l00AA = gVBCodeTokenWordsExE(p00A4)
  Select Case -l00AC
  Case mc00BA
    l00A8 = RotateRight3(l00AA)
    l00A8 = NamesLookUpTbl(l00A8 + mc00BE)
    fn0DDA = getFrmNamebyOffbyIdx(l00A8)
  Case mc00C2
    l00A6 = gVBCodeTokenWordsFinal(p00A4)
    fn0DDA = getFrmNamebyOffbyIdx(m000E(l00A6))
  Case mc00CA
    fn0DDA = fn0CDA(UDiv(l00AA, 2))
  Case Else
    fn0DDA = fnPreFix(-l00AC) & IntToHexStr(l00A6)
  End Select
End Function

Function fn0E18(p00DA As Integer) As Integer
24070
On Error Resume Next
  fn0E18 = gVBCodeTokenWordsFinal(p00DA)
End Function

Private Sub sub0E27(ByVal pv00DC As Integer)
Dim l00DE As Integer
Dim l00E0
Dim l00E8 As T14EB
Dim l00EA As Integer

24080
  If pv00DC Then
    ReDim gVBCodeTokenWordsFinal(gVBCodeSizeWords)
    Exit Sub
  End If
  If gSubroutines = 0 Then MakeErrBeep: Exit Sub
  For pv00DC = 1 To FrmFilesCount
    l00DE = 0
    l00EA = gFrmStruct(pv00DC).mSubroutinesNameLen
    For l00E0 = 1 To gSubroutines
      If g_TKStruct(l00E0).mSubroutinesNameLen = l00EA Then
        Get hVBFile, VBAbsOffs(g_TKStruct(l00E0).TKFrmRes), l00E8
        If l00E8.M15BE And &H1000000 Then
          l00DE = l00DE + 1
          g_TKStruct(l00E0).M251F = l00DE
        ElseIf g_TKStruct(l00E0).segm0 Then
          l00DE = l00DE + 1
          g_TKStruct(l00E0).M251F = l00DE
        End If
      End If
    Next
  Next
End Sub

Sub sub0A39(p00F6 As T2679, p00F8 As Integer, p00FA As String)
Dim l00FC As String
Dim l0100 As Integer
Dim l0102 As Integer
Dim l0104 As Integer
Dim l0106 As Integer
 Const mc0108 = 2048 ' &H800%
 Const MASK_0000_0010 = 2 ' &H2%
 Const mc0116 = 15 ' &HF%
 Const mc0118 = 13 ' &HD%
 Const mc011A = 48 ' &H30%

24090
  l0104 = p00F6.TokenIndex
  l0106 = p00F6.M26A0
  l00FC = g_TKStruct(l0106).M1CC7
  p00F8 = p00F8 + 1
  If Len(l00FC) <> p00F8 Then
    If Len(l00FC) Then
      ShowErrMsg mc0108
    Else
      l00FC = String$(p00F8, 0)
    End If
  End If
  While p00F8 > 1
    If l0104 < 0 Then
      If ScanOnly And MASK_0000_0010 Then
        p00FA = ",?" + p00FA
      End If
      p00F8 = p00F8 - 1
    Else
      l0100 = ScanedTokens(l0104).M26AC
      If l0100 Then
        l0102 = Asc(Mid$(l00FC, p00F8, 1))
        If l0102 = 0 Then
          Mid$(l00FC, p00F8, 1) = Chr$(l0100)
        ElseIf (l0100 Xor l0102) And &HF Then
          If g_TKStruct(l0106).Size_M13F4 = -1 Then
            If (l0102 And mc0116) <> mc0118 Then
              l0102 = mc0118 Or (l0102 And mc011A)
              Mid$(l00FC, p00F8, 1) = Chr$(l0102)
            End If
          ElseIf (l0102 And &HF) > (l0100 And &HF) Then
            Mid$(l00FC, p00F8, 1) = Chr$(l0100)
          End If
        ElseIf (l0100 Or l0102) <> l0102 Then
          Mid$(l00FC, p00F8, 1) = Chr$(l0100 Or l0102)
        End If
        If ScanOnly And MASK_0000_0010 Then
          p00FA = "," + ScanedTokens(l0104).M26B2 + p00FA
        End If
        p00F8 = p00F8 - 1
      End If
      l0104 = ScanedTokens(l0104).TokenIndex
    End If
  Wend
  p00F6.TokenIndex = l0104
  If l00FC <> g_TKStruct(l0106).M1CC7 Then
    g_TKStruct(l0106).M1CC7 = l00FC
    bPrjFilesCreated = True
  End If
End Sub

Sub ScanSth()
Dim Offs As Integer
Dim i
Dim Token As Integer
Dim l0126 As Integer
Dim idx2 As Integer
Dim idx
 Const mc012C = 8 ' &H8%
 Const mc012E = 32 ' &H20%
 Const mc0130 = 8192 ' &H2000%
 Const mc0132 = 128 ' &H80%
 Const mask_1110_0000 = 224 ' &HE0%
 Const mc013E = 96 ' &H60%
 Const mc0142 = 192 ' &HC0%
 Const mc0144 = 160 ' &HA0%
 Const mc0146 = 32 ' &H20%

24100
  intTokens = &HFFFE
  Offs = mCurPos - 1
  If Asc(Mid$(gIntTokenStrBuff, mCurPos, 1)) And mc012C Then Offs = Offs - 1
  
' Init Buff with 00
  For i = 1 To Offs: gVBCodeTokenWordsFinal(i) = 0: Next
  
  For i = 1 To m0040
    Offs = 0
    For idx = gVBCodeSizeWords To 1 Step -1
      If gVBCodeTokenWordsFinal(idx) = i Then Offs = idx: Exit For
    Next
    
    If Offs Then
      Do While Offs > 3
        Offs = Offs - 1
        Token = gVBCodeTokenWordsFinal(Offs)
        If Token And Token <> i Then
          For idx = 1 To Offs: gVBCodeTokenWordsFinal(idx) = 0: Next
          If Offs >= mCurPos Then
            mCurPos = Offs + 1
          End If
          Exit For
        End If
      Loop
      Exit For
    End If
  Next
  
  mMaxPos = gVBCodeSizeWords
  For i = m0040 To 1 Step -1
    l0126 = m000E(i)
    If g_TKStruct(l0126).Size_M13F4 <> -1 Then
      idx2 = 0
      For idx = 1 To gVBCodeSizeWords
        If gVBCodeTokenWordsFinal(idx) = i Then idx2 = idx: Exit For
      Next
      
      If idx2 = 0 Then
        g_TKStruct(l0126).mCodeSize = mMaxPos
      Else
        mMaxPos = idx2
        g_TKStruct(l0126).mCodeSize = idx2
      End If
      
      idx2 = g_TKStruct(l0126).M2515
      If idx2 Then
        If gVBCodeTokenWordsFinal(mMaxPos) > i Then
          mMaxPos = mMaxPos - 1
        End If
        Token = gVBCodeTokenWordsExE(mMaxPos)
        
        Do While Token <> idx2
          GoSub L1B302
          
          If Token = i Then
            If NamesLookUpTbl(g_TKStruct(l0126).TKFrmRes + mc008C) = 6 Then
              mMaxPos = mMaxPos - 1
              GoSub L1B302
            End If
          End If
          
          If Token = i Then
            g_TKStruct(l0126).mCodeSize = mMaxPos
            gVBCodeTokenWordsFinal(mMaxPos) = i
            Token = idx2
          
          ElseIf Token Then
            ShowErrMsg mc012E
            Token = 0
            Exit Do
          
          Else
            mMaxPos = mMaxPos - 1
            Token = gVBCodeTokenWordsExE(mMaxPos)
            
            If mMaxPos < mCurPos Then
              ShowErrMsg mc0130
              Token = 0
              Exit Do
            End If
          
          End If
          
        Loop
        
        If Token = idx2 Then
          Token = g_TKStruct(l0126).TKFrmRes
          sub0A82 mMaxPos, 0, mc0132, NamesLookUpTbl(Token + mc008C), 0
          If gVBCodeTokenWords(mMaxPos) <> -mc00C2 Then
            gVBCodeTokenWords(mMaxPos) = -mc00C2
            gWriteOutPut = True
          End If
        Else
          ShowErrMsg mc0130
        End If
      End If
    End If
  Next
  If mMaxPos <> mCurPos Then
    mCurPos = mMaxPos
  End If
  idx2 = 0
  For Offs = mCurPos To gVBCodeSizeWords
    Token = gVBCodeTokenWordsFinal(Offs)
    If Token > m0040 Then
      Token = 0
      gVBCodeTokenWordsFinal(Offs) = Token
    End If
    
    i = Asc(Mid$(gIntTokenStrBuff, Offs, 1))
    
    If Token Then
      l0126 = m000E(Token)
      If Token < idx2 Then
        
        MakeErrBeep
        For idx = 1 To Offs - 1: gVBCodeTokenWordsFinal(idx) = 0: Next
        idx2 = Token
      
      ElseIf Token > idx2 Then
        
        mMaxPos = Offs
        
        Do
          idx2 = idx2 + 1
          g_TKStruct(m000E(idx2)).mCodeSize = mMaxPos
        Loop While idx2 < Token
        
      End If
      
      Select Case i And mask_1110_0000
      Case mc007A
        If g_TKStruct(l0126).M2503 Then
          sub0DE6 Offs, l0126
        End If
      Case mc0132
      Case mc013E
      Case &HE0
      Case mc0142
      Case mc0144
      Case Else
        ShowErrMsg mc0130
      End Select
      
      
    Else
      If i >= mc0146 Then gVBCodeTokenWordsFinal(Offs) = idx2
    End If
  Next
Exit Sub

L1B302:
  Token = gVBCodeTokenWordsExE(mMaxPos)
  
  If Token And Token <> idx2 Then
    Token = 0
  Else
    Token = gVBCodeTokenWordsFinal(mMaxPos)
    If Token = 0 Then
    ElseIf Token <> i Then
      ShowErrMsg mc012E
      Return
    End If
    Offs = Asc(Mid$(gIntTokenStrBuff, mMaxPos, 1))
    If Offs And (Offs And mask_1110_0000) <> mc0132 Then
      If Token = i Then Token = 0
    Else
      If Token = 0 Then Token = i
    End If
  End If
Return
End Sub

Sub sub0A47(ByVal pv0148 As Integer, p014A As Integer)
Dim l014C As Integer
 Const mc014E = 2 ' &H2%
 Const mc0150 = 3 ' &H3%

24110
  pv0148 = RotateRight3(pv0148)
  Select Case NamesLookUpTbl(pv0148)
  Case mc014E, mc0150
    l014C = NamesLookUpTbl(pv0148 + mc00BE)
    If l014C Then
    If Len(g_TKStruct(l014C).M1CC7) <> p014A + 1 Then
      If Len(g_TKStruct(l014C).M1CC7) = 0 Then
        g_TKStruct(l014C).M1CC7 = Chr$(NamesLookUpTbl(pv0148 + mc008C)) & String$(p014A, 0)
      Else
        MakeErrBeep
      End If
    End If
    Else
      MakeErrBeep
    End If
  Case Else
    MakeErrBeep
  End Select
End Sub

Sub Write_main_txt()
Dim pos As Integer
Dim CurToken As Integer

 Const mc0162 = 225 ' &HE1%
 Const mc0164 = 226 ' &HE2%
 Const mc0166 = 233 ' &HE9%
 Const mc016E = 231 ' &HE7%

24120
  Writing_main_txt = True
  DisOutput "' main.txt - global definitions": DisOutputWrite
  pos = 3
  
  If Asc(Right$(StrTokens, 1)) <> EOS_TOKEN Then
    MakeErrBeep
    Mid$(StrTokens, gv09BA + 1, 1) = Chr$(EOS_TOKEN)
  End If
  
  Do
    CurToken = Asc(Mid$(StrTokens, pos, 1))
    
    Select Case CurToken And mask_1110_0000
    
    Case 0, &HE0
      
      Select Case CurToken
      Case 0
        While Asc(Mid$(StrTokens, pos, 1)) = 0
          pos = pos + 1
        Wend
      
      Case mc0162
        
        sub0D56 pos
       
       'Find StrEnd
        pos = InStr(pos, StrTokens, Chr$(mc0164))
        If pos = 0 Then
          ShowErrMsg mc012E
          Exit Do
        Else
          pos = pos + 1
        End If
      
    ' Finish
      Case EOS_TOKEN: Exit Do
      
      Case 8 To &H1F
        pos = pos + 1
        
      Case mc009E: MakeErrBeep: pos = pos + 1
      
      Case mc0166: pos = pos + 1
      
      Case &HE0
        If gv78AA(pos) Then MakeErrBeep
        pos = pos + 1
        
      Case mc016E
        sub076F pos, StrTokens, True
        DisOutputWrite
        pos = pos + 1
        
      Case Else
        pos = pos + 1
        ShowErrMsg mc012E
        
      End Select
    
    Case mc0146
      ShowErrMsg mc012E
      DisOutput "Global " & fn0CDA(pos) & " ' As ?"
      DisOutputWrite
      pos = pos + 1
    
    Case mc0142
      sub076F pos, StrTokens, True
      DisOutputWrite
      pos = pos + 1
    
    Case mc0144
      sub07A1 pos, SomeDataBuff(), StrTokens, True
      DisOutputWrite
      pos = pos + 1
    
    Case Else
      MakeErrBeep
      pos = pos + 1
    
    End Select
  Loop
End Sub

Sub CreateDeclarations(ModulNum As Integer)
Dim i%
Dim l017A As Integer
Dim l017C As Integer
 Const mc0184 = 160 ' &HA0%

24130
  If ModulNum <> NewFileExt Then MakeErrBeep: ScanModuleToken1 ModulNum
  
  sub0E0B gSubroutinesNameLen
  
  If ScanOnly Then
    mCurPos = gVBCodeSizeWords
    mMaxPos = mCurPos
    Exit Sub
  End If
  
  If m0040 Then
    mCurPos = g_TKStruct(m000E(1)).mCodeSize
  Else
    mCurPos = gVBCodeSizeWords
  End If
  
  Writing_main_txt = True
  
  If mc000C Then DisOutput "Option Explicit": DisOutputWrite
  
  For i = 1 To gSubroutines
    If g_TKStruct(i).mSubroutinesNameLen = gSubroutinesNameLen Then
      If g_TKStruct(i).segm0 = 0 Then
        sub0785 i
      End If
    End If
  Next
  
  If curFrmStruct.M_2_curr Then
    i = &H1A \ 2
  Else
    i = 3
  End If
  
  Do While i < gVBCodeSizeWords
    l017A = gVBCodeTokenWordsFinal(i)
    l017C = Asc(Mid$(gIntTokenStrBuff, i, 1))
    
    Select Case l017C And mask_1110_0000
    Case 0, &HE0
      Select Case l017C
      Case 0
        While Asc(Mid$(gIntTokenStrBuff, i, 1)) = 0
          i = i + 1
        Wend
      
      Case EOS_TOKEN
        Exit Do
      
      Case mc009E
        i = i + 1
      
      Case 8 To &H1F
        i = i + 1
      
      Case mc0166
        i = i + 2
      
      Case mc016E
        i = i + 1
      
      Case Else
        MakeErrBeep
        i = i + 1
      
      End Select
    
    Case mc0146
      MakeErrBeep
      If l017A = 0 Then DisOutput "Dim " & fn0DDA(i) & " ' As ?": DisOutputWrite
      i = i + 1
    Case mc0142
      sub076F i, gIntTokenStrBuff, False: DisOutputWrite
      i = i + 1
    Case mc0184
      i = i + 1
    Case mc013E
    If l017A = 0 Then
      DisOutput "Dim "
      sub07A1 i, gVBCodeTokenWordsExE(), gIntTokenStrBuff, False
      DisOutputWrite
    End If
      i = i + 1
    Case mc0132, mc007A
      i = i + 1
    Case Else
      MakeErrBeep
      i = i + 1
    End Select
  Loop

End Sub

Sub sub07F8()
Dim l0186
Dim l0188 As Integer
Dim l018A As Integer
Dim l018C As Integer
Dim l018E As Integer

24140
  l018C = g_TKStruct(g_CurSub).M251F
  mMaxPos = g_TKStruct(g_CurSub).mCodeSize
  If ScanOnly Then
    
    If mCurPos = 0 Then
    End If
    
    If mMaxPos = 0 Then
      mMaxPos = mCurPos
    End If
    
    Exit Sub
    
  End If
  
  If TokenStruct1.Type_3 <= &H16 Then
    If TokenStruct1.Type_3 <> &H16 Then MakeErrBeep: Exit Sub
  End If
  
  If mMaxPos = 0 Then
  
    For l0186 = 1 To gVBCodeSizeWords
      If gVBCodeTokenWordsFinal(l0186) = l018C Then mMaxPos = l0186: Exit For
    Next
    
    If mMaxPos = 0 Then Exit Sub
    g_TKStruct(g_CurSub).mCodeSize = mMaxPos
    
  End If
  
  l0186 = mMaxPos
  If gVBCodeTokenWordsFinal(l0186) > l018C Then Exit Sub
  
  l018E = g_TKStruct(g_CurSub).M2515
  
  If l018E Then
    If gVBCodeTokenWordsExE(l0186) = l018E Or gVBCodeTokenWordsExE(l0186) = 0 Then
      If gVBCodeTokenWords(l0186) <> -mc00C2 Then
        gVBCodeTokenWords(l0186) = -mc00C2
        gWriteOutPut = True
      End If
      For l0186 = l0186 + 1 To gVBCodeSizeWords
        If gVBCodeTokenWordsFinal(l0186) = l018C Then Exit For
      Next
      If l0186 > gVBCodeSizeWords Then l0186 = 0
    End If
  End If
  
  Do While l0186
    If gVBCodeTokenWordsFinal(l0186) > l018C Then Exit Do
    l0188 = Asc(Mid$(gIntTokenStrBuff, l0186, 1))
    Select Case l0188 And mask_1110_0000
    Case mc013E
      DisOutput "Static "
      sub07A1 l0186, gVBCodeTokenWordsExE(), gIntTokenStrBuff, False
      DisOutputWrite
    Case mc0132
      If gVBCodeTokenWordsExE(l0186) <> l018E Then
        If gVBCodeTokenWordsExE(l0186) And &H8001 Then
            DisOutput "Dim "
            sub07A1 l0186, gVBCodeTokenWordsExE(), gIntTokenStrBuff, False
            DisOutputWrite
        End If
      End If
    Case mc0144
      l018A = Asc(Mid$(StrTokens, UDiv(gVBCodeTokenWordsExE(l0186), 2), 1))
      If (l018A And mask_1110_0000) = mc0142 Then
        DisOutput "' "
        sub076F UDiv(gVBCodeTokenWordsExE(l0186), 2), StrTokens, True
        DisOutputWrite
      End If
    Case mc0142
      DisOutput "' "
      sub076F l0186, gIntTokenStrBuff, False: DisOutputWrite
    End Select
    If l0186 >= gVBCodeSizeWords Then
      l0186 = 0
    Else
      For l0186 = l0186 + 1 To gVBCodeSizeWords
        If gVBCodeTokenWordsFinal(l0186) = l018C Then Exit For
      Next
      If l0186 > gVBCodeSizeWords Then Exit Do
    End If
  Loop
  
End Sub

Sub sub0792(p0196 As Integer)
Dim l0198 As Integer
Dim l019A As Integer
Dim l019C As Integer
Dim l019E As Integer
Dim l01A0 As Integer
Dim l01A2 As Integer
Dim l01A4 As String
Dim l01A6 As Integer
Dim l01A8 As String
Dim l01AA As Integer
Dim l01AE As Integer
 Const mc01B0 = 12 ' &HC%
 Const mc01B2 = 32 ' &H20%
 Const mc01B4 = 2 ' &H2%
 Const mc01B6 = 6 ' &H6%
 Const mc01BA = 7 ' &H7%
 Const mc01BC = 8 ' &H8%
 Const mc01BE = 9 ' &H9%
 Const mc01C0 = 11 ' &HB%

24150
  If ScanOnly Then Exit Sub
  DisOutput " ("
  l0198 = g_TKStruct(p0196).M2503
  If l0198 Then
    If g_TKStruct(p0196).M250C Then
      DisOutput ("Index%")
      If Len(g_EventNames(l0198).M2C33) Then DisOutput ", "
    End If
    DisOutput g_EventNames(l0198).M2C33
  ElseIf g_TKStruct(p0196).Size_M13F4 = -1 Then
    l01A4 = g_TKStruct(p0196).M1CC7
    l01A8 = l01A4
    GoSub L1CB28
  Else
    l01A2 = g_TKStruct(p0196).M251F
    For l0198 = 1 To gVBCodeSizeWords
      If gVBCodeTokenWordsFinal(l0198) = l01A2 Then Exit For
    Next
    If l0198 > gVBCodeSizeWords Then l0198 = 0
    l01A4 = g_TKStruct(p0196).M1CC7
    l01A6 = g_TKStruct(p0196).mCodeSize
    l01A8 = g_TKStruct(p0196).M2533
    If l0198 <> l01A6 Then
      If l01A6 = 0 Or l01A6 < l0198 Then
        l01A6 = l0198
        g_TKStruct(p0196).mCodeSize = l0198
        bPrjFilesCreated = True
      ElseIf l0198 = 0 Then
        l0198 = l01A6
      Else
        l01A6 = l0198
      End If
    End If
    l01A0 = g_TKStruct(p0196).M2515
    If l01A0 Then
      For l0198 = l0198 + 1 To gVBCodeSizeWords
        If gVBCodeTokenWordsFinal(l0198) = l01A2 Then Exit For
      Next
      If l0198 > gVBCodeSizeWords Then l0198 = 0
      l01A6 = l0198
    End If
    If Len(l01A4) > 1 And Len(l01A8) <> Len(l01A4) Then
      l01A8 = Chr$(0)
      If l01A6 Then
        Do While l0198
          l019C = Asc(Mid$(gIntTokenStrBuff, l0198, 1))
          Select Case l019C And mask_1110_0000
          Case mc0132
            If (gVBCodeTokenWordsExE(l0198) And &H8001) = 0 Then
              l01A8 = l01A8 + Chr$(l0198 - l01A6)
            Else
              Exit Do
            End If
          Case mc007A
            l01A8 = l01A8 + Chr$(l0198 - l01A6)
          End Select
          For l0198 = l0198 + 1 To gVBCodeSizeWords
            If gVBCodeTokenWordsFinal(l0198) = l01A2 Then Exit For
          Next
          If l0198 > gVBCodeSizeWords Then l0198 = 0
        Loop
      Else
      End If
    End If
    If Len(l01A8) <> Len(l01A4) Then l01A4 = ""
    l0198 = l01A6
    Do While l0198
      l019E = gVBCodeTokenWordsFinal(l0198)
      If l019E <> 0 And l019E <> l01A2 Then Exit Do
      l019C = Asc(Mid$(gIntTokenStrBuff, l0198, 1))
      Select Case l019C And mask_1110_0000
      Case mc0132
        If (gVBCodeTokenWordsExE(l0198) And &H8001) = 0 Then
          If l019A Then DisOutput ", "
          l019A = l019A + 1
          DisOutput "ByVal "
          sub07A1 l0198, gVBCodeTokenWordsExE(), gIntTokenStrBuff, False
        Else
          Exit Do
        End If
      Case mc007A
        If l019A Then DisOutput ", "
        l019A = l019A + 1
        If Len(l01A4) > l019A Then
          l01AE = Asc(Mid$(l01A4, l019A + 1, 1))
          If l01AE Then
            If (l01AE And mc0116) = mc01B0 Then
              If l01AE And mc01B2 Then
              Else
                DisOutput "ByVal "
              End If
            ElseIf l01AE And &H38 Then
            Else
              DisOutput "ByVal "
            End If
          End If
        End If
        sub07A1 l0198, gVBCodeTokenWordsExE(), gIntTokenStrBuff, False
      End Select
      For l0198 = l0198 + 1 To gVBCodeSizeWords
        If gVBCodeTokenWordsFinal(l0198) = l01A2 Then Exit For
        If gVBCodeTokenWords(l0198) Then
          Exit For
        End If
      Next
      If l0198 > gVBCodeSizeWords Then
        l0198 = 0
      ElseIf gVBCodeTokenWordsFinal(l0198) <> l01A2 Then
        If gVBCodeTokenWords(l0198) <= 0 Then l0198 = 0
      End If
    Loop
    l01A4 = g_TKStruct(p0196).M1CC7
  End If
  DisOutput ")"
  l0198 = g_TKStruct(p0196).TKFrmRes
  l0198 = NamesLookUpTbl(l0198 + mc008C)
  If l0198 Then sub07BA l0198, 0, False
  DisOutputWrite
  If Len(l01A4) > 1 And Len(l01A4) <> Len(l01A8) Then
    If Len(l01A4) > l019A + 1 Then
      DisOutput "' called ("
      GoSub L1CB28
      DisOutput ")"
      DisOutputWrite
    End If
  End If
Exit Sub

L1CB28:
  l019A = 0
  For l0198 = 2 To Len(l01A4)
    If l019A Then DisOutput ", "
    l019A = l019A + 1
    l01AE = Asc(Mid$(l01A4, l019A + 1, 1))
    l01AA = l01AE And mc0116
    If l01AE Then
      If l01AA = mc01B0 Then
        l01AE = l01AE And mc01B2
      ElseIf l01AE = mc0118 Then
        l01AE = l01AE And mc011A
      Else
        l01AE = l01AE And &H38
      End If
      If l01AE = 0 Then DisOutput "ByVal "
    End If
    DisOutput "p" & Format$(l019A)
    Select Case l01AA
    Case 0
    Case mc01B4 To mc01B6, mc01B0
      DisOutput Mid$(gc117C, l01AA, 1)
    Case mc01BA: DisOutput " As Variant"
    Case mc01BC: DisOutput " As object"
    Case mc01BE: DisOutput " As type"
    Case mc01C0: DisOutput " As array"
    Case Else
      DisOutput " As Any"
    End Select
  Next
Return
End Sub

Sub TKInit1()
24160
  
  mMaxPos = g_TKStruct(g_CurSub).mCodeSize
  
  If mMaxPos = 0 Then
    mMaxPos = gVBCodeSizeWords
  End If
  
  ScanedTokensCount = 0

End Sub

Function fn0E36(p01C6%) As String
24170
  fn0E36 = fn0DDA(p01C6)
End Function

Sub sub0A57(p01C8 As Integer)
Dim l01CA As Integer
Dim l01CC As Integer
Dim l01CE As Integer

24020
  l01CE = g_TKStruct(g_CurSub).M251F
  If mMaxPos = gVBCodeSizeWords Then
    Exit Sub
  End If
    l01CC = mMaxPos
    Do
      l01CA = gVBCodeTokenWordsExE(l01CC)
      If l01CA <= p01C8 Then
        If l01CA < p01C8 Then Exit Do
        If Asc(Mid$(gIntTokenStrBuff, l01CC - 1, 1)) = mc0078 Then Exit Sub
        sub0A82 l01CC, 1, mc0132, mc0078, 0
        Exit Sub
      End If
      l01CA = l01CC + 1
      l01CC = 0
      For l01CA = l01CA To gVBCodeSizeWords
        If gVBCodeTokenWordsFinal(l01CA) = l01CE Then l01CC = l01CA: Exit For
      Next
    Loop While l01CC
End Sub

Sub sub0C56(p01D0 As Integer, p01D2 As Integer)
Dim l01D4 As Integer
Dim l01D6 As Integer
Dim l01D8 As Integer
Dim l01DA As Integer
 Const mc01DC = 16 ' &H10%

24030
  l01D6 = Asc(Mid$(gIntTokenStrBuff, p01D0, 1)) And mask_1110_0000
  If l01D6 > p01D2 Then Exit Sub
  l01DA = gVBCodeTokenWordsFinal(p01D0)
  If g_CurSub Then
    l01D8 = g_TKStruct(g_CurSub).M251F
    If l01DA Then
      If l01DA <> l01D8 Then
        If p01D2 = mc013E Then
          l01D6 = Asc(Mid$(gIntTokenStrBuff, p01D0, 1)) And Not mask_1110_0000
          If l01D6 And mc01DC Then
            Exit Sub
          End If
          p01D2 = mc0142
          Mid$(gIntTokenStrBuff, p01D0, 1) = Chr$(l01D6 Or p01D2)
          gWriteOutPut = True
          Exit Sub
        ElseIf p01D2 <> mc0142 Then
          If gv0BC6 Then MakeErrBeep
        End If
      End If
    End If
  End If
  l01D4 = gVBCodeTokenWordsExE(p01D0)
  If p01D2 <= mc0146 Then
    If l01D4 = 0 Then
      p01D2 = mc013E
    ElseIf l01D4 < 0 Then
    ElseIf l01D4 And 1 Then
      MakeErrBeep
    ElseIf l01D4 And &H8001 Then
      MakeErrBeep
      If -l01D4 > TokenStruct1.Type_3 Then p01D2 = mc0132
    ElseIf l01D4 >= 6 And l01D4 < TokenStruct1.PubOrPrivate \ &H200 Then
      p01D2 = mc007A
    Else
      p01D2 = mc0146
    End If
  End If
  If l01DA Or g_CurSub = 0 Then
    Exit Sub
  End If
  Select Case p01D2
  Case mc007A, mc0132
    If p01D0 < mCurPos Then mCurPos = p01D0
    GoSub L1D2C0
  Case mc0144
  Case mc013E
    If l01D4 Then
      p01D2 = mc0142
    Else
      GoSub L1D2C0
    End If
  Case mc0142
  Case mc0146
    GoSub L1D2C0
  Case &HE0
  Case Else
    MakeErrBeep
  End Select
Exit Sub

L1D2C0:
  If p01D0 < mCurPos Then Return
  If l01D8 > m0040 Then
  Else
    gVBCodeTokenWordsFinal(p01D0) = l01D8
    If p01D0 < mMaxPos Then
      g_TKStruct(g_CurSub).mCodeSize = p01D0
      mMaxPos = p01D0
    End If
    gWriteOutPut = True
  End If
Return
End Sub
