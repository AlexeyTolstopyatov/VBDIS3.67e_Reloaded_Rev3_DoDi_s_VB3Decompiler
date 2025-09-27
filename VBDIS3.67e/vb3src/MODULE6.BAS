'Attribute VB_Name = "MODULE6"
' Module6
Option Explicit
Dim m001C As Integer
Dim FrmStruct2 As Integer
Dim hFile_TKexe As Integer
Dim m0032 As Long
Dim m0036 As Single
Dim m003A As Double
Dim m0042 As Currency
Dim m0062 As Integer
Dim m0064 As Integer
Dim IntBuff1  As Integer
Dim IntBuff2  As Integer
Dim IntBuff3 As Integer
Dim m006C As Integer
Dim RawTokenForResume As Integer
Dim m0070 As String
Dim m0074 As String
Dim m0078 As Integer
Dim m007A As Integer
Dim m007C(255) As Integer
Dim VarNameTbl() As String
Dim m00C4 As Integer
Const mc00CE = 1000 ' &H3E8%
Const mc00F4 = 9 ' &H9%
Const mc00FE = 12 ' &HC%
Const mc0100 = 10 ' &HA%
Const mc010C = 1 ' &H1%
Const mc0130 = 10 ' &HA%
Const mc0134 = 8 ' &H8%
Const mc013E = 9 ' &H9%
Const mc0140 = 8192 ' &H2000%
Const mc015E = 1 ' &H1%
Const mc0160 = 2 ' &H2%
Const mc0162 = 4 ' &H4%
Const mc0164 = 5 ' &H5%
Const mc0166 = 8 ' &H8%
Const mc0168 = 10 ' &HA%
Const mc016A = 11 ' &HB%
Const mc016C = 12 ' &HC%
Const mc016E = 14 ' &HE%
Const mc0170 = 13 ' &HD%
Const mc0176 = 256 ' &H100%
Const mc0178 = 1 ' &H1%
Const mc017A = 3 ' &H3%
Const mc017C = 1024 ' &H400%
Const mc017E = 9 ' &H9%
Const mc0180 = 4 ' &H4%
Const mc0182 = 5 ' &H5%
Const mc0184 = 6 ' &H6%
Const mc0186 = 7 ' &H7%
Const mc0188 = 8 ' &H8%
Const mc018A = 10 ' &HA%
Const mc018C = 2 ' &H2%
Const mc018E = 512 ' &H200%
Const mc0192 = 256 ' &H100%
Const mc0198 = 32 ' &H20%
Const mc022A = 32 ' &H20%
Const mc0252 = 224 ' &HE0%
Const mc0254 = 192 ' &HC0%
Const mc0256 = 16 ' &H10%
Const mc0258 = 17 ' &H11%
Const mc026C = 32 ' &H20%

Sub sub0728(p00D2 As Integer)
10030
  If p00D2 Then sub05D7 p00D2 Else StoreInVarNameTbl ""
  OutputTxt "(" & fn074C() & ")"
End Sub

Sub sub07EA()
Dim l00D6 As Integer

10040
  l00D6 = gv0BC8
  gv0C5E = 0
  FileOff_PrevToken = 0
  gv0BC8 = True
  
  If TokenStruct1.SubOrFunc And 1 Then
    DisOut = "Sub"
  Else
    DisOut = "Function"
  End If
  
  If ScanOnly Then
  Else
    If TokenStruct1.PubOrPrivate And &H40 Then DisOutput "Private "
    
    DisOutput DisOut & " " & getFrmNamebyOffbyIdx(g_CurSub)
    
  End If
  
  sub0792 g_CurSub
  
  If gv0BCE = 0 Then sub07F8
  
  gv0BC8 = l00D6
  
End Sub

Sub sub0649(p00E4 As Integer, p00E6%)
Dim l00E8 As Integer
Dim l00EA As String
Dim l00EC As Integer
Dim l00EE As Integer
 Const mc00F4 = 9 ' &H9%
 Const mc00FE = 12 ' &HC%
 Const mc0100 = 10 ' &HA%
 Const mc010C = 1 ' &H1%

10050
  On (p00E4 \ &H10) + 1 GoTo L9D02, LA2B0, LA9AA, LAB80, LAD2C, LB07C, LB292, LB404, LB576, LBD76, LBE2E, LC0C8, LC0D2, LC0DE, LC0EA, LC37C, LC4FC, LC518, LC676, LC682, LC7D2, LC8E8

L9CEC:
  Error mc00CE
Exit Sub

L9D02: On p00E4 And &HF GoTo L9D4E, L9D66, L9D6C, L9D78, L9DB4, L9DF0, L9E1A, L9E50, L9E56, LB4C4, L9E64, L9F7A, L9FD2, LA03E, LA234
    sub06E9 p00E6: p00E6 = 0
    Exit Sub
L9D4E:
    sub06E9 IntBuff1
    Exit Sub
L9D66:
L9D6C:
    GoTo L9CEC
L9D78:
    OutputTxt fn06D5(IntBuff2) & " "
    Exit Sub
L9DB4:
    OutputTxt fn06D5(IntBuff2) & Space$(IntBuff3)
    Exit Sub
L9DF0:
    OutputTxt ": "
    Exit Sub
L9E1A:
    OutputTxt ": "
    sub06F3 IntBuff1
    Exit Sub
L9E50:
L9E56:
    GoTo L9CEC

L9E64:
    sub06F9 IntBuff1, p00E6: p00E6 = 0
      If tkMain.iToken2 = mc00F4 Then
        l00EC = UDiv(IntBuff1, 2)
        l00EE = gVBCodeTokenWordsExE(l00EC - 1)
        Select Case l00EE And &HC000
        Case 0, &H8000
          sub068D l00EE
        Case &H4000
          If l00EE = &H4006 Then
            sub068D l00EE
            l00EA = fn07C6(gVBCodeTokenWordsExE(l00EC))
          Else
            sub068D l00EE
          End If
        Case Else
          MakeErrBeep
        End Select
      End If
    Exit Sub

L9F7A:
    sub0705 IntBuff1, p00E6, 0
    p00E6 = 0
    OutputTxt " = " & fn074C()
  Exit Sub
L9FD2:
    sub0705 IntBuff1, p00E6, -1
    p00E6 = 0
      If tkMain.iToken2 = mc00F4 Then
        sub068D gVBCodeTokenWordsExE(UDiv(IntBuff1, 2) - 1)
      End If
  Exit Sub

LA03E:
    l00E8 = IntBuff1
    If l00E8 >= 0 Then
        l00EE = UDiv(IntBuff2, 2)
        gv0BFC.Type_3 = l00EE
        Select Case tkMain.iToken2
        Case mc00FE
          sub0680 GetVarName(l00EE * 2), p00E6, l00E8
          Exit Sub
        Case mc00F4
          sub068D gVBCodeTokenWordsExE(l00EE - 1)
          sub0710
        Case mc0100
          Select Case RawTokenForResume
          Case &H7E0 To &H816
            If l00E8 Then sub05D7 l00E8
          Case Else
  If l00EE <= Len(gIntTokenStrBuff) Then
            If (Asc(Mid$(gIntTokenStrBuff, l00EE, 1)) And &HF) = 9 Then
             sub068D gVBCodeTokenWordsExE(l00EE - 1)
            End If
  End If
          End Select
        Case Else
          MakeErrBeep
        End Select
        If l00E8 > 1 Then sub05C5 l00E8
      sub0680 GetVarName(IntBuff2), p00E6, l00E8
    Else
      sub06F9 IntBuff2, p00E6
      p00E6 = 0
    End If
    Exit Sub
LA234:
    sub0680 GetVarName(IntBuff2), p00E6, IntBuff1
    sub071E fn074C()
    OutputTxt " = " & fn074C()
    Exit Sub
LA2B0: On p00E4 And &HF GoTo LA3A6, LA592, LA698, LA74A, LA7D0, LA83C, LA8CA, LA8D6, LA8E4, LA8F2, LA946, LA952, LA992, LA998, LA99E
    sub0705 IntBuff2, p00E6, 0
    sub0728 IntBuff1
    If p00E6 = 0 Then
      If m007A Then
        sub0705 0, 0, -1
      Else
        sub0705 0, 0, 0
      End If
    Else
      p00E6 = 0
    End If
    Exit Sub
LA36C:
    sub0705 IntBuff2, p00E6, 0
    sub0728 IntBuff1
    m0078 = -1
    Exit Sub
LA3A6:
    Select Case p00E6
    Case 0 To 7
      If tkMain.iToken2 = mc00F4 Then
        sub066C "." & fn07C6(IntBuff1)
      Else
        sub066C fn0753(IntBuff1, False, 0)
      End If
      If p00E6 Then
        sub066C Mid$("%&!#@?$", p00E6, 1)
        p00E6 = 0
      End If
    Case 9
      If IntBuff1 < 0 Then
        sub066C "!" & fn07C6(IntBuff1)
      ElseIf IntBuff1 And 1 Then
        sub066C fn0753(IntBuff1, False, p00E6)
      ElseIf VBver_11 < 3 Then
        sub066C "!" & fn07C6(IntBuff1)
      Else
        sub066C fn0760(IntBuff1)
      End If
      p00E6 = 0
    Case Else
      MakeErrBeep
    End Select
    Exit Sub
LA592:
    If p00E6 = 9 Then
      sub071E fn074C() & fn0753(IntBuff1, False, p00E6)
      p00E6 = 0
    Else
      sub071E fn074C() & fn0753(IntBuff1, False, 0)
      If p00E6 Then
        OutputTxt Mid$("%&!#@?$", p00E6, 1)
        p00E6 = 0
      End If
    End If
    OutputTxt " = " & fn074C()
    Exit Sub
LA698:
    l00EA = fn074C() & fn0753(IntBuff2, True, p00E6 = 9)
    If IntBuff1 >= 0 Then
      sub0680 l00EA, p00E6, IntBuff1
    Else
      If p00E6 Then
        l00EA = l00EA & Mid$("%&!#@?$", p00E6, 1)
        p00E6 = 0
      End If
      StoreInVarNameTbl l00EA
    End If
    Exit Sub
LA74A:
    l00EA = fn074C() & fn0753(IntBuff2, True, p00E6 = 9)
    sub0680 l00EA, p00E6, IntBuff1
    sub0656 fn074C() & " = ", 1
    Exit Sub
LA7D0:
    gv0BFC.Type_3 = IntBuff1
    gv0BFC.M2226 = IntBuff2
    If m007A Then MakeErrBeep
    m007A = p00E4
    If m0078 Then
      sub0705 m0078, 0, -1
    End If
    Exit Sub
LA83C:
    If m007A Then MakeErrBeep
    gv0BFC.Type_3 = IntBuff1
    If gv0BC8 And gv0BFC.Type_3 > 7 Then gv0BFC.Type_3 = 0
    gv0BFC.M2226 = IntBuff2
    m007A = p00E4
    If m0078 Then
      sub0705 m0078, 0, -1
    End If
    Exit Sub
LA8CA:
  GoTo L9CEC
LA8D6:
  GoTo L9CEC
LA8E4:
  GoTo L9CEC
LA8F2:
LA8F8:
    sub0656 fn0BA0(), 1
    If m007A Then
      sub0705 0, 0, -1
    End If
    Exit Sub
LA946:
    GoTo L9CEC
LA952:
    OutputTxt "Type " & GetVarName(IntBuff2)
    Exit Sub
LA992:
LA998:
LA99E:
    GoTo L9CEC
LA9AA:
  On p00E4 And &HF GoTo LA9DE, LA9E4, LA9F0, LAA16, LAAAA, LAAB0, LAABC, LAB0A, LAB10, LAB16, LAB1C, LAB22, LAB28, LAD60, LAB74
LA9DE:
LA9E4:
  GoTo L9CEC
LA9F0:  OutputTxt "Call "
LAA16:
      l00EA = getFrmNamebyOffbyIdx(NamesLookUpTbl((RotateRight3(IntBuff2)) + mc010C))
    If IntBuff1 And p00E4 = &H23 Then
      sub0680 l00EA, 0, IntBuff1
      OutputTxt fn074C()
    Else
      sub0656 l00EA, IntBuff1
    End If
    Exit Sub
LAAAA:
LAAB0:
    GoTo L9CEC
LAABC:
    sub0677 "To"
    sub0656 "Case", 1
    Exit Sub
LAB0A:
LAB10:
LAB16:
LAB1C:
LAB22:
LAB28:
    sub0662 fn0BA0()
    sub0656 "Case", 1
    Exit Sub

LAB74:
    GoTo L9CEC
LAB80: On p00E4 And &HF GoTo LABB4, LABBA, LABC6, LABF2, LABFA, LAE0C, LAC04, LAC0A, LAC16, LAC4E, LAC54, LAC5A, LAC66, LACC4, LAC9E

LABB4:
LABBA:
  GoTo L9CEC
LABC6:
  sub071E MakeLocalName(IntBuff1)
  Exit Sub
LABF2:
LABFA:

LAC04:
LAC0A:
  GoTo L9CEC
LAC16:
  OutputTxt "End " & DisOut
  Exit Sub
LAC4E:
LAC54:
LAC5A:
    GoTo L9CEC
LAC66:
  sub071E "Exit " & DisOut
  Exit Sub
LAC9E:
    sub0677 "Step"
LACC4:
    sub0677 "To"
    sub0677 "="
    sub0656 "For ", 1
    Exit Sub
LAD2C: On p00E4 And &HF GoTo LAD78, LADBA, LADCC, LADC0, LAE12, LAE70, LAECE, LAEDA, LAE18, LAF3A, LAF40, LAF46, LAF52, LAF90, LAF9C

LAD60:
LAD68:
    Exit Sub

LAD78:
  sub071E "GoSub " & MakeLocalName(IntBuff1)
  Exit Sub
LADBA:
LADC0:
  GoTo L9CEC
LADCC:
    sub071E "GoTo " & MakeLocalName(IntBuff1)
    Exit Sub
LAE0C:
LAE12:
LAE18:
    sub071E fn0BA0() & " " & fn074C() & " Then "
    Exit Sub
LAE70:
    sub071E "If " & fn074C() & " Then " & MakeLocalName(IntBuff1)
    Exit Sub
LAECE:
  GoTo L9CEC
LAEDA:
    sub071E "If " & fn074C() & " GoTo " & MakeLocalName(IntBuff1)
    Exit Sub

LAF3A:
LAF40:
LAF46:
  GoTo L9CEC
LAF52:
    sub0656 fn0BA0(), 0
    If gv0BC8 Then m00C4 = m00C4 - 1
    Exit Sub
LAF90:
  GoTo L9CEC
LAF9C:
    sub071E fn0BA0()
    Select Case IntBuff1
    Case -1
      OutputTxt "GoTo 0"
    Case -2
      OutputTxt "Resume Next"
    Case Else
      OutputTxt "GoTo " & MakeLocalName(IntBuff1)
    End Select
    Exit Sub
LB07C: On p00E4 And &HF GoTo LB0AE, LB172, LB17E, LB1C0, LB1C6, LB1CC, LB1D2, LAD68, LB1DA, LB1E0, LB1EC, LB218, LB21E, LB224, LB22A
LB0AE:
    sub071E "On " & fn074C() & fn0BA0()
    OutputTxt MakeLocalName(m007C(1))
    For l00EC = 2 To m007C(0)
      OutputTxt ", " & MakeLocalName(m007C(l00EC))
    Next l00EC
    Exit Sub
LB172:
    GoTo L9CEC
LB17E:
    sub071E "Resume " & MakeLocalName(IntBuff1)
    Exit Sub
LB1C0:
LB1C6:
LB1CC:
LB1D2:

LB1DA:
LB1E0:
  GoTo L9CEC
LB1EC:
    sub066C ", "
    Exit Sub
LB218:
LB21E:
LB224:
LB22A:
    l00EA = fn074C()
    While m006C > AltToken_9Bit
      l00EA = ", " & l00EA
      m006C = m006C - 1
    Wend
    sub066C l00EA
    Exit Sub
LB292: On p00E4 And &HF GoTo LB2C4, LB2CA, LB2D0, LB33A, LB35E, LB38C, LB3A4, LB3AA, LB3D4, LB3DA, LB3E0, LB3E6, LB3EC, LB3F2, LB3F8
LB2C4:
LB2CA:
LB2D0:
    If p00E4 And 1 Then l00EA = "Step"
    If p00E4 And 2 Then l00EA = "-" & l00EA
    sub0680 l00EA, 0, 2
    Exit Sub
LB33A:
    m0070 = "Debug."
    Exit Sub
LB35E:
    m0070 = fn074C() & "."
    Exit Sub
LB38C:
    m006C = &H60
    Exit Sub
LB3A4:
LB3AA:
    sub0656 fn0BA0(), m00C4
    Exit Sub
LB3D4:
LB3DA:
LB3E0:
LB3E6:
LB3EC:
LB3F2:
LB3F8:
  GoTo L9CEC
LB404: On p00E4 And &HF GoTo LB436, LB44E, LB43C, LB442, LB454, LB4CA, LB500, LB50E, LB51C, LB522, LB52E, LB558, LB564, LB56A

LB436:
LB43C:
LB442:
  GoTo L9CEC
LB44E:
LB454:
    While m00C4 > 1
      sub066C fn074C()
    Wend
    sub0656 fn0740("Print"), m00C4
    Exit Sub
LB4C4:
LB4CA:
    FrmStruct2 = m001C
    sub0734 IntBuff1, fn074C()
    Exit Sub
LB500:
    Exit Sub
LB50E:
    Exit Sub
LB51C:
LB522:
    GoTo L9CEC
LB52E:
    sub0656 fn0BA0(), IntBuff1
    Exit Sub
LB558:
LB55E:
LB564:
LB56A:
    GoTo L9CEC
LB576: On p00E4 And &HF GoTo LB5F0, LB5F6, LB5FC, LB61C, LB608, LB726, LB732, LB75E, LB83C, LB8A6, LB8AC, LB9B4, LB9C0, LBA42, LBA08
LB5A8:
    sub0662 ", "
    sub0656 fn0BA0(), 2
    Exit Sub
LB5F0:
LB5F6:
LB5FC:
    GoTo L9CEC
LB608:
    l00EA = fn074C()
LB61C:
    If m00C4 > 1 Then sub066C fn074C()
    Select Case IntBuff1
    Case 0
    Case 1: l00EA = l00EA & ", B"
    Case 2: l00EA = l00EA & ", BF"
    Case Else: MakeErrBeep
    End Select
    If Len(l00EA) Then StoreInVarNameTbl l00EA
    sub0656 fn0740(fn0BA0()), m00C4
    Exit Sub
LB726:
    GoTo L9CEC
LB732:
    sub0656 fn0BA0(), m00C4
    Exit Sub
LB75E:
LB764:
    Select Case IntBuff1 And &HFFFE
    Case &H8002
    Case &H4002
      l00EA = "To " & fn074C()
      VarNameTbl(m00C4) = l00EA
    Case &H2
      sub0677 "To"
    Case Else: MakeErrBeep
    End Select
    sub0656 fn0BA0(), 2
    Exit Sub
LB83C:
LB842:
    sub071E fn0BA0() & fn074C() & " = "
    OutputTxt fn074C()
    Exit Sub

LB8A6:
LB8AC:
    If p00E6 Then
      sub071E "Mid(" & fn074C() & ", "
      p00E6 = 0
    Else
      sub071E "Mid$(" & fn074C() & ", "
    End If
    l00EA = fn074C()
    If p00E4 And 1 Then sub05C5 2
    OutputTxt fn074C() & ") = " & l00EA
    Exit Sub
LB9B4:
    GoTo L9CEC
LB9C0:
    sub0677 "As"
    sub0656 fn0BA0(), 1
    Exit Sub
LBA08:
    sub066C " Len = " & fn074C()
LBA42:
    l00EA = "As " & fn074C()
    l00EE = IntBuff1
    If l00EE And &H8CD0 Then MakeErrBeep
    Select Case l00EE And &H7000
    Case &H1000: l00EA = "Lock Read Write " & l00EA
    Case &H2000: l00EA = "Lock Write " & l00EA
    Case &H3000: l00EA = "Lock Read " & l00EA
    Case &H4000: l00EA = "Shared " & l00EA
    End Select
    If l00EE And &H300 Then
      If l00EE And &H200 Then l00EA = "Write " & l00EA
      If l00EE And &H100 Then l00EA = "Read " & l00EA
      l00EA = "Access " & l00EA
    End If
    Select Case l00EE And &H3F
    Case 1: l00EA = "Input " & l00EA
    Case 2: l00EA = "Output " & l00EA
    Case 4: l00EA = "Random " & l00EA
    Case 8: l00EA = "Append " & l00EA
    Case &H20: l00EA = "Binary " & l00EA
    End Select
    sub071E fn0BA0() & fn074C() & " For " & l00EA
    Exit Sub
LBD76: On p00E4 And &HF GoTo LBDA8, LBDAE, LBDB4, LB5A8, LBDBC, LBDC2, LBDC8, LBDD4, LBE06, LBE0C, LB842, LBE14, LBE1A, LB764, LBE22
LBDA8:
LBDAE:
LBDB4:

LBDBC:
LBDC2:
LBDC8:
    GoTo L9CEC
LBDD4:
    FrmStruct2 = m001C
    sub071E fn074C()
    Exit Sub
LBE06:
LBE0C:

LBE14:
LBE1A:

LBE22:
    GoTo L9CEC
LBE2E: On p00E4 And &HF GoTo LBF4E, LBF5A, LBF76, LBF7C, LBF82, LBF88, LBF8E, LBF94, LBF9A, LBFA0, LBFAC, LC0AA, LC0B0, LC0B6, LC0BC
    l00EA = fn074C() & ")"
    l00EA = fn074C() & ", " & l00EA
    l00EA = fn074C() & ")-(" & l00EA
    l00EA = " (" & fn074C() & ", " & l00EA
    sub0656 fn0740(fn0BA0()) & l00EA, 0
    Exit Sub
LBF4E:
    GoTo L9CEC
LBF5A:
    m0074 = fn0BA0()
    Exit Sub
LBF76:
LBF7C:
LBF82:
LBF88:
LBF8E:
LBF94:
LBF9A:
LBFA0:
  GoTo L9CEC
LBFAC:
      If m00C4 > 0 Then
      If p00E6 And &H20 Then
        p00E6 = InStr(VarNameTbl(m00C4), "$")
        VarNameTbl(m00C4) = Left$(VarNameTbl(m00C4), p00E6 - 1) & Mid$(VarNameTbl(m00C4), p00E6 + 1)
      ElseIf p00E6 Then
        VarNameTbl(m00C4) = "C" & gVB_a_TypeConv(p00E6) & "(" & VarNameTbl(m00C4) & ")"
      End If
      End If
      p00E6 = 0
      Exit Sub
LC0AA:
LC0B0:
LC0B6:
LC0BC:
  GoTo L9CEC
LC0C8: GoTo L9CEC
LC0D2: GoTo L9CEC

LC0DE: GoTo L9CEC

LC0EA: On p00E4 And &HF GoTo LC11C, LC122, LC128, LC134, LC162, LC188, LC1C0, LC1F8, LC23E, LC276, LC2BC, LC2F4, LC364, LC370

LC11C:
LC122:
LC128:
    GoTo L9CEC
LC134:
    StoreInVarNameTbl Format$(p00E6): p00E6 = 0
    Exit Sub
LC162:
    StoreInVarNameTbl Format$(IntBuff1)
    Exit Sub
LC188:
    StoreInVarNameTbl Format$(m0032) & "&"
    Exit Sub
LC1C0:
    StoreInVarNameTbl "&H" & Hex$(IntBuff1)
    Exit Sub
LC1F8:
    StoreInVarNameTbl "&H" & Hex$(m0032) & "&"
    Exit Sub
LC23E:
    StoreInVarNameTbl "&O" & Oct$(IntBuff1)
    Exit Sub
LC276:
    StoreInVarNameTbl "&O" & Oct$(m0032) & "&"
    Exit Sub
LC2BC:
    StoreInVarNameTbl Trim$(Str$(m0036)) & "!"
    Exit Sub
LC2F4:
    StoreInVarNameTbl Trim$(Str$(m003A)) & "#"
    Exit Sub
LC32C:
    StoreInVarNameTbl Trim$(Str$(m0042)) & "@"
    Exit Sub
LC364:
LC36A:
LC370:
  GoTo L9CEC
LC37C: On p00E4 And &HF GoTo LC3B0, LC3B6, LC3BC, LC3C2, LC3C8, LC3CE, LC3D4, LC3E0, LC476, LC47C, LC488, LC4DE, LC4E4, LC4EA, LC4F0

LC3B0:
LC3B6:
LC3BC:
LC3C2:
LC3C8:
LC3CE:
LC3D4:
    GoTo L9CEC
LC3E0:
    sub0662 "TypeOf "
    sub066C " Is "
      l00EA = gVBObjsDefs(IntBuff1).Name_4
      If Len(l00EA) = 0 Then l00EA = Hex$(IntBuff1): MakeErrBeep
    sub066C l00EA
    Exit Sub
LC476:
LC47C:
    GoTo L9CEC
LC488:
    gv0BFC.Type_3 = -IntBuff2
    gv0BFC.M2226 = IntBuff3
    m007A = True
    If m0078 Then sub0705 m0078, 0, -1
    Exit Sub
LC4DE:
LC4E4:
LC4EA:
LC4F0:
    GoTo L9CEC
LC4FC: If p00E4 = &H107 Then GoTo LC32C
  GoTo L9CEC

LC518: On p00E4 And &HF GoTo LC54C, LC552, LC558, LC55E, LC564, LC56A, LC570, LC576, LC57C, LC588, LC58E, LC614, LC620, LC664, LC66A

LC54C:
LC552:
LC558:
LC55E:
LC564:
LC56A:
LC570:
LC576:
LC57C:
    GoTo L9CEC
LC588:
LC58E:
    l00EC = IntBuff1
    l00EA = fn0BBC(l00EC)
    If p00E4 And 1 Then l00EA = fn0740(l00EA)
    If l00EC >= 0 Then
      sub0656 l00EA, IntBuff2
    Else
      sub0680 l00EA, 0, IntBuff2
    End If
    Exit Sub
LC614:
    GoTo L9CEC
LC620:
    If IntBuff1 <> &HFC Then GoTo L9CEC
    sub0656 "Dim Shared ", 0
    Exit Sub
LC664:
LC66A:
  GoTo L9CEC
LC676:
  GoTo L9CEC
LC682: On p00E4 And &HF GoTo LC6B4, LC6BA, LC6C6, LC720, LA36C, LC72E, LC796, LC79C, LC7A2, LC7A8, LC7AE, LC7B4, LC7C0, LC7C6

LC6B4:
LC6BA:
    GoTo L9CEC
LC6C6:
    sub0656 "Set", 1
    OutputTxt " = " & fn074C()
    Exit Sub
LC720:
    GoTo L9CEC

LC72E:
    If IntBuff1 Then l00EA = "Text" Else l00EA = "Binary"
    sub0656 fn0BA0() & l00EA, 0
    Exit Sub
LC796:
LC79C:
LC7A2:
LC7A8:
LC7AE:
LC7B4:
LC7BA:
LC7C0:
LC7C6:
  GoTo L9CEC
LC7D2: On p00E4 And &HF GoTo LC806, LC80C, LC812, LA8F8, LC81A, LC826, LC87E, LC8B0, LC8B6, LC8BC, LC8C2, LC8C8, LC8CE, LC8D4, LC8DA

LC806:
LC80C:
LC812:

LC81A:
    GoTo L9CEC
LC826:
  sub068D IntBuff1
  l00EA = gVBObjsDefs(IntBuff1).Name_4
  StoreInVarNameTbl "New " & l00EA
  Exit Sub
LC87E:
    StoreInVarNameTbl fn0BA0()
    sub068D &H4006
    Exit Sub
LC8B0:
LC8B6:
LC8BC:
LC8C2:
LC8C8:
LC8CE:
LC8D4:
LC8DA:
  GoTo L9CEC

LC8E8: On p00E4 And &HF GoTo LC91A, LC920, LC926, LC92C, LC932, LC938, LC93E, LC944, LC94A, LC950, LC956, LC95C, LC962, LC96E, LC9BA
LC91A:
LC920:
LC926:
LC92C:
LC932:
LC938:
LC93E:
LC944:
LC94A:
LC950:
LC956:
LC95C:
LC962:
  GoTo L9CEC
LC96E:
  l00EA = fn074C() & "." & fn03DF(IntBuff2)
  sub0656 l00EA, IntBuff1
Exit Sub
LC9BA:
  GoTo L9CEC
End Sub

Private Sub sub0705(ByVal pv0128 As Integer, p012A As Integer, p012C As Integer)
Dim l012E As Integer
 Const mc0130 = 10 ' &HA%
 Const mc0134 = 8 ' &H8%
 Const mc013E = 9 ' &H9%
 Const mc0140 = 8192 ' &H2000%

10060
  If p012A = 0 And p012C = -1 And m007A Then
    p012A = gv0BFC.Type_3
    p012C = gv0BFC.M2226
    l012E = True
    m007A = False
  End If
  If p012C = -1 Then
    p012C = 0
  ElseIf p012C < -1 Then
    l012E = True
  End If
  If p012A Then m0078 = 0
  If pv0128 Then
    If m0062 Then OutputTxt ", " Else m0062 = True
    OutputTxt GetVarName(pv0128)
    If Not l012E And p012A > 0 And p012A < 8 And p012A <> 6 Then
      OutputTxt Mid$("%&!#@?$", p012A, 1)
      Exit Sub
    End If
  End If
  If p012A = 0 Then Exit Sub
  If p012C < 0 Then
    sub0734 (p012C And &H7F) - 1, " As New "
  Else
    sub0734 p012C - 1, " As "
  End If
  If p012A < 0 Then
    OutputTxt "String *" & Str$(-p012A)
    Exit Sub
  End If
    pv0128 = gVBCodeTokenWordsExE(p012A - 1)
    p012A = Asc(Mid$(gIntTokenStrBuff, p012A, 1))
    Select Case p012A And &HF
    Case 0
      MakeErrBeep
    Case mc0130
        OutputTxt fn0CDA(UDiv(pv0128, 2))
    Case mc0134
      OutputTxt "String *" & Str$(Uint32(pv0128))
    Case Is < 8
      OutputTxt gVB_a_DataTypes(p012A And &HF)
    Case mc013E
      If gVBObjsDefs(pv0128 And &HFF).id_5 < 0 Then
        ShowErrMsg mc0140
        OutputTxt "()"
      End If
      OutputTxt gVBObjsDefs(pv0128 And &HFF).Name_4
    Case Else
      MakeErrBeep
      OutputTxt "?" & Hex$(p012A)
    End Select
End Sub

Sub sub05F4(BlockSize As Long, p0144 As Integer)

Dim StringBuff As String
Dim BlockEnd As Long
Dim i%
Dim l_Int_Buff1 As Integer
Dim RawToken As Integer
Dim l_Int_Buff2  As Integer
Dim RawToken_ROR9 As Integer
 Const mc015E = 1 ' &H1%
 Const mc0160 = 2 ' &H2%
 Const mc0162 = 4 ' &H4%
 Const mc0164 = 5 ' &H5%
 Const mc0166 = 8 ' &H8%
 Const mc0168 = 10 ' &HA%
 Const mc016A = 11 ' &HB%
 Const mc016C = 12 ' &HC%
 Const mc016E = 14 ' &HE%
 Const mc0170 = 13 ' &HD%
 Const mc0176 = 256 ' &H100%
 Const mc0178 = 1 ' &H1%
 Const mc017A = 3 ' &H3%
 Const mc017C = 1024 ' &H400%
 Const mc017E = 9 ' &H9%
 Const mc0180 = 4 ' &H4%
 Const mc0182 = 5 ' &H5%
 Const mc0184 = 6 ' &H6%
 Const mc0186 = 7 ' &H7%
 Const mc0188 = 8 ' &H8%
 Const mc018A = 10 ' &HA%
 Const mc018C = 2 ' &H2%
 Const mc018E = 512 ' &H200%
Static m0190 As Integer
 Const mc0192 = 256 ' &H100%
 Const mc0198 = 32 ' &H20%

10070
  gv0BC8 = p0144 And gc0BCC
  m001C = gv0BCE And gc0BD2
  FrmStruct2 = False
  m00C4 = 0
  ReDim VarNameTbl(1 To 50)

On Error GoTo ERR_TK_Unknown
    hFile_TKexe = hVBFile
  gv0C5E = -1
  
  BlockEnd = Seek(hFile_TKexe) + BlockSize
  
  Get hFile_TKexe, , RawTokenForResume
  While Seek(hFile_TKexe) < BlockEnd

Err_ResumeFromUnkown:
    RawToken = RawTokenForResume
      CurOffset = Seek(hFile_TKexe) - 2
      ConvertToken RawToken
        RawToken = tkMain.AltToken
        If tkMain.iToken = 0 Then
        End If
    
    RawToken_ROR9 = (RawToken And &HFC00) \ &H400
    
    Select Case ATK9B_VBDat_1.TK_Case1_8bit And &HF
    ' 1,2
    Case mc015E, mc0160

LCF0A:
      
      Select Case ATK9B_VBDat_1.ForwardWords
      Case 0
      Case 1
        Get hFile_TKexe, , IntBuff1
      Case 2
        Get hFile_TKexe, , IntBuff1
        Get hFile_TKexe, , IntBuff2
      Case 3
        Get hFile_TKexe, , IntBuff1
        Get hFile_TKexe, , IntBuff2
        Get hFile_TKexe, , IntBuff3
      Case Else
        MakeErrBeep
      End Select
    
    '4
    Case mc0162
      DisOutputWrite
      hFile_TKexe = 0
      Exit Sub
    '5
    Case mc0164
      FrmStruct2 = False
      
      If m00C4 And gv0BD8 = 0 Then
        OutputTxt " ?"
        gv0BD8 = -2
      End If
      
      DisOutputWrite
      
      If m00C4 Then sub0629
      
      m0062 = 0
      
      FileOff_PrevToken = Loc(hFile_TKexe) - 2
      
      m0078 = 0: m007A = False
      
      If AltToken_9Bit Then GoTo LCF0A
    
    Case mc0166
    ' String Found
      Get hFile_TKexe, , l_Int_Buff2
      Get hFile_TKexe, , l_Int_Buff1
      If l_Int_Buff2 > 2 Then
        Get hFile_TKexe, , l_Int_Buff1
      Else
        l_Int_Buff1 = 0
      End If
      
    ' Read String
      StringBuff = Space$(l_Int_Buff1)
      If l_Int_Buff1 Then Get hFile_TKexe, , StringBuff
      MakeQuotedString StringBuff
      StoreInVarNameTbl StringBuff
      
      l_Int_Buff1 = l_Int_Buff2 - l_Int_Buff1 - 4
      If l_Int_Buff1 > 0 Then
        Seek hFile_TKexe, Seek(hFile_TKexe) + l_Int_Buff1
      End If
    
    Case mc0168: Get hFile_TKexe, , m0032
    Case mc016A: Get hFile_TKexe, , m0036
    Case mc016C: Get hFile_TKexe, , m003A
    Case mc016E: Get hFile_TKexe, , m0042
    Case mc0170
      Get hFile_TKexe, , l_Int_Buff1: l_Int_Buff1 = l_Int_Buff1 \ 2
      m007C(0) = l_Int_Buff1
      For i = 1 To l_Int_Buff1
        Get hFile_TKexe, , m007C(i)
      Next i
      
    Case Else
      gv7B36 = gv0980
      ShowErrMsg mc0176
      gv0BD8 = AltToken_9Bit
      If ATK9B_VBDat_1.ForwardWords Then
        Seek hFile_TKexe, Seek(hFile_TKexe) + ATK9B_VBDat_1.ForwardWords * 2
      End If
    End Select
    
    Get hFile_TKexe, , RawTokenForResume
    
    Select Case ATK9B_VBDat_1.mTokenFlags And &HF
    Case mc0178: sub0649 AltToken_9Bit, RawToken_ROR9
    Case mc017A
      If ATK9B_VBDat_1.mTokenFlags And mc017C Then m0064 = 0
      sub0656 fn0BA0(), ATK9B_VBDat_1.M2CF2
    Case mc017E
      sub0656 fn0740(fn0BA0()), m00C4
    Case mc0180: StoreInVarNameTbl fn0BA0()
    Case mc0182: sub0662 fn0BA0()
    Case mc0184: sub066C fn0BA0()
    Case mc0186: sub0677 fn0BA0()
    Case mc0188: sub0680 fn0BA0(), RawToken_ROR9, ATK9B_VBDat_1.M2CF2
    Case mc018A
      sub0680 fn0740(fn0BA0()), RawToken_ROR9, ATK9B_VBDat_1.M2CF2
      
    Case mc018C
      If ATK9B_VBDat_1.mTokenFlags And mc018E Then RawToken_ROR9 = 0
      
    Case m0190:
      gv0BD8 = AltToken_9Bit
      
    Case Else
      MakeErrBeep
      
    End Select
    
    If ATK9B_VBDat_1.mTokenFlags And mc0192 Then
      If gv0BD8 = 0 Then gv0BD8 = AltToken_9Bit
    End If
    
    If ATK9B_VBDat_1.mTokenFlags And mc017C Then
      m0064 = 0: m0062 = 0: FrmStruct2 = False
      sub068D 0
    End If
    
  Wend
  hFile_TKexe = 0
Exit Sub

ERR_TK_Unknown:
  MakeErrBeep
  OutputTxt " ?"
  DisOutputWrite
  
  If hFile_Out Then
    Print #hFile_Out, "' --- unknown "; Hex$(AltToken_9Bit); " @ "; Hex$(FileOff_PrevToken); ":"; Hex$(Loc(hFile_TKexe) - 2); " ---"
  End If
  
    While (ScanTokens() And mc0198) = 0: Wend
    RawTokenForResume = intTokens
    ScanedTokensCount = 0
    
  On Error GoTo ERR_TK_Unknown
  If Seek(hFile_TKexe) < BlockEnd Then Resume Err_ResumeFromUnkown
  hFile_TKexe = 0
  Exit Sub
End Sub

Sub sub0698()
End Sub

Sub sub06A6()
End Sub

Sub sub06B3()
End Sub

Sub sub0629()
10110
  If m00C4 = 0 Then Exit Sub
  If hFile_Out Then
    Print #hFile_Out, "--- unused expressions (top down) ---"
    While m00C4 > 0
      Print #hFile_Out, , fn074C()
    Wend
  End If
End Sub

Sub DisOutputWrite()

10120
  If gv0C5E < 0 Then gv0C5E = 0: gDisOutputTxt = "": Exit Sub
  gDisOutputTxt = RTrim$(gDisOutputTxt)
  While m00C4 > 0
    gDisOutputTxt = gDisOutputTxt & " ?+" & fn074C() & "?"
  Wend
  If m001C And Len(gDisOutputTxt) = 0 Then Exit Sub
  If hFile_Out Then
    Print #hFile_Out, gDisOutputTxt
  Else
    On Error Resume Next
    Listbox_File_Out.AddItem gDisOutputTxt
    If Err Then
      If Len(gDisOutputTxt) = 0 Then Exit Sub
      gv0980 = gv0980 + 1
      If gv0980 >= UBound(gv094E) Then ReDim Preserve gv094E(gv0980 + 1)
      gv094E(gv0980).M1E72 = gDisOutputTxt
      gv094E(gv0980).Offs = FileOff_PrevToken
      If hFile_TKexe Then gv094E(gv0980 + 1).Offs = Seek(hFile_TKexe)
    End If
  End If
  gDisOutputTxt = ""
End Sub

Sub MakeQuotedString(p01BC As String)
Dim l01BE As Integer

10130
  If Len(p01BC) Then
    If InStr(p01BC, Chr$(&HD)) Then MakeErrBeep
    l01BE = -1
    Do
      l01BE = InStr(l01BE + 2, p01BC, """")
      If l01BE = 0 Then Exit Do
      p01BC = Left$(p01BC, l01BE) & Mid$(p01BC, l01BE)
    Loop
  End If
  p01BC = """" & p01BC & """"
End Sub

Sub sub06C0(p01C0 As Integer)
Dim l01C2 As Integer
Dim l01C4 As String

10140
  l01C4 = Space$(p01C0)
  If p01C0 > 0 Then
    Get hFile_TKexe, , l01C4
    If p01C0 And 1 Then Get hFile_TKexe, , FixedByteBuff
    l01C2 = InStr(l01C4, Chr$(&HD))
    While l01C2
      l01C4 = Left$(l01C4, l01C2 - 1) & String$(Asc(Mid$(l01C4, l01C2 + 1, 1)), Mid$(l01C4, l01C2 + 2, 1)) & Mid$(l01C4, l01C2 + 3)
      l01C2 = InStr(l01C4, Chr$(&HD))
    Wend
  End If
  StoreInVarNameTbl fn0BA0() & l01C4
End Sub

Sub sub0677(ByVal p01CA As String)
Dim l01CC As String

10150
  l01CC = fn074C()
  If m00C4 < 1 Then
    m00C4 = 1
    VarNameTbl(m00C4) = "<missed arg>"
  End If
  VarNameTbl(m00C4) = VarNameTbl(m00C4) & " " & p01CA & " " & l01CC
End Sub

Sub sub06C8()
10160
  ReDim gv094E(10), VarNameTbl(50)
  gv0BC8 = True
  gv0980 = 0
End Sub

Private Function fn06D5(p01D0 As Integer) As String
10170
    Select Case p01D0
    Case -6 To -1
        fn06D5 = "L" & Hex$(FileOff_PrevToken) & ":"
    Case Else
      fn06D5 = Format$(Uint32(p01D0))
    End Select
End Function

Private Function MakeLocalName(Offset As Integer) As String
Dim OldPos As Long
Dim intBuff As Integer

10180
    OldPos = Seek(hFile_TKexe)
    
    Get hFile_TKexe, curSegOffs + Uint32(Offset) + 4, intBuff
    Select Case intBuff
    Case -1
        MakeLocalName = "L" & Hex$(Uint32(Offset) + curSegOffs - 1)
    Case Else
      MakeLocalName = Format$(Uint32(intBuff))
    End Select
    
    Seek hFile_TKexe, OldPos
    
End Function

Function fn0740(p01DE As String) As String
10190
  If m0074 = "" Then
    fn0740 = m0070 & p01DE
  Else
    fn0740 = m0070 & m0074
    m0074 = ""
  End If
  m0070 = ""
End Function

Function fn074C() As String
10200
  If m00C4 <= 0 Then
    fn074C = "<missed arg>"
  Else
    fn074C = VarNameTbl(m00C4)
    VarNameTbl(m00C4) = ""
    m00C4 = m00C4 - 1
  End If
End Function

Sub sub066C(p01E2 As String)
10210
  If m00C4 <= 0 Then StoreInVarNameTbl "<missed arg>"
  VarNameTbl(m00C4) = VarNameTbl(m00C4) & p01E2
End Sub

Sub sub0662(ByVal p01E4 As String)
10220
  If m00C4 > 0 Then VarNameTbl(m00C4) = p01E4 & VarNameTbl(m00C4)
End Sub

Private Function fn0753(p01E8%, p01EA As Integer, p01EC As Integer) As String
Dim l01EE As String

10230
  If gv0BC8 Then
    If tkMain.iToken2 = mc00F4 Then
      l01EE = fn07C6(p01E8)
      If p01EA Then sub0710
    Else
      l01EE = fn0CDA(UDiv(p01E8, 2))
    End If
  Else
    If p01E8 < 0 Then
      If p01E8 <> -2 Then MakeErrBeep
      Exit Function
    End If
    l01EE = GetVarName(p01E8)
  End If
  If Len(l01EE) Then
    If p01EC Then
      fn0753 = "!" & l01EE
    Else
      fn0753 = "." & l01EE
    End If
  End If
End Function

Function fn0760(p01F2 As Integer) As String

10240
  If p01F2 > 0 And Len(gv0B9E) Then
    fn0760 = "!" & fn03DF(p01F2)
    sub0710
  Else
    fn0760 = "!?" & IntToHexStr(p01F2)
  End If
End Function

Sub StoreInVarNameTbl(VarName As String)
10250
On Error GoTo 501
  m00C4 = m00C4 + 1
  VarNameTbl(m00C4) = VarName
Exit Sub

501 If m00C4 < 0 Then
      m00C4 = 1
    Else
10258
      ReDim Preserve VarNameTbl(1 To UBound(VarNameTbl) + 10)
    End If
10259
    Resume
End Sub

Sub sub0680(p01FC As String, p01FE As Integer, ByVal pv0200 As Integer)
Dim l0202 As String

10260
  If p01FE Then
    p01FC = p01FC & Mid$("%&!#@?$", p01FE, 1)
    p01FE = 0
  End If
  If pv0200 Then
    sub05C5 pv0200
    l0202 = fn074C()
  End If
  StoreInVarNameTbl p01FC & "(" & l0202 & ")"
End Sub

Private Sub sub06F9(p0204%, p0206 As Integer)
10270
  StoreInVarNameTbl GetVarName(p0204)
  If p0206 Then VarNameTbl(m00C4) = VarNameTbl(m00C4) & Mid$("%&!#@?$", p0206, 1)
End Sub

Sub sub0656(p0208, ByVal pv020A As Integer)
10280
  If m0064 Then
    OutputTxt ", "
  Else
    m0064 = Len(gDisOutputTxt) + 1
    sub071E p0208
  End If
  If pv020A Then
    sub05C5 pv020A
    sub071E fn074C()
  End If
End Sub

Sub sub071E(ByVal p020C As String)
10290
  If FrmStruct2 Then Exit Sub
  If Len(gDisOutputTxt) Then
    If Right$(gDisOutputTxt, 1) <> " " Then
      gDisOutputTxt = gDisOutputTxt & " " & p020C
      Exit Sub
    End If
  End If
  gDisOutputTxt = gDisOutputTxt & p020C
End Sub

Sub OutputTxt(p020E As String)
10300
  If FrmStruct2 Then Exit Sub
  gDisOutputTxt = gDisOutputTxt & p020E
End Sub

Sub sub0734(p0210%, ByVal p0212$)
10310
  If FrmStruct2 Then Exit Sub
  sub06F3 p0210
  OutputTxt p0212
End Sub

Sub sub076F(ByVal pv0214 As Integer, p0216 As String, p0218 As Integer)
Dim l021A As Long
Dim l021C As Long
Dim l021E As Integer
Dim l0220 As String
 Const mc022A = 32 ' &H20%

10320
  l021E = Asc(Mid$(p0216, pv0214, 1)) And &HF
  If l021E >= mc0134 Then Exit Sub
  If p0218 Then
    l021A = gv0B8E + 2 * CLng(pv0214)
    DisOutput "Global Const " & fn0CDA(pv0214) & " ="
  Else
    l021A = gCurModStructOffset + 2 * CLng(pv0214)
    DisOutput "Const " & fn0DDA(pv0214) & " ="
  End If
  l021C = Seek(hVBFile)
  Seek hVBFile, l021A
  Select Case l021E
  Case 0:
    Get hVBFile, , l021E
    DisOutput Str$(l021E) & " ' ? &H" & fn02AC(l021E)
    Do
      pv0214 = pv0214 + 1
      If Asc(Mid$(p0216, pv0214, 1)) Then Exit Do
      Get hVBFile, , l021E
      DisOutput fn02AC(l021E)
    Loop
  Case 1: Get hVBFile, , l021E: DisOutput Str$(l021E) & " ' &H" & Hex$(l021E) & "%"
  Case 2: Get hVBFile, , m0032: DisOutput Str$(m0032) & "& ' &H" & Hex$(m0032) & "&"
  Case 3: Get hVBFile, , m0036: DisOutput Str$(m0036) & "!"
  Case 4:
    Get hVBFile, , m003A: l0220 = Str$(m003A)
    If InStr(l0220, ".") = 0 Then l0220 = l0220 & "#"
    DisOutput l0220
  Case 5: Get hVBFile, , m0042: DisOutput Str$(m0042) & "@"
  Case 6: ShowErrMsg mc022A
  Case 7
    Get hVBFile, , l021E
    Get hVBFile, gv0B8A + l021E, l021E
    Get hVBFile, gv0B8A + l021E, l021E
    l0220 = Space$(l021E)
    If l021E > 0 Then Get hVBFile, , l0220
    MakeQuotedString l0220
    DisOutput " " & l0220
  Case mc0134: MakeErrBeep
  Case mc013E
    Get hVBFile, l021A - 2, l021E
    DisOutput " As " & gVBObjsDefs(l021E And &HFF).Name_4
  Case mc0130: MakeErrBeep
  Case Else
    MakeErrBeep
    Do
      Get hVBFile, , l021E: DisOutput fn02AC(l021E)
      pv0214 = pv0214 + 1
    Loop Until Asc(Mid$(p0216, pv0214, 1))
  End Select
  Seek hVBFile, l021C
End Sub

Sub sub0785(p022E As Integer)
Dim l0232 As T14EB
Dim l0234 As String

10330
  Get hVBFile, VBAbsOffs(g_TKStruct(p022E).TKFrmRes), l0232
  If l0232.SubOrFunc And 1 Then
    DisOut = "Sub"
  Else
    DisOut = "Function"
  End If
  If l0232.PubOrPrivate <> &HC Then MakeErrBeep
  l0234 = getFrmNamebyOffbyIdx(p022E)
  gv0C4A.M225B = fn03D3(l0232.M1587)
  gv0C4A.M2262 = fn03D3(l0232.M159D)
  DisOutput "Declare " & DisOut & " " & l0234
  If Right$(gv0C4A.M225B, 1) = "." Then
    gv0C4A.M225B = Left$(gv0C4A.M225B, Len(gv0C4A.M225B) - 1)
  End If
  DisOutput " Lib """ & gv0C4A.M225B & """"
  If gv0C4A.M2262 <> l0234 Then
    DisOutput " Alias """ & gv0C4A.M2262 & """"
  End If
  sub0792 p022E
End Sub

Sub sub07A1(ByVal pv0242 As Integer, p0244() As Integer, p024A As String, p024C%)
Dim l024E As Integer
Dim l0250 As Integer
 Const mc0252 = 224 ' &HE0%
 Const mc0254 = 192 ' &HC0%
 Const mc0256 = 16 ' &H10%
 Const mc0258 = 17 ' &H11%

10340
  l0250 = Asc(Mid$(p024A, pv0242, 1))
  If (l0250 And mc0252) = mc0254 Then sub076F pv0242, p024A, p024C: Exit Sub
  If p024C Then
    DisOutput "Global " & fn0CDA(pv0242)
  Else
    DisOutput fn0DDA(pv0242)
  End If
  If l0250 And mc0256 Then
    l024E = Asc(Mid$(p024A, pv0242 + 1, 1))
    Select Case l024E
    Case mc0258
      sub07AD p0244(), pv0242 + 3
    Case Else
      MakeErrBeep
    End Select
  End If
  sub07BA l0250, p0244(pv0242 - 1), True
End Sub

Function fn07C6(p025C%) As String
10350
  fn07C6 = fn0B6F(p025C)
End Function

Function GetVarName(p0262%) As String
Dim l0264 As Long
Dim l0266 As Integer
Dim l0268 As Integer
 Const mc026C = 32 ' &H20%

10360
  l0264 = Uint32(p0262)
    If p0262 = 0 Then Beep: Exit Function
    If l0264 > gRawVBCodeSize Then Beep: Exit Function
      
      l0268 = l0264 \ 2
      
      l0266 = Asc(Mid$(gIntTokenStrBuff, l0268, 1))
      
      If (l0266 And &H1F) = 0 Or (l0266 And mc0252) = mc026C Then
        sub07DC p0262, tkMain.iToken1, tkMain.iToken2
      End If
      GetVarName = fn0DDA(l0268)
End Function

Sub sub05C5(p026E As Integer)
Dim l0270 As String
Dim l0272 As String

10000
  If (p026E And &H7FFF) = 0 Then Exit Sub
  If Abs(p026E) > m00C4 Then
    StoreInVarNameTbl "<concat?>"
    Exit Sub
  End If
  If p026E < 0 Then
    p026E = -p026E
    Do While p026E
      l0270 = fn074C(): p026E = p026E - 1
      If Left$(l0270, 8) <> "<default" Then Exit Do
    Loop
    While p026E
      l0272 = fn074C(): p026E = p026E - 1
      If Left$(l0272, 8) = "<default" Then
        l0270 = ", " & l0270
      Else
        l0270 = l0272 & ", " & l0270
      End If
    Wend
  Else
    l0270 = fn074C(): p026E = p026E - 1
    While p026E > 0
      l0270 = fn074C() & ", " & l0270
      p026E = p026E - 1
    Wend
  End If
  StoreInVarNameTbl l0270
  p026E = 1
End Sub

Sub sub05D7(p0276 As Integer)
Dim l0278 As String

10010
  If m001C Or (p0276 And 1) Then sub05C5 p0276: Exit Sub
  While p0276 > 0
    l0278 = fn074C() & l0278
    If VarNameTbl(m00C4) = "" Then
      m00C4 = m00C4 - 1
    Else
      l0278 = fn074C() & " To " & l0278
    End If
    p0276 = p0276 - 2
    If p0276 Then l0278 = ", " & l0278
  Wend
  StoreInVarNameTbl l0278
  p0276 = 1
End Sub

Sub sub05E6()
10020
  ShowErrMsg mc022A
End Sub
