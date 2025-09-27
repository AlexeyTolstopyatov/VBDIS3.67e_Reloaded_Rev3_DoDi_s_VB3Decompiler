Attribute VB_Name = "MODULE6"
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
Dim VarNameStack() As String
Dim ESP As Integer
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
  If p00D2 Then sub05D7 p00D2 Else PushOnNameStack ""
  OutputTxt "(" & PopArg() & ")"
End Sub

Sub sub07EA()
Dim l00D6 As Integer

10040
  l00D6 = gv0BC8
  gCodeIndent = 0
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
  
  DecompParamList g_CurSub
  
  If gv0BCE = 0 Then DecompDeclarations
  
  gv0BC8 = l00D6
  
End Sub

Sub OneParamCmds(pCmd As Integer, pParam%)
Dim l00E8 As Integer
Dim l00EA As String
Dim l00EC As Integer
Dim l00EE As Integer
 Const mc00F4 = 9 ' &H9%
 Const mc00FE = 12 ' &HC%
 Const mc0100 = 10 ' &HA%
 Const mc010C = 1 ' &H1%

'10,20,30,40,50....
10050
  On (pCmd \ &H10) + 1 GoTo L9D02, LA2B0, LA9AA, LAB80, LAD2C, LB07C, LB292, LB404, LB576, LBD76, LBE2E, LC0C8, LC0D2, LC0DE, LC0EA, LC37C, LC4FC, LC518, LC676, LC682, LC7D2, LC8E8

L9CEC:
  Error mc00CE
Exit Sub


L9D02: On pCmd And &HF GoTo L_0, L_1, L_2, L_3, L_4, L_5, L_6, L_7, L_8, L_9, L_A, L_Assign, L_B, L_C, L_D
' Set CodeIndent Level
    SetCodeIndent pParam: pParam = 0
    Exit Sub
L_0:
    SetCodeIndent IntBuff1
    Exit Sub
L_1:
L_2:
    GoTo L9CEC
L_3:
    OutputTxt fn06D5(IntBuff2) & " "
    Exit Sub
L_4:
    OutputTxt fn06D5(IntBuff2) & Space$(IntBuff3)
    Exit Sub
L_5:
    OutputTxt ": "
    Exit Sub
L_6:
    OutputTxt ": "
    sub06F3 IntBuff1
    Exit Sub
L_7:
L_8:
    GoTo L9CEC

L_A:
    pushVarNameByIdx IntBuff1, pParam: pParam = 0
      If tkMain.iToken2 = mc00F4 Then
        l00EC = UDiv(IntBuff1, 2)
        l00EE = gLocalVarsTypeArrExE(l00EC - 1)
        Select Case l00EE And &HC000
        Case 0, &H8000
          sub068D l00EE
        Case &H4000
          If l00EE = &H4006 Then
            sub068D l00EE
            l00EA = fn07C6(gLocalVarsTypeArrExE(l00EC))
          Else
            sub068D l00EE
          End If
        Case Else
          MakeErrBeep
        End Select
      End If
    Exit Sub

L_Assign:
    sub0705 IntBuff1, pParam, 0
    pParam = 0
    OutputTxt " = " & PopArg()
  Exit Sub
L_B:
    sub0705 IntBuff1, pParam, -1
    pParam = 0
      If tkMain.iToken2 = mc00F4 Then
        sub068D gLocalVarsTypeArrExE(UDiv(IntBuff1, 2) - 1)
      End If
  Exit Sub

L_C:
    l00E8 = IntBuff1
    If l00E8 >= 0 Then
        l00EE = UDiv(IntBuff2, 2)
        gv0BFC.Type_3 = l00EE
        Select Case tkMain.iToken2
        Case mc00FE
          sub0680 GetVarName(l00EE * 2), pParam, l00E8
          Exit Sub
        Case mc00F4
          sub068D gLocalVarsTypeArrExE(l00EE - 1)
          sub0710
        Case mc0100
          Select Case RawTokenForResume
          Case &H7E0 To &H816
            If l00E8 Then sub05D7 l00E8
          Case Else
  If l00EE <= Len(gLocalVarsTypeBuff) Then
            If (Asc(Mid$(gLocalVarsTypeBuff, l00EE, 1)) And &HF) = 9 Then
             sub068D gLocalVarsTypeArrExE(l00EE - 1)
            End If
  End If
          End Select
        Case Else
          MakeErrBeep
        End Select
        If l00E8 > 1 Then sub05C5 l00E8
      sub0680 GetVarName(IntBuff2), pParam, l00E8
    Else
      pushVarNameByIdx IntBuff2, pParam
      pParam = 0
    End If
    Exit Sub
L_D:
    sub0680 GetVarName(IntBuff2), pParam, IntBuff1
    sub071E PopArg()
    OutputTxt " = " & PopArg()
    Exit Sub
LA2B0: On pCmd And &HF GoTo LA3A6, LA592, LA698, LA74A, LA7D0, LA83C, LA8CA, LA8D6, LA8E4, LA8F2, LA946, LA952, LA992, LA998, LA99E
    sub0705 IntBuff2, pParam, 0
    sub0728 IntBuff1
    If pParam = 0 Then
      If m007A Then
        sub0705 0, 0, -1
      Else
        sub0705 0, 0, 0
      End If
    Else
      pParam = 0
    End If
    Exit Sub
LA36C:
    sub0705 IntBuff2, pParam, 0
    sub0728 IntBuff1
    m0078 = -1
    Exit Sub
LA3A6:
    Select Case pParam
    Case 0 To 7
      If tkMain.iToken2 = mc00F4 Then
        sub066C "." & fn07C6(IntBuff1)
      Else
        sub066C fn0753(IntBuff1, False, 0)
      End If
      If pParam Then
        sub066C Mid$("%&!#@?$", pParam, 1)
        pParam = 0
      End If
    Case 9
      If IntBuff1 < 0 Then
        sub066C "!" & fn07C6(IntBuff1)
      ElseIf IntBuff1 And 1 Then
        sub066C fn0753(IntBuff1, False, pParam)
      ElseIf VBver_11 < 3 Then
        sub066C "!" & fn07C6(IntBuff1)
      Else
        sub066C fn0760(IntBuff1)
      End If
      pParam = 0
    Case Else
      MakeErrBeep
    End Select
    Exit Sub
LA592:
    If pParam = 9 Then
      sub071E PopArg() & fn0753(IntBuff1, False, pParam)
      pParam = 0
    Else
      sub071E PopArg() & fn0753(IntBuff1, False, 0)
      If pParam Then
        OutputTxt Mid$("%&!#@?$", pParam, 1)
        pParam = 0
      End If
    End If
    OutputTxt " = " & PopArg()
    Exit Sub
LA698:
    l00EA = PopArg() & fn0753(IntBuff2, True, pParam = 9)
    If IntBuff1 >= 0 Then
      sub0680 l00EA, pParam, IntBuff1
    Else
      If pParam Then
        l00EA = l00EA & Mid$("%&!#@?$", pParam, 1)
        pParam = 0
      End If
      PushOnNameStack l00EA
    End If
    Exit Sub
LA74A:
    l00EA = PopArg() & fn0753(IntBuff2, True, pParam = 9)
    sub0680 l00EA, pParam, IntBuff1
    sub0656 PopArg() & " = ", 1
    Exit Sub
LA7D0:
    gv0BFC.Type_3 = IntBuff1
    gv0BFC.M2226 = IntBuff2
    If m007A Then MakeErrBeep
    m007A = pCmd
    If m0078 Then
      sub0705 m0078, 0, -1
    End If
    Exit Sub
LA83C:
    If m007A Then MakeErrBeep
    gv0BFC.Type_3 = IntBuff1
    If gv0BC8 And gv0BFC.Type_3 > 7 Then gv0BFC.Type_3 = 0
    gv0BFC.M2226 = IntBuff2
    m007A = pCmd
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
    sub0656 GetVBDISKeyWord(), 1
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
  On pCmd And &HF GoTo LA9DE, LA9E4, LA9F0, LAA16, LAAAA, LAAB0, LAABC, LAB0A, LAB10, LAB16, LAB1C, LAB22, LAB28, LAD60, LAB74
LA9DE:
LA9E4:
  GoTo L9CEC
LA9F0:  OutputTxt "Call "
LAA16:
      l00EA = getFrmNamebyOffbyIdx(NamesLookUpTbl((RotateRight3(IntBuff2)) + mc010C))
    If IntBuff1 And pCmd = &H23 Then
      sub0680 l00EA, 0, IntBuff1
      OutputTxt PopArg()
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
    sub0662 GetVBDISKeyWord()
    sub0656 "Case", 1
    Exit Sub

LAB74:
    GoTo L9CEC
LAB80: On pCmd And &HF GoTo LABB4, LABBA, LABC6, LABF2, LABFA, LAE0C, LAC04, LAC0A, LAC16, LAC4E, LAC54, LAC5A, LAC66, LACC4, LAC9E

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
LAD2C: On pCmd And &HF GoTo LAD78, LADBA, LADCC, LADC0, LAE12, LAE70, LAECE, LAEDA, LAE18, LAF3A, LAF40, LAF46, LAF52, LAF90, LAF9C

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
    sub071E GetVBDISKeyWord() & " " & PopArg() & " Then "
    Exit Sub
LAE70:
    sub071E "If " & PopArg() & " Then " & MakeLocalName(IntBuff1)
    Exit Sub
LAECE:
  GoTo L9CEC
LAEDA:
    sub071E "If " & PopArg() & " GoTo " & MakeLocalName(IntBuff1)
    Exit Sub

LAF3A:
LAF40:
LAF46:
  GoTo L9CEC
LAF52:
    sub0656 GetVBDISKeyWord(), 0
    If gv0BC8 Then ESP = ESP - 1
    Exit Sub
LAF90:
  GoTo L9CEC
LAF9C:
    sub071E GetVBDISKeyWord()
    Select Case IntBuff1
    Case -1
      OutputTxt "GoTo 0"
    Case -2
      OutputTxt "Resume Next"
    Case Else
      OutputTxt "GoTo " & MakeLocalName(IntBuff1)
    End Select
    Exit Sub
LB07C: On pCmd And &HF GoTo LB0AE, LB172, LB17E, LB1C0, LB1C6, LB1CC, LB1D2, LAD68, LB1DA, LB1E0, LB1EC, LB218, LB21E, LB224, LB22A
LB0AE:
    sub071E "On " & PopArg() & GetVBDISKeyWord()
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
    l00EA = PopArg()
    While m006C > AltToken_9Bit
      l00EA = ", " & l00EA
      m006C = m006C - 1
    Wend
    sub066C l00EA
    Exit Sub
LB292: On pCmd And &HF GoTo LB2C4, LB2CA, LB2D0, LB33A, LB35E, LB38C, LB3A4, LB3AA, LB3D4, LB3DA, LB3E0, LB3E6, LB3EC, LB3F2, LB3F8
LB2C4:
LB2CA:
LB2D0:
    If pCmd And 1 Then l00EA = "Step"
    If pCmd And 2 Then l00EA = "-" & l00EA
    sub0680 l00EA, 0, 2
    Exit Sub
LB33A:
    m0070 = "Debug."
    Exit Sub
LB35E:
    m0070 = PopArg() & "."
    Exit Sub
LB38C:
    m006C = &H60
    Exit Sub
LB3A4:
LB3AA:
    sub0656 GetVBDISKeyWord(), ESP
    Exit Sub
LB3D4:
LB3DA:
LB3E0:
LB3E6:
LB3EC:
LB3F2:
LB3F8:
  GoTo L9CEC
LB404: On pCmd And &HF GoTo LB436, LB44E, LB43C, LB442, LB454, LB4CA, LB500, LB50E, LB51C, LB522, LB52E, LB558, LB564, LB56A

LB436:
LB43C:
LB442:
  GoTo L9CEC
LB44E:
LB454:
    While ESP > 1
      sub066C PopArg()
    Wend
    sub0656 fn0740("Print"), ESP
    Exit Sub
L_9:
LB4CA:
    FrmStruct2 = m001C
    sub0734 IntBuff1, PopArg()
    Exit Sub
LB500:
    Exit Sub
LB50E:
    Exit Sub
LB51C:
LB522:
    GoTo L9CEC
LB52E:
    sub0656 GetVBDISKeyWord(), IntBuff1
    Exit Sub
LB558:
LB55E:
LB564:
LB56A:
    GoTo L9CEC
LB576: On pCmd And &HF GoTo LB5F0, LB5F6, LB5FC, LB61C, LB608, LB726, LB732, LB75E, LB83C, LB8A6, LB8AC, LB9B4, LB9C0, LBA42, LBA08
LB5A8:
    sub0662 ", "
    sub0656 GetVBDISKeyWord(), 2
    Exit Sub
LB5F0:
LB5F6:
LB5FC:
    GoTo L9CEC
LB608:
    l00EA = PopArg()
LB61C:
    If ESP > 1 Then sub066C PopArg()
    Select Case IntBuff1
    Case 0
    Case 1: l00EA = l00EA & ", B"
    Case 2: l00EA = l00EA & ", BF"
    Case Else: MakeErrBeep
    End Select
    If Len(l00EA) Then PushOnNameStack l00EA
    sub0656 fn0740(GetVBDISKeyWord()), ESP
    Exit Sub
LB726:
    GoTo L9CEC
LB732:
    sub0656 GetVBDISKeyWord(), ESP
    Exit Sub
LB75E:
LB764:
    Select Case IntBuff1 And &HFFFE
    Case &H8002
    Case &H4002
      l00EA = "To " & PopArg()
      VarNameStack(ESP) = l00EA
    Case &H2
      sub0677 "To"
    Case Else: MakeErrBeep
    End Select
    sub0656 GetVBDISKeyWord(), 2
    Exit Sub
LB83C:
LB842:
    sub071E GetVBDISKeyWord() & PopArg() & " = "
    OutputTxt PopArg()
    Exit Sub

LB8A6:
LB8AC:
    If pParam Then
      sub071E "Mid(" & PopArg() & ", "
      pParam = 0
    Else
      sub071E "Mid$(" & PopArg() & ", "
    End If
    l00EA = PopArg()
    If pCmd And 1 Then sub05C5 2
    OutputTxt PopArg() & ") = " & l00EA
    Exit Sub
LB9B4:
    GoTo L9CEC
LB9C0:
    sub0677 "As"
    sub0656 GetVBDISKeyWord(), 1
    Exit Sub
LBA08:
    sub066C " Len = " & PopArg()
LBA42:
    l00EA = "As " & PopArg()
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
    sub071E GetVBDISKeyWord() & PopArg() & " For " & l00EA
    Exit Sub
LBD76: On pCmd And &HF GoTo LBDA8, LBDAE, LBDB4, LB5A8, LBDBC, LBDC2, LBDC8, LBDD4, LBE06, LBE0C, LB842, LBE14, LBE1A, LB764, LBE22
LBDA8:
LBDAE:
LBDB4:

LBDBC:
LBDC2:
LBDC8:
    GoTo L9CEC
LBDD4:
    FrmStruct2 = m001C
    sub071E PopArg()
    Exit Sub
LBE06:
LBE0C:

LBE14:
LBE1A:

LBE22:
    GoTo L9CEC
LBE2E: On pCmd And &HF GoTo LBF4E, LBF5A, LBF76, LBF7C, LBF82, LBF88, LBF8E, LBF94, LBF9A, LBFA0, LBFAC, LC0AA, LC0B0, LC0B6, LC0BC
    l00EA = PopArg() & ")"
    l00EA = PopArg() & ", " & l00EA
    l00EA = PopArg() & ")-(" & l00EA
    l00EA = " (" & PopArg() & ", " & l00EA
    sub0656 fn0740(GetVBDISKeyWord()) & l00EA, 0
    Exit Sub
LBF4E:
    GoTo L9CEC
LBF5A:
    m0074 = GetVBDISKeyWord()
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
      If ESP > 0 Then
      If pParam And &H20 Then
        pParam = InStr(VarNameStack(ESP), "$")
        VarNameStack(ESP) = Left$(VarNameStack(ESP), pParam - 1) & Mid$(VarNameStack(ESP), pParam + 1)
      ElseIf pParam Then
        VarNameStack(ESP) = "C" & gVB_a_TypeConv(pParam) & "(" & VarNameStack(ESP) & ")"
      End If
      End If
      pParam = 0
      Exit Sub
LC0AA:
LC0B0:
LC0B6:
LC0BC:
  GoTo L9CEC
LC0C8: GoTo L9CEC
LC0D2: GoTo L9CEC

LC0DE: GoTo L9CEC

LC0EA: On pCmd And &HF GoTo LC11C, LC122, LC128, LC134, LC162, LC188, LC1C0, LC1F8, LC23E, LC276, LC2BC, LC2F4, LC364, LC370

LC11C:
LC122:
LC128:
    GoTo L9CEC
LC134:
    PushOnNameStack Format$(pParam): pParam = 0
    Exit Sub
LC162:
    PushOnNameStack Format$(IntBuff1)
    Exit Sub
LC188:
    PushOnNameStack Format$(m0032) & "&"
    Exit Sub
LC1C0:
    PushOnNameStack "&H" & Hex$(IntBuff1)
    Exit Sub
LC1F8:
    PushOnNameStack "&H" & Hex$(m0032) & "&"
    Exit Sub
LC23E:
    PushOnNameStack "&O" & Oct$(IntBuff1)
    Exit Sub
LC276:
    PushOnNameStack "&O" & Oct$(m0032) & "&"
    Exit Sub
LC2BC:
    PushOnNameStack Trim$(Str$(m0036)) & "!"
    Exit Sub
LC2F4:
    PushOnNameStack Trim$(Str$(m003A)) & "#"
    Exit Sub
LC32C:
    PushOnNameStack Trim$(Str$(m0042)) & "@"
    Exit Sub
LC364:
LC36A:
LC370:
  GoTo L9CEC
LC37C: On pCmd And &HF GoTo LC3B0, LC3B6, LC3BC, LC3C2, LC3C8, LC3CE, LC3D4, LC3E0, LC476, LC47C, LC488, LC4DE, LC4E4, LC4EA, LC4F0

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
LC4FC: If pCmd = &H107 Then GoTo LC32C
  GoTo L9CEC

LC518: On pCmd And &HF GoTo LC54C, LC552, LC558, LC55E, LC564, LC56A, LC570, LC576, LC57C, LC588, LC58E, LC614, LC620, LC664, LC66A

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
    If pCmd And 1 Then l00EA = fn0740(l00EA)
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
LC682: On pCmd And &HF GoTo LC6B4, LC6BA, LC6C6, LC720, LA36C, LC72E, LC796, LC79C, LC7A2, LC7A8, LC7AE, LC7B4, LC7C0, LC7C6

LC6B4:
LC6BA:
    GoTo L9CEC
LC6C6:
    sub0656 "Set", 1
    OutputTxt " = " & PopArg()
    Exit Sub
LC720:
    GoTo L9CEC

LC72E:
    If IntBuff1 Then l00EA = "Text" Else l00EA = "Binary"
    sub0656 GetVBDISKeyWord() & l00EA, 0
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
LC7D2: On pCmd And &HF GoTo LC806, LC80C, LC812, LA8F8, LC81A, LC826, LC87E, LC8B0, LC8B6, LC8BC, LC8C2, LC8C8, LC8CE, LC8D4, LC8DA

LC806:
LC80C:
LC812:

LC81A:
    GoTo L9CEC
LC826:
  sub068D IntBuff1
  l00EA = gVBObjsDefs(IntBuff1).Name_4
  PushOnNameStack "New " & l00EA
  Exit Sub
LC87E:
    PushOnNameStack GetVBDISKeyWord()
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

LC8E8: On pCmd And &HF GoTo LC91A, LC920, LC926, LC92C, LC932, LC938, LC93E, LC944, LC94A, LC950, LC956, LC95C, LC962, LC96E, LC9BA
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
  l00EA = PopArg() & "." & fn03DF(IntBuff2)
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
    
    pv0128 = gLocalVarsTypeArrExE(p012A - 1)
    
    p012A = Asc(Mid$(gLocalVarsTypeBuff, p012A, 1))
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

Sub CodeDecompile(BlockSize As Long, p144 As Integer)

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
  gv0BC8 = p144 And gc0BCC
  m001C = gv0BCE And gc0BD2
  FrmStruct2 = False
  ESP = 0
  ReDim VarNameStack(1 To 50)

On Error GoTo ERR_TK_Unknown
    hFile_TKexe = hVBFile
  gCodeIndent = -1
  
  BlockEnd = Seek(hFile_TKexe) + BlockSize
  
  
'111111
  Get hFile_TKexe, , RawTokenForResume
  While Seek(hFile_TKexe) < BlockEnd

Err_ResumeFromUnkown:
    RawToken = RawTokenForResume
      CurOffset = Seek(hFile_TKexe) - 2
      ConvertToken RawToken
        RawToken = tkMain.AltToken
        If tkMain.iToken = 0 Then
        End If
    
    '1111 1100 0000 0000   &HFC00
    '0000 0100 0000 0000   &H400 ; SHR 0A ->
    
    'Example: If RawToken is ...
    '0001 0000 0000 0000   &h1000  it get...
    '0000 0000 0000 0100   &h0004
    RawToken_ROR9 = (RawToken And &HFC00) \ &H400
    
    Select Case Cur_VBDat.TK_Case1_8bit And &HF
    ' 1,2
    Case mc015E, mc0160

LCF0A:
      
     'Param
      Select Case Cur_VBDat.Num_Param_0123
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
    
    '5 LineFinished
    Case mc0164
      FrmStruct2 = False
      
     'Error check stack empty and ??? also
      If ESP And gv0BD8 = 0 Then
        OutputTxt " ?"
        gv0BD8 = -2
      End If
      
      DisOutputWrite
      
      If ESP Then sub0629
      
      m0062 = 0
      
      FileOff_PrevToken = Loc(hFile_TKexe) - 2
      
      m0078 = 0: m007A = False
      
      If AltToken_9Bit Then GoTo LCF0A
    
    
    '8
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
      PushOnNameStack StringBuff
      
    ' AlignSeek
      l_Int_Buff1 = l_Int_Buff2 - l_Int_Buff1 - 4
      If l_Int_Buff1 > 0 Then
        Seek hFile_TKexe, Seek(hFile_TKexe) + l_Int_Buff1
      End If
    '10
    Case mc0168: Get hFile_TKexe, , m0032
    '11
    Case mc016A: Get hFile_TKexe, , m0036
    '12
    Case mc016C: Get hFile_TKexe, , m003A
    '14
    Case mc016E: Get hFile_TKexe, , m0042
    '13
    Case mc0170
      Get hFile_TKexe, , l_Int_Buff1: l_Int_Buff1 = l_Int_Buff1 \ 2
      m007C(0) = l_Int_Buff1
      For i = 1 To l_Int_Buff1
        Get hFile_TKexe, , m007C(i)
      Next i
      
    Case Else
      gv7B36 = gDisOutLogDbgLineCount
      ShowErrMsg mc0176
      gv0BD8 = AltToken_9Bit
      If Cur_VBDat.Num_Param_0123 Then
        Seek hFile_TKexe, Seek(hFile_TKexe) + Cur_VBDat.Num_Param_0123 * 2
      End If
    End Select
    
'222
    Get hFile_TKexe, , RawTokenForResume
    
    Select Case Cur_VBDat.mTokenFlags And &HF
    Case mc0178: OneParamCmds AltToken_9Bit, RawToken_ROR9
    
    Case mc017A
      If Cur_VBDat.mTokenFlags And mc017C Then m0064 = 0
      sub0656 GetVBDISKeyWord(), Cur_VBDat.Bit_Forward2
    
    Case mc017E
      sub0656 fn0740(GetVBDISKeyWord()), ESP
    
    Case mc0180: PushOnNameStack GetVBDISKeyWord()
    Case mc0182: sub0662 GetVBDISKeyWord()
    Case mc0184: sub066C GetVBDISKeyWord()
    Case mc0186: sub0677 GetVBDISKeyWord()
    Case mc0188: sub0680 GetVBDISKeyWord(), RawToken_ROR9, Cur_VBDat.Bit_Forward2
    Case mc018A
      sub0680 fn0740(GetVBDISKeyWord()), RawToken_ROR9, Cur_VBDat.Bit_Forward2
      
    Case mc018C
      If Cur_VBDat.mTokenFlags And mc018E Then RawToken_ROR9 = 0
      
    Case m0190:
      gv0BD8 = AltToken_9Bit
      
    Case Else
      MakeErrBeep
      
    End Select
    
    
    
    If Cur_VBDat.mTokenFlags And mc0192 Then
      If gv0BD8 = 0 Then gv0BD8 = AltToken_9Bit
    End If
    
    If Cur_VBDat.mTokenFlags And mc017C Then
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
  If ESP = 0 Then Exit Sub
  If hFile_Out Then
    Print #hFile_Out, "--- unused expressions (top down) ---"
    While ESP > 0
      Print #hFile_Out, , PopArg()
    Wend
  End If
End Sub

Sub DisOutputWrite()

10120
  If gCodeIndent < 0 Then gCodeIndent = 0: gDisOutputTxt = "": Exit Sub
  gDisOutputTxt = RTrim$(gDisOutputTxt)
  
  While ESP > 0
    gDisOutputTxt = gDisOutputTxt & " ?+" & PopArg() & "?"
  Wend
  
  If m001C And Len(gDisOutputTxt) = 0 Then Exit Sub
  
 'Out if file is open
  If hFile_Out Then
    Print #hFile_Out, gDisOutputTxt
  Else
 
 'Open outputfile
    On Error Resume Next
    Listbox_File_Out.AddItem gDisOutputTxt
    DoEvents
    If Err Then
      
      If Len(gDisOutputTxt) = 0 Then Exit Sub
      
      gDisOutLogDbgLineCount = gDisOutLogDbgLineCount + 1
      If gDisOutLogDbgLineCount >= UBound(gDisOutLogDbg) Then ReDim Preserve gDisOutLogDbg(gDisOutLogDbgLineCount + 1)
      
      gDisOutLogDbg(gDisOutLogDbgLineCount).OutputTxt_M1E72 = gDisOutputTxt
      gDisOutLogDbg(gDisOutLogDbgLineCount).Offs = FileOff_PrevToken
      
      If hFile_TKexe Then gDisOutLogDbg(gDisOutLogDbgLineCount + 1).Offs = Seek(hFile_TKexe)
    
    End If
  End If
  
  
  Debug.Print Hex(Seek(hVBFile) - 1), gDisOutputTxt
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
  PushOnNameStack GetVBDISKeyWord() & l01C4
End Sub

Sub sub0677(ByVal p01CA As String)
Dim l01CC As String

10150
  l01CC = PopArg()
  If ESP < 1 Then
    ESP = 1
    VarNameStack(ESP) = "<missed arg>"
  End If
  VarNameStack(ESP) = VarNameStack(ESP) & " " & p01CA & " " & l01CC
End Sub

Sub sub06C8()
10160
  ReDim gDisOutLogDbg(10), VarNameStack(50)
  gv0BC8 = True
  gDisOutLogDbgLineCount = 0
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
Dim IntBuff As Integer

10180
    OldPos = Seek(hFile_TKexe)
    
    Get hFile_TKexe, curSegOffs + Uint32(Offset) + 4, IntBuff
    Select Case IntBuff
    Case -1
        MakeLocalName = "L" & Hex$(Uint32(Offset) + curSegOffs - 1)
    Case Else
      MakeLocalName = Format$(Uint32(IntBuff))
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

Function PopArg() As String
10200
  If ESP <= 0 Then
    PopArg = "<missed arg>"
  Else
    PopArg = VarNameStack(ESP)
    VarNameStack(ESP) = ""
    ESP = ESP - 1
  End If
End Function

Sub sub066C(p01E2 As String)
10210
  If ESP <= 0 Then PushOnNameStack "<missed arg>"
  VarNameStack(ESP) = VarNameStack(ESP) & p01E2
End Sub

Sub sub0662(ByVal p01E4 As String)
10220
  If ESP > 0 Then VarNameStack(ESP) = p01E4 & VarNameStack(ESP)
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

Sub PushOnNameStack(VarName As String)
10250
On Error GoTo 501
  ESP = ESP + 1
  VarNameStack(ESP) = VarName
Exit Sub

501 If ESP < 0 Then
      ESP = 1
    Else
10258
      ReDim Preserve VarNameStack(1 To UBound(VarNameStack) + 10)
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
    l0202 = PopArg()
  End If
  PushOnNameStack p01FC & "(" & l0202 & ")"
End Sub

Private Sub pushVarNameByIdx(VarNameIdx%, VarType As Integer)
10270
  PushOnNameStack GetVarName(VarNameIdx)
  If VarType Then VarNameStack(ESP) = VarNameStack(ESP) & Mid$("%&!#@?$", VarType, 1)
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
    sub071E PopArg()
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
Dim VBModuleDef As VBModuleDef_struc
Dim l0234 As String

10330
  Get hVBFile, VBAbsOffs(g_TKStruct(p022E).TKFrmRes), VBModuleDef
  If VBModuleDef.SubOrFunc And 1 Then
    DisOut = "Sub"
  Else
    DisOut = "Function"
  End If
  
  If VBModuleDef.PubOrPrivate <> &HC Then MakeErrBeep
  
  l0234 = getFrmNamebyOffbyIdx(p022E)
  
  gv0C4A.M225B = fn03D3(VBModuleDef.M1587)
  gv0C4A.M2262 = fn03D3(VBModuleDef.M159D)
  DisOutput "Declare " & DisOut & " " & l0234
  
  If Right$(gv0C4A.M225B, 1) = "." Then
    gv0C4A.M225B = Left$(gv0C4A.M225B, Len(gv0C4A.M225B) - 1)
  End If
  DisOutput " Lib """ & gv0C4A.M225B & """"
  
  If gv0C4A.M2262 <> l0234 Then
    DisOutput " Alias """ & gv0C4A.M2262 & """"
  End If
  
  DecompParamList p022E
  
End Sub

Sub DeclareVarNameAndType(ByVal TokenIndex As Integer, pVBCodeTokenWordsExE() As Integer, TokenString As String, IsGlobalDef%)
Dim l024E As Integer
Dim TokenByte As Integer
 Const mc0252 = 224 ' &HE0%
 Const mc0254 = 192 ' &HC0%
 Const mc0256 = 16 ' &H10%
 Const mc0258 = 17 ' &H11%

10340
  TokenByte = Asc(Mid$(TokenString, TokenIndex, 1))
  If (TokenByte And mc0252) = mc0254 Then sub076F TokenIndex, TokenString, IsGlobalDef: Exit Sub
  
  If IsGlobalDef Then
    DisOutput "Global " & fn0CDA(TokenIndex)
  Else
    DisOutput fn0DDA(TokenIndex)
  End If
  
  If TokenByte And mc0256 Then
    
    l024E = Asc(Mid$(TokenString, TokenIndex + 1, 1))
    Select Case l024E
    Case mc0258
      sub07AD pVBCodeTokenWordsExE(), TokenIndex + 3
    Case Else
      MakeErrBeep
    End Select
    
  End If
  
  DetermDeclareType TokenByte, pVBCodeTokenWordsExE(TokenIndex - 1), True
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
      
      l0266 = Asc(Mid$(gLocalVarsTypeBuff, l0268, 1))
      
      If (l0266 And &H1F) = 0 Or (l0266 And mc0252) = mc026C Then
        ProcessType p0262, tkMain.iToken1, tkMain.iToken2
      End If
      
      GetVarName = fn0DDA(l0268)
End Function

Sub sub05C5(p026E As Integer)
Dim l0270 As String
Dim l0272 As String

10000
  If (p026E And &H7FFF) = 0 Then Exit Sub
  If Abs(p026E) > ESP Then
    PushOnNameStack "<concat?>"
    Exit Sub
  End If
  If p026E < 0 Then
    p026E = -p026E
    Do While p026E
      l0270 = PopArg(): p026E = p026E - 1
      If Left$(l0270, 8) <> "<default" Then Exit Do
    Loop
    While p026E
      l0272 = PopArg(): p026E = p026E - 1
      If Left$(l0272, 8) = "<default" Then
        l0270 = ", " & l0270
      Else
        l0270 = l0272 & ", " & l0270
      End If
    Wend
  Else
    l0270 = PopArg(): p026E = p026E - 1
    While p026E > 0
      l0270 = PopArg() & ", " & l0270
      p026E = p026E - 1
    Wend
  End If
  PushOnNameStack l0270
  p026E = 1
End Sub

Sub sub05D7(p0276 As Integer)
Dim l0278 As String

10010
  If m001C Or (p0276 And 1) Then sub05C5 p0276: Exit Sub
  While p0276 > 0
    l0278 = PopArg() & l0278
    If VarNameStack(ESP) = "" Then
      ESP = ESP - 1
    Else
      l0278 = PopArg() & " To " & l0278
    End If
    p0276 = p0276 - 2
    If p0276 Then l0278 = ", " & l0278
  Wend
  PushOnNameStack l0278
  p0276 = 1
End Sub

Sub sub05E6()
10020
  ShowErrMsg mc022A
End Sub
