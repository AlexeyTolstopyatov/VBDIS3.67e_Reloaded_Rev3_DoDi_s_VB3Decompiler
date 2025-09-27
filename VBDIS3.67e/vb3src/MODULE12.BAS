'Attribute VB_Name = "MODULE12"
' Module12
Option Explicit
Const VBX_VER3 = 3 ' &H3%
Const VBX_VER4 = 4 ' &H4%
Const mc003E = 17 ' &H11%
Const mc0040 = 18 ' &H12%
Const mc0042 = 19 ' &H13%
Const mc0044 = 20 ' &H14%
Const mc0046 = 21 ' &H15%
Const mc004C = 31 ' &H1F%
Dim m004E As Integer
Dim m008E As Long
Dim m0092 As Long
Dim m00A0 As T2C56
Dim m00A8(3) As String


Sub sub0B3E(ByVal p00BE%)
17000
  m00A8(0) = "<no/std obj>"
  m00A8(3) = "<ctrl>"
  m00A8(2) = "<class>"
  m00A8(1) = "<prop>"
  sub0B31 p00BE
End Sub

Sub sub0B4B()
End Sub

Function fn0B57(p00C8 As Integer, p00CA As Long, p00CC%) As String
Dim l00D0 As String * 256
Dim l00D2 As Integer

17040
  Get p00C8, p00CA + Uint32(p00CC), l00D0
  l00D2 = InStr(l00D0, Chr$(0))
  If l00D2 Then
    fn0B57 = Left$(l00D0, l00D2 - 1)
  Else
    MsgBox gc038A
    fn0B57 = l00D0
  End If
End Function

Function fn0B63(p00DA As Integer) As String
17050
  fn0B63 = gv12BE(gv1444(p00DA).Type_3 And gc129E)
End Function

Sub sub068D(ByVal pv00DC As Integer)
17060
  Select Case pv00DC And &HC000
  Case 0, &H8000
17061
    pv00DC = pv00DC And &HFF
    m00A0.Type_3 = 2
    m00A0.M1D80 = pv00DC
    m00A0.Name_4 = gVBObjsDefs(pv00DC).Name_4
  Case &H4000
17062
    If pv00DC = &H4006 Then
      m00A0.Type_3 = 2
      m00A0.M1D80 = curFrmStruct.M_2_curr
      m00A0.Name_4 = "me"
    Else
17063
      m00A0.Type_3 = 2
      pv00DC = pv00DC And &HFF
      m00A0.M1D80 = pv00DC
      m00A0.Name_4 = gVBObjsDefs(pv00DC).Name_4
    End If
  Case Else
17064
    Beep
    m00A0.Type_3 = 0
  End Select
End Sub

Function fn0B6F(p00EC As Integer) As String
Dim l00EE As Integer
Dim l00F0 As Integer
Dim l00F2 As String

17070
  l00EE = m00A0.M1D80
  l00F0 = p00EC And &H3FFF
  Select Case p00EC
  Case -2
17071
    If l00EE Then
      If m00A0.Type_3 = 2 Then
        l00EE = gVBObjsDefs(l00EE).id_add_5
      Else
        l00EE = 0
      End If
    End If
    If l00EE Then
      m00A0.Name_4 = m00A0.Name_4 & "<." & gVBObjsDefs(l00EE).Name_4 & ">"
      m00A0.M1D80 = l00EE
    Else
      m00A0.Name_4 = m00A0.Name_4 & "<.hide2>"
      m00A0.Type_3 = 0
      MakeErrBeep
    End If
    Exit Function
  Case &HC0FE
17072
    l00F2 = "Controls"
    m00A0.Type_3 = 2
    l00EE = 5
  Case &HC0FD
17073
    l00F2 = "Count"
    m00A0.Type_3 = 0
    l00EE = 0
  Case Else
17074
    If l00EE = 0 Then
      If gv120C Then
      Else
        ShowErrMsg 1
      End If
    End If
17075
    Select Case m00A0.Type_3
    Case 2
      Select Case p00EC And &HC000
      Case &HC000
        l00EE = gVBObjsDefs(l00EE).id_5
L14914:
17076
        If l00EE > 0 Then
          If l00F0 > gKnowControlsData(l00EE).M2BC4 Then
            l00EE = 0
          Else
            l00EE = gKnowControlsData(l00EE).M2BBB
            l00EE = gv1478(l00EE + l00F0)
            l00F2 = gv1444(l00EE).Name_4
          End If
          m00A0.Type_3 = 1
        End If
      Case &H8000
L149BE:
17077
        l00EE = gVBObjsDefs(l00EE).mFrmCount
        If l00EE = 0 Then
          If gv120C Then
            Beep
          Else
            MakeErrBeep
          End If
        Else
          l00EE = l00EE + l00F0
          m00A0.Type_3 = 3
          l00F2 = gv0FF6(l00EE).Name_4
        End If
      Case 0
        Select Case l00EE
        Case 0
17078
          MakeErrBeep
        Case 1, 2, 4
17079
          l00EE = 0
          If l00F0 <= UBound(gv1076) Then
            l00EE = gv1076(l00F0).M2651
            If l00EE Then
              l00F0 = gv1076(l00F0).M2658
              l00EE = gVBObjsDefs(l00EE).mFrmCount
              l00F2 = gv0FF6(l00EE + l00F0).Name_4
              l00EE = 4
              m00A0.Type_3 = 2
            End If
          End If
          If l00EE = 0 Then
            If l00F0 <= UBound(gv1042) Then
              l00EE = gv1042(l00F0)
              If l00EE Then
                m00A0.Type_3 = 1
                l00F2 = gv1444(l00EE).Name_4
              End If
            End If
          End If
          If l00EE = 0 Then
            MakeErrBeep
          End If
        Case Else
17370
          l00F2 = fn03DF(l00F0)
        End Select
      Case Else
17371
        MakeErrBeep
      End Select
    Case 3
      Select Case p00EC And &HC000
      Case &HC000
17372
      Case Else
17373
      End Select
      l00EE = Abs(gv0FF6(l00EE).id_5)
      GoTo L14914
    Case Else
17374
      If l00F0 <= UBound(gv1042) Then
        l00EE = gv1042(l00F0)
        m00A0.Type_3 = 1
        l00F2 = gv1444(l00EE).Name_4
      Else
17375
        If gv120C Then
          Beep
        Else
          MakeErrBeep
        End If
17376
        l00F2 = fn03DF(p00EC)
        m00A0.Type_3 = 0
      End If
    End Select
  End Select
  m00A0.M1D80 = l00EE
17377
  If m00A0.Type_3 = 1 Then
    l00EE = gv1444(l00EE).Type_3
    Select Case l00EE And gc129E
    Case mc003E: l00EE = &HB
    Case mc0040: l00EE = &HE
    Case mc0042: l00EE = &HF
    Case mc0044: l00EE = &HC
    Case mc004C: l00EE = 1
    Case m004E: l00EE = 4
    Case gc12BC: l00EE = &H14
    Case mc0046: l00EE = &H10
    Case Else: l00EE = 0
    End Select
    If l00EE Then

17378
      m00A0.M1D80 = l00EE
      m00A0.Type_3 = 2
    End If
  End If
  If Len(l00F2) = 0 Then
    l00F2 = "P?" & Hex$(p00EC)
    m00A0.Type_3 = 0
  End If

17379
  m00A0.Name_4 = m00A0.Name_4 & "." & l00F2
  fn0B6F = l00F2
End Function

Sub sub0B31(p010E As Integer)
Dim hFile As Integer
Dim l0112%
Dim l0114%

17100
  gv1476 = 0: gv14AA = 1
  ReDim gv1444(0), gv1478(0)
  gv150E = 0: gv1542 = 1
  ReDim g_EventNames(0), gv1510(0)
  gKnowControlsCount = 0
  ReDim gKnowControlsData(0)
  
  hFile = FreeFile
  std_p_e = AppPathwithSlash & "std_p_e.300"
  
  If p010E Then DispLog2 p010E, "Loading " & std_p_e
  Open std_p_e For Binary Access Read As hFile
  
  Get hFile, , l0112
  If l0112 <> (VBX_VER3 Or &H8000) Then
    Close hFile
    ShowErrMsg &H4000
    End
  End If
  
  Get hFile, , l0114
  For l0112 = 0 To l0114
    Get hFile, , l0114
    gv12D4(l0112) = Space$(l0114)
    Get hFile, , gv12D4(l0112)
  Next l0112
  
  Get hFile, , l0114
  For l0112 = 0 To l0114
    Get hFile, , l0114
    gv12BE(l0112) = Space$(l0114)
    Get hFile, , gv12BE(l0112)
  Next l0112
  
  Get hFile, , l0114: ReDim gv1444(-l0114 To 0)
  m008E = Uint32(-l0114) - 1
  For l0112 = 1 To l0114 + 1
    Get hFile, , gv1444(1 - l0112)
  Next l0112
  
  Get hFile, , l0114: ReDim g_EventNames(-l0114 To 0)
  m0092 = Uint32(-l0114) - 1
  For l0112 = 1 To l0114 + 1
    Get hFile, , g_EventNames(1 - l0112)
  Next
  Close hFile
  
  If gv12BE(mc0042) = "" Then
    gv12BE(mc0042) = "Field"
    gv12BE(mc0044) = "TableDef"
    gv12BE(mc0046) = "Index"
  End If
  
End Sub

Sub sub0B15()
' called (p1 As type, p2 As type, p3 As type, p4 As type)
17010
End Sub

Sub sub0B24()
End Sub

Sub sub0710()
Dim l012E As Integer

17080
  Select Case m00A0.Type_3
  Case 1
17081
    l012E = gv1444(m00A0.M1D80).Type_3
    If l012E And gc129C Then
      l012E = l012E Xor gc129C
    Else
    End If
  Case 2
17082
    l012E = m00A0.M1D80
    If l012E < 0 Then
      l012E = -l012E
    ElseIf l012E = 20 Then
      l012E = 0
    Else
      If gVBObjsDefs(l012E).id_5 < 0 Then
17083
        If -gVBObjsDefs(l012E).id_5 = gVBObjsDefs(l012E - 1).id_5 Then
          l012E = l012E - 1
        Else
          l012E = -gVBObjsDefs(l012E).id_5
        End If
      Else
17084
        l012E = gVBObjsDefs(l012E).id_add_5
        If l012E < 0 Then
          l012E = -l012E
        Else
          l012E = Abs(gVBObjsDefs(l012E).id_add_5)
        End If
      End If
17085
      If l012E Then
        m00A0.Name_4 = m00A0.Name_4 & "<." & gVBObjsDefs(l012E).Name_4 & ">"
      Else
        m00A0.Name_4 = m00A0.Name_4 & "<.defcoll>"
      End If
    End If
  Case 3
17086
    l012E = m00A0.M1D80
  End Select
17087
  m00A0.Name_4 = m00A0.Name_4 & "()"
  If l012E Then
    m00A0.M1D80 = l012E
  Else
    m00A0.Type_3 = 0
  End If
End Sub

Sub LoadVBX(LogCount As Integer, ByVal VBXFile As String)
Dim hFile As Integer
Dim l0136%
Dim l0138%
Dim i%
Dim PathLen As Integer
Dim VBX_Size%
Dim FileLength As Long
Dim KnownControls As T2BA1
Dim VBX_Version As Integer
Dim l014A%  'Type
Dim l014C%  'Type
Dim RawToken%  'Type
Dim AppPath As String
 Const mc0164 = 4 ' &H4%

17090
  
'Replace ".VBX" with ".300"
  PathLen = ExtractPath(VBXFile, ".")
  If PathLen = 0 Then
    VBXFile = VBXFile & ".300"
  Else
    VBXFile = Left$(VBXFile, PathLen) & "300"
  End If
  
  If LogCount Then DispLog2 LogCount, "Loading " & VBXFile

On Error GoTo File_300_NotFound

Retry:
17091

' Look for .300 in ProjectDir
  AppPath = gCurPath
  std_p_e = AppPath & VBXFile
  PathLen = Len(Dir$(std_p_e))
  
  
  If PathLen = 0 Then
    
  ' Look for .300 in AppDir
    AppPath = AppPathwithSlash
    std_p_e = AppPath & VBXFile
    PathLen = Len(Dir$(std_p_e))
  End If
  
 'Raise Err if not found
  If PathLen = 0 Then Error 1000
17092
  
  FileLength = FileLen(std_p_e)
  
  hFile = FreeFile
  
'open .300 File
  Open std_p_e For Binary Access Read As hFile
  Get hFile, , VBX_Version
  Get hFile, , VBX_Size
  
  If VBX_Version = VBX_VER4 Then
17093
17095
    FileLength = 0
    Error 1005
    If FileLength = 0 Then
      ShowErrMsg &H4000
      Error 1002
    End If
  ElseIf VBX_Version <> VBX_VER3 Then
    Error 1001
  End If
  
  ReDim Preserve gKnowControlsData(gKnowControlsCount + VBX_Size)
  For l0136 = 1 To VBX_Size
    If VBX_Version = VBX_VER3 Then
      Get hFile, , KnownControls
    Else
      Stop
      Get hFile, , l014C
      Get hFile, , RawToken
'      sub0B15 l014A, l014C, RawToken, KnownControls
    End If
    
'Sub sub0B15()
' called (p1 As type, p2 As type, p3 As type, p4 As T2BA1)
    
17096
    If ScanOnly Then
      If LogCount Then DispLog2 LogCount, "Loading " & KnownControls.ControlTypeName
      FrmMain.dum900.AddItem KnownControls.ControlTypeName
    End If
    
    If KnownControls.M2BC4 >= 0 Then
      If VBX_Version = VBX_VER3 Then
        KnownControls.M2BBB = gv14AA
        ReDim Preserve gv1478(gv14AA + KnownControls.M2BC4 + 1)
        ReDim Preserve gv1444(LBound(gv1444) To gv1476 + KnownControls.M2BC4 + 1)
      End If
17097
      For i = 0 To KnownControls.M2BC4
        
        If VBX_Version = VBX_VER3 Then
          
          Get hFile, , PathLen
          
          If PathLen > 0 Then
            gv1476 = gv1476 + 1
            Get hFile, , gv1444(gv1476)
            PathLen = gv1476
          End If
          
          gv1478(gv14AA) = PathLen
          gv14AA = gv14AA + 1
        Else
          PathLen = gv1478(KnownControls.M2BBB + i)
        End If
        
        If PathLen > 0 Then
          If (gv1444(PathLen).Type_3 And &HFF) = &H70 Then
            Select Case gv1444(PathLen).Name_4
            Case "Fields": gv1444(PathLen).Type_3 = mc0042 Or gc129C
            Case "TableDefs": gv1444(PathLen).Type_3 = mc0044 Or gc129C
            Case "Indexes": gv1444(PathLen).Type_3 = mc0046 Or gc129C
            Case Else: ShowErrMsg &H1000
            End Select
          End If
        End If
        
      Next i
    Else
      KnownControls.M2BBB = 0
    End If
17098
    If KnownControls.M2BE4 >= 0 Then
      If VBX_Version = VBX_VER3 Then
        KnownControls.M2BDA = gv1542
        ReDim Preserve gv1510(gv1542 + KnownControls.M2BE4 + 1)
        ReDim Preserve g_EventNames(LBound(g_EventNames) To gv150E + KnownControls.M2BE4 + 1)
17099
        For l0138 = 0 To KnownControls.M2BE4
          Get hFile, , PathLen
          If PathLen > 0 Then
            gv150E = gv150E + 1
            PathLen = gv150E
            Get hFile, , g_EventNames(gv150E)
          End If
          gv1510(gv1542) = PathLen
          gv1542 = gv1542 + 1
        Next l0138
      End If
    Else
      KnownControls.M2BDA = 0
    End If
    gKnowControlsCount = gKnowControlsCount + 1
    
    gKnowControlsData(gKnowControlsCount) = KnownControls
  Next l0136

'-------------------------------

Ignore:
  If hFile Then Close hFile
  If l014A Then Close l014A
Exit Sub

File_300_NotFound:
  If hFile Then Close hFile: hFile = 0
  If l014A Then Close l014A: l014A = 0
  gv7B32 = gv7B32 And Not &H40
'  ShowErrMsg &H40
'Stop
'  If MsgBoxResponse = vbRetry Then Resume Retry
  Resume Ignore
End Sub
