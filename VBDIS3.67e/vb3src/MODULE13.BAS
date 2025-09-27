'Attribute VB_Name = "MODULE13"
' Module13
Option Explicit
Const mc00AE = 16384 ' &H4000%
Const mc00B0 = 1 ' &H1%
Const mc00B8 = 1 ' &H1%
Const mc00BA = 2 ' &H2%
Const mc00BC = 2 ' &H2%
Const mc00BE = 3 ' &H3%
Const mc00C0 = 3 ' &H3%
Const mc00C2 = 4 ' &H4%
Const mc00C4 = 4 ' &H4%
Const mc00C6 = 5 ' &H5%
Const mc00C8 = 5 ' &H5%
Const mc00CA = 6 ' &H6%
Const mc00CC = 6 ' &H6%
Const mc00CE = 7 ' &H7%
Const mc00D0 = 7 ' &H7%
Const mc00D2 = 12 ' &HC%
Const mc00D4 = 8 ' &H8%
Const mc00D6 = 10 ' &HA%
Const mc00D8 = 9 ' &H9%
Const mc00DA = 8 ' &H8%
Const mc00DC = 10 ' &HA%
Const mc00DE = 9 ' &H9%
Const mc010E = 4 ' &H4%

Sub sub0BE9(ByVal LogCount As Integer)
Dim hFile As Integer
Dim IntBuff%
 Const mc00AE = 16384 ' &H4000%
 Const mc00B0 = 1 ' &H1%
 Const mc00B8 = 1 ' &H1%
 Const mc00BA = 2 ' &H2%
 Const mc00BC = 2 ' &H2%
 Const mc00BE = 3 ' &H3%
 Const mc00C0 = 3 ' &H3%
 Const mc00C2 = 4 ' &H4%
 Const mc00C4 = 4 ' &H4%
 Const mc00C6 = 5 ' &H5%
 Const mc00C8 = 5 ' &H5%
 Const mc00CA = 6 ' &H6%
 Const mc00CC = 6 ' &H6%
 Const mc00CE = 7 ' &H7%
 Const mc00D0 = 7 ' &H7%
 Const mc00D2 = 12 ' &HC%
 Const mc00D4 = 8 ' &H8%
 Const mc00D6 = 10 ' &HA%
 Const mc00D8 = 9 ' &H9%
 Const mc00DA = 8 ' &H8%
 Const mc00DC = 10 ' &HA%
 Const mc00DE = 9 ' &H9%

18030

' Load vbdis3i.dat into VBdis_Struct

  hFile = FreeFile
  std_p_e = "vbdis3i.dat"
  If LogCount Then DispLog2 LogCount, "Loading " & std_p_e
  
  Open AppPathwithSlash & std_p_e For Binary As hFile
  Get hFile, , IntBuff
  
  'File not found or wrong version
  If IntBuff = 0 Then ShowErrMsg 16384
  
  VBdis_String = String$(IntBuff, 0)
  
  Get hFile, , VBdis_Struct
  Get hFile, , VBdis_String
  
  Close hFile
  
  
  gvConv_ff(1) = 0
  gvConv_ff(2) = 1
  gvConv_ff(3) = 2
  gvConv_ff(4) = 3
  gvConv_ff(5) = 4
  gvConv_ff(6) = 5
  gvConv_ff(7) = 6
  gvConv_ff(12) = 7
  gvConv_ff(10) = 8
  gvConv_ff(8) = 9
  gvConv_ff(9) = 10
  
  gvConv_bw(0) = 1
  gvConv_bw(1) = 2
  gvConv_bw(2) = 3
  gvConv_bw(3) = 4
  gvConv_bw(4) = 5
  gvConv_bw(5) = 6
  gvConv_bw(6) = 7
  gvConv_bw(7) = 12
  gvConv_bw(8) = 10
  gvConv_bw(9) = 8
  gvConv_bw(10) = 9
  gVB_a_TypeConv(1) = "Int"
  gVB_a_TypeConv(2) = "Lng"
  gVB_a_TypeConv(3) = "Sng"
  gVB_a_TypeConv(4) = "Dbl"
  gVB_a_TypeConv(5) = "Cur"
  gVB_a_TypeConv(6) = "Var"
  gVB_a_TypeConv(7) = "Str"
  gVB_a_DataTypes(1) = "Integer"
  gVB_a_DataTypes(2) = "Long"
  gVB_a_DataTypes(3) = "Single"
  gVB_a_DataTypes(4) = "Double"
  gVB_a_DataTypes(5) = "Currency"
  gVB_a_DataTypes(6) = "Variant"
  gVB_a_DataTypes(7) = "String"
  gVB_a_DataTypes(8) = "String *"
  
  sub0AE1
End Sub

Function fn0BD9(p00E8 As Integer) As String
Dim l00EA As Integer
Dim l00EC As Integer

18050
  l00EA = Abs(VBdis_Struct.Token9Bit(p00E8).M2CE2)
  If l00EA = 0 Then Exit Function
  
  l00EC = InStr(l00EA, VBdis_String, "§")
  fn0BD9 = Mid$(VBdis_String, l00EA, l00EC - l00EA)
End Function

Function fn0BA0() As String
Dim l00F0 As Integer
Dim l00F2 As Integer

18000
  l00F0 = ATK9B_VBDat_1.M2CE2
  If l00F0 <= 0 Then Exit Function
  l00F2 = InStr(l00F0, VBdis_String, "§")
  fn0BA0 = Mid$(VBdis_String, l00F0, l00F2 - l00F0)
End Function

Sub sub0BAC()
End Sub

Function fn0BBC(p00FC As Integer) As String
Dim l00FE As Integer
Dim l0100 As Integer

18020
  l00FE = VerDepSig + (p00FC And &HFF) * &H18
  l0100 = Asc(Mid$(VBDIS_FlagToken2, l00FE, 1))
  fn0BBC = Mid$(VBDIS_FlagToken2, l00FE + 1, l0100)
End Function

Sub LoadVBDisDatFiles(LogLvl As Integer)
Dim hFilel As Integer
Dim Token2Size As Integer


18040
  Select Case VBver_11
  Case 2: VerDepSig = &H7500
  Case 3: VerDepSig = &H5600
  Case Else
    ShowErrMsg mc010E
    Exit Sub
  End Select
  
  hFilel = FreeFile
  std_p_e = "vbdis" & Format$(VBver_11) & "x.dat"
  
 'show log
  If LogLvl Then DispLog2 LogLvl, "Loading " & std_p_e
  
  Open AppPathwithSlash & std_p_e For Binary As hFilel
  
 'Test Ver
  Get hFilel, , Token2Size
  'Token2Size  0302                   << 7
  Token2Size = Token2Size Xor (VBver_11 * &H100)
  '
  If Token2Size = 2 Then
    
  ' Size of VBDIS_FlagToken = 10837+1(0x2A56) Words
  '                        => 21676(0x54AC) Bytes
    Get hFilel, , VBDIS_FlagToken
    
   '10837 *3 => 32511 (7EFF)
    VBDIS_ControlToken = String$(VBDIS_FlagTokenSize * 3, 0): Get hFilel, , VBDIS_ControlToken
    
   '2+10837+65022 = 75861 '0x12855
    Get hFilel, , Token2Size
    
    
    VBDIS_FlagToken2 = String$(Token2Size, 0): Get hFilel, , VBDIS_FlagToken2
  
  Else
    
    ShowErrMsg mc00AE
    Cancel = True
  
  End If
  
  Close hFilel

End Sub
