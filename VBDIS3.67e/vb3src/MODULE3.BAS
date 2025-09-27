'Attribute VB_Name = "MODULE3"
' Module3
Option Explicit

Sub sub0272(p0076 As ComboBox, p0078 As Variant)
Dim l007C
Dim l007E As Long

5000
  If VarType(p0078) = 8 Then
    For l007C = p0076.ListCount - 1 To 0 Step -1
      If p0076.List(l007C) = p0078 Then Exit For
    Next
  Else
    l007E = p0078
    For l007C = p0076.ListCount - 1 To 0 Step -1
      If p0076.ItemData(l007C) = l007E Then Exit For
    Next
  End If
  p0076.ListIndex = l007C
End Sub

Function fn027F(p0082 As String) As Long
5010
  gv0A90.M1EE9 = p0082
  LSet gv0A96 = gv0A90
  fn027F = gv0A96.M1F1A
End Function

Sub sub0286(p0084 As Integer, p0086 As Integer)
Dim l008C As Long

5020
  gv0B0E.M1F57 = p0086
  gv0B0E.M1F5F = p0084
  l008C = p0086 And &HFFFF&
  gv0B0E.M1F94 = (l008C And &H1F) * 2
  gv0B0E.M1F8D = (l008C \ &H20) And &H3F
  gv0B0E.M1F86 = (l008C \ &H800) And &H1F
  l008C = p0084 And &HFFFF&
  gv0B0E.M1F7F = l008C And &H1F
  gv0B0E.M1F78 = (l008C \ &H20) And &HF
  gv0B0E.M1F70 = ((l008C \ &H200) And &H7F) + 1980
  gv0B0E.M1F52 = DateSerial(gv0B0E.M1F70, gv0B0E.M1F78, gv0B0E.M1F7F) + TimeSerial(gv0B0E.M1F86, gv0B0E.M1F8D, gv0B0E.M1F94)
  gv0B0E.M034F = Format$(gv0B0E.M1F52, "hh:mm:ss")
  gv0B0E.M1F68 = Format$(gv0B0E.M1F52, "dd.mm.yyyy")
End Sub

Function fn0292(p0090 As Integer) As String
5030
  fn0292 = Right$(Hex$(p0090 Or &H100), 2)
 End Function

Function fn029A(p0094 As Integer) As String
5040
  fn029A = " " & Right$(Hex$(p0094 Or &H100), 2)
End Function

Function fn02A3(p0098 As Long) As String
Dim l009C As String * 9

5050
  RSet l009C = Hex$(p0098)
  fn02A3 = l009C
End Function

Function fn02AC(p00A0 As Integer) As String
5060
  fn02AC = " " & Right$(Hex$(p00A0 Or &H10000), 4)
End Function

Function IntToHexStr(ByVal p00A4 As Long) As String
5070
  IntToHexStr = Right$(Hex$(p00A4 Or &H10000), 4)
End Function

Function fn02BD(p00A8 As Long) As Integer
5080
  fn02BD = (p00A8 And &HFFFF0000) \ &H10000
End Function

Function PageToOffset(p00AC As Integer, p00AE As Integer) As Long
5090
  PageToOffset = (CLng(p00AC) And &HFFFF&) * p00AE
End Function

Function ToSInt16(ByVal p00B2 As Long) As Integer
5100
'TestSign
  If p00B2 And &H8000& Then
  ' negative
    ToSInt16 = CInt(p00B2 Or &HFFFF0000)
  Else
  ' positive
    ToSInt16 = p00B2 And &HFFFF&
  End If
End Function

Function fn02D9(p00B6 As Integer) As String
5110
  gv0A52.M1EFD = p00B6
  LSet gv0A3E = gv0A52
  fn02D9 = gv0A3E.M1EE9
End Function

Function fn02E0(p00BA As Long) As String
5120
  gv0A96.M1F1A = p00BA
  LSet gv0A90 = gv0A96
  fn02E0 = gv0A90.M1EE9
End Function

Function fn02E7(p00BE As String, p00C0 As Integer) As Integer
5130
  gv0A3E.M1EE9 = Mid$(p00BE, p00C0, 2)
  LSet gv0A52 = gv0A3E
  fn02E7 = gv0A52.M1EFD
End Function

Function AddPathSlash(ByVal Path As String) As String
5140
  If Right$(Path, 1) = "\" Then AddPathSlash = Path Else AddPathSlash = Path & "\"
End Function

Function ExtractPath(ByVal p00C8 As String, ByVal p00CA As String) As Integer
Dim l00CC As Integer

5150
  Do
    ExtractPath = l00CC
    l00CC = InStr(l00CC + 1, p00C8, p00CA)
  Loop While l00CC
End Function

Function fn0309() As String
End Function

Sub DateTimeSth1(p00D2 As Variant)

5170
  gv0B0E.M1F52 = p00D2
  gv0B0E.M1F70 = Year(p00D2)
  gv0B0E.M1F78 = Month(p00D2)
  gv0B0E.M1F7F = Day(p00D2)
  gv0B0E.M1F86 = Hour(p00D2)
  gv0B0E.M1F8D = Minute(p00D2)
  gv0B0E.M1F94 = Second(p00D2)
  gv0B0E.M1F57 = (fn033D(gv0B0E.M1F86, &H800) Or gv0B0E.M1F8D * &H20 Or gv0B0E.M1F94 \ 2)
  gv0B0E.M1F5F = (gv0B0E.M1F70 - 1980) * &H200 + gv0B0E.M1F78 * &H20 + gv0B0E.M1F7F
  gv0B0E.M034F = Format$(gv0B0E.M1F52, "hh:mm:ss")
  gv0B0E.M1F68 = Format$(gv0B0E.M1F52, "dd.mm.yyyy")
End Sub

Function fn0325(p00D8%, p00DA As Integer) As Integer
Dim l00DC As Long

5180
  l00DC = Uint32(p00D8) + p00DA
  If l00DC >= &H8000& Then l00DC = l00DC Or &HFFFF0000
  fn0325 = l00DC
End Function

Function UDiv(p00E0%, p00E2 As Integer) As Integer
5190
  UDiv = Uint32(p00E0) \ p00E2
End Function

Function Uint32(p00E6 As Integer) As Long
5200
  Uint32 = CLng(p00E6) And &HFFFF&
End Function

Function fn033D(p00EA%, p00EC As Integer) As Integer
Dim l00EE As Long

5210
  l00EE = Uint32(p00EA) * p00EC
  If l00EE >= &H8000& Then l00EE = l00EE Or &HFFFF0000
  fn033D = l00EE
End Function

Sub sub0345(p00F0 As Variant)
5220
  gv0B0E.M1F52 = p00F0
End Sub

Function fn034F(p00F6 As String, p00F8 As Integer) As String
Dim l00FA As Integer

5230
  If p00F8 < 0 Then
    fn034F = "$" & IntToHexStr(p00F8) & "?"
  Else
    l00FA = InStr(p00F8, p00F6, Chr$(0)) - p00F8
    fn034F = Mid$(p00F6, p00F8, l00FA)
  End If
End Function

Function strNullCut(p00FE As String) As String
Dim l0100 As Integer

5240
  l0100 = InStr(p00FE, Chr$(0))
  If l0100 Then
    strNullCut = Left$(p00FE, l0100 - 1)
  Else
    strNullCut = p00FE
  End If
End Function
