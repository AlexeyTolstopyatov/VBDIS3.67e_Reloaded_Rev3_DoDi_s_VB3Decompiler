' Module1
Option Explicit

Function Descramble (ByVal iodata As String) As String
Dim i
  For i = 1 To Len(iodata)
    Mid$(iodata, i, 1) = Chr$(Asc(Mid$(iodata, i, 1)) Xor &H1F Xor i)
  Next
  Descramble = iodata
End Function

Function fn008D (p0082 As String) As Long
  gv009E.M044F = p0082
  LSet gv00A4 = gv009E
  fn008D = gv00A4.M0480
End Function

Function fn00A0 (p0090 As Integer) As String
  fn00A0 = Right$(Hex$(p0090 Or &H100), 2)
 End Function

Function fn00A8 (p0094 As Integer) As String
  fn00A8 = " " & Right$(Hex$(p0094 Or &H100), 2)
End Function

Function fn00B1 (p0098 As Long) As String
Dim l009C As String * 9
  RSet l009C = Hex$(p0098)
  fn00B1 = l009C
End Function

Function fn00BA (p00A0 As Integer) As String
  fn00BA = " " & Right$(Hex$(p00A0 Or &H10000), 4)
End Function

Function fn00C3 (ByVal p00A4 As Long) As String
  fn00C3 = Right$(Hex$(p00A4 Or &H10000), 4)
End Function

Function fn00CB (p00A8 As Long) As Integer
  fn00CB = (p00A8 And &HFFFF0000) \ &H10000
End Function

Function fn00D5 (p00AC As Integer, p00AE As Integer) As Long
  fn00D5 = (CLng(p00AC) And &HFFFF&) * p00AE
End Function

Function fn00DD (ByVal p00B2 As Long) As Integer
  If p00B2 And &H8000& Then
    fn00DD = CInt(p00B2 Or &HFFFF0000)
  Else
    fn00DD = p00B2 And &HFFFF&
  End If
End Function

Function fn00E7 (p00B6 As Integer) As String
  gv0060.M0463 = p00B6
  LSet gv004C = gv0060
  fn00E7 = gv004C.M044F
End Function

Function fn00EE (p00BA As Long) As String
  gv00A4.M0480 = p00BA
  LSet gv009E = gv00A4
  fn00EE = gv009E.M044F
End Function

Function fn00F5 (p00BE As String, p00C0 As Integer) As Integer
  gv004C.M044F = Mid$(p00BE, p00C0, 2)
  LSet gv0060 = gv004C
  fn00F5 = gv0060.M0463
End Function

Function fn00FE (p00C4 As String) As String
  If Right$(p00C4, 1) = "\" Then fn00FE = p00C4 Else fn00FE = p00C4 & "\"
End Function

Function fn010D (p00C8 As String, p00CA As String) As Integer
Dim l00CC As Integer
  Do
    fn010D = l00CC
    l00CC = InStr(l00CC + 1, p00C8, p00CA)
  Loop While l00CC
End Function

Function fn0133 (p00DA%, p00DC As Integer) As Integer
Dim l00DE As Long
  l00DE = IntToULng(p00DA) + p00DC
  If l00DE >= &H8000& Then l00DE = l00DE Or &HFFFF0000
  fn0133 = l00DE
End Function

Function fn013B (p00E2%, p00E4 As Integer) As Integer
  fn013B = IntToULng(p00E2) \ p00E4
End Function

Function fn014B (p00EC%, p00EE As Integer) As Integer
Dim l00F0 As Long
  l00F0 = IntToULng(p00EC) * p00EE
  If l00F0 >= &H8000& Then l00F0 = l00F0 Or &HFFFF0000
  fn014B = l00F0
End Function

Function fn015D (p00F8 As String, p00FA As Integer) As String
Dim l00FC As Integer
  If p00FA < 0 Then
    fn015D = "$" & fn00C3(p00FA) & "?"
  Else
    l00FC = InStr(p00FA, p00F8, Chr$(0)) - p00FA
    fn015D = Mid$(p00F8, p00FA, l00FC)
  End If
End Function

Function fn0165 (p0100 As String) As String
Dim l0102 As Integer
  l0102 = InStr(p0100, Chr$(0))
  If l0102 Then
    fn0165 = Left$(p0100, l0102 - 1)
  Else
    fn0165 = p0100
  End If
End Function

Function IntToULng (Value As Integer) As Long
  
  IntToULng = CLng(Value) And &HFFFF&

End Function

Sub sub0080 (p0076 As ComboBox, p0078 As Variant)
Dim l007C
Dim l007E As Long
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

Sub sub0094 (p0084 As Integer, p0086 As Integer)
Dim l008C As Long
  gv011C.M04BD = p0086
  gv011C.M04C5 = p0084
  l008C = p0086 And &HFFFF&
  gv011C.M04FA = (l008C And &H1F) * 2
  gv011C.M04F3 = (l008C \ &H20) And &H3F
  gv011C.M04EC = (l008C \ &H800) And &H1F
  l008C = p0084 And &HFFFF&
  gv011C.M04E5 = l008C And &H1F
  gv011C.M04DE = (l008C \ &H20) And &HF
  gv011C.M04D6 = ((l008C \ &H200) And &H7F) + 1980
  gv011C.M04B8 = DateSerial(gv011C.M04D6, gv011C.M04DE, gv011C.M04E5) + TimeSerial(gv011C.M04EC, gv011C.M04F3, gv011C.M04FA)
  gv011C.M015D = Format$(gv011C.M04B8, "hh:mm:ss")
  gv011C.M04CE = Format$(gv011C.M04B8, "dd.mm.yyyy")
End Sub

Sub sub0127 (p00D4 As Variant)
  gv011C.M04B8 = p00D4
  gv011C.M04D6 = Year(p00D4)
  gv011C.M04DE = Month(p00D4)
  gv011C.M04E5 = Day(p00D4)
  gv011C.M04EC = Hour(p00D4)
  gv011C.M04F3 = Minute(p00D4)
  gv011C.M04FA = Second(p00D4)
  gv011C.M04BD = (fn014B(gv011C.M04EC, &H800) Or gv011C.M04F3 * &H20 Or gv011C.M04FA \ 2)
  gv011C.M04C5 = (gv011C.M04D6 - 1980) * &H200 + gv011C.M04DE * &H20 + gv011C.M04E5
  gv011C.M015D = Format$(gv011C.M04B8, "hh:mm:ss")
  gv011C.M04CE = Format$(gv011C.M04B8, "dd.mm.yyyy")
End Sub

Sub sub0153 (p00F2 As Variant)
  gv011C.M04B8 = p00F2
End Sub

