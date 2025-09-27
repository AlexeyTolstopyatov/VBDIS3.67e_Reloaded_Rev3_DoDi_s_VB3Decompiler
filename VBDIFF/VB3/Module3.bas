' Module3
Option Explicit

Function fn0102 (p0082 As String) As Long
  gv0104.M04AC = p0082
  LSet gv010A = gv0104
  fn0102 = gv010A.M04DD
End Function

Function fn0115 (p0090 As Integer) As String
  fn0115 = Right$(Hex$(p0090 Or &H100), 2)
 End Function

Function fn011D (p0094 As Integer) As String
  fn011D = " " & Right$(Hex$(p0094 Or &H100), 2)
End Function

Function fn0126 (p0098 As Long) As String
Dim l009C As String * 9
  RSet l009C = Hex$(p0098)
  fn0126 = l009C
End Function

Function fn012F (p00A0 As Integer) As String
  fn012F = " " & Right$(Hex$(p00A0 Or &H10000), 4)
End Function

Function fn0138 (ByVal p00A4 As Long) As String
  fn0138 = Right$(Hex$(p00A4 Or &H10000), 4)
End Function

Function fn0140 (p00A8 As Long) As Integer
  fn0140 = (p00A8 And &HFFFF0000) \ &H10000
End Function

Function fn014A (p00AC As Integer, p00AE As Integer) As Long
  fn014A = (CLng(p00AC) And &HFFFF&) * p00AE
End Function

Function fn0152 (p00B2 As Long) As Integer
  If p00B2 And &H8000& Then
    fn0152 = CInt(p00B2 Or &HFFFF0000)
  Else
    fn0152 = p00B2 And &HFFFF&
  End If
End Function

Function fn015C (p00B6 As Integer) As String
  gv00C6.M04C0 = p00B6
  LSet gv00B2 = gv00C6
  fn015C = gv00B2.M04AC
End Function

Function fn0163 (p00BA As Long) As String
  gv010A.M04DD = p00BA
  LSet gv0104 = gv010A
  fn0163 = gv0104.M04AC
End Function

Function fn016A (p00BE As String, p00C0 As Integer) As Integer
  gv00B2.M04AC = Mid$(p00BE, p00C0, 2)
  LSet gv00C6 = gv00B2
  fn016A = gv00C6.M04C0
End Function

Function fn0173 (p00C4 As String) As String
  If Right$(p00C4, 1) = "\" Then fn0173 = p00C4 Else fn0173 = p00C4 & "\"
End Function

Function fn0182 (p00C8 As String, p00CA As String) As Integer
Dim l00CC As Integer
  Do
    fn0182 = l00CC
    l00CC = InStr(l00CC + 1, p00C8, p00CA)
  Loop While l00CC
End Function

Function fn018C (p00D0 As String) As String
Dim l00D2
  For l00D2 = 1 To Len(p00D0)
    Mid$(p00D0, l00D2, 1) = Chr$(Asc(Mid$(p00D0, l00D2, 1)) Xor &H1F Xor l00D2)
  Next
  fn018C = p00D0
End Function

Function fn01A8 (p00DA%, p00DC As Integer) As Integer
Dim l00DE As Long
  l00DE = fn01B8(p00DA) + p00DC
  If l00DE >= &H8000& Then l00DE = l00DE Or &HFFFF0000
  fn01A8 = l00DE
End Function

Function fn01B0 (p00E2%, p00E4 As Integer) As Integer
  fn01B0 = fn01B8(p00E2) \ p00E4
End Function

Function fn01B8 (p00E8 As Integer) As Long
  fn01B8 = CLng(p00E8) And &HFFFF&
End Function

Function fn01C0 (p00EC%, p00EE As Integer) As Integer
Dim l00F0 As Long
  l00F0 = fn01B8(p00EC) * p00EE
  If l00F0 >= &H8000& Then l00F0 = l00F0 Or &HFFFF0000
  fn01C0 = l00F0
End Function

Function fn01D2 (p00F8 As String, p00FA As Integer) As String
Dim l00FC As Integer
  If p00FA < 0 Then
    fn01D2 = "$" & fn0138(p00FA) & "?"
  Else
    l00FC = InStr(p00FA, p00F8, Chr$(0)) - p00FA
    fn01D2 = Mid$(p00F8, p00FA, l00FC)
  End If
End Function

Function fn01DA (p0100 As String) As String
Dim l0102 As Integer
  l0102 = InStr(p0100, Chr$(0))
  If l0102 Then
    fn01DA = Left$(p0100, l0102 - 1)
  Else
    fn01DA = p0100
  End If
End Function

Sub sub00F5 (p0076 As ComboBox, p0078 As Variant)
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

Sub sub0109 (p0084 As Integer, p0086 As Integer)
Dim l008C As Long
  gv0182.M051A = p0086
  gv0182.M0522 = p0084
  l008C = p0086 And &HFFFF&
  gv0182.M0557 = (l008C And &H1F) * 2
  gv0182.M0550 = (l008C \ &H20) And &H3F
  gv0182.M0549 = (l008C \ &H800) And &H1F
  l008C = p0084 And &HFFFF&
  gv0182.M0542 = l008C And &H1F
  gv0182.M053B = (l008C \ &H20) And &HF
  gv0182.M0533 = ((l008C \ &H200) And &H7F) + 1980
  gv0182.M0515 = DateSerial(gv0182.M0533, gv0182.M053B, gv0182.M0542) + TimeSerial(gv0182.M0549, gv0182.M0550, gv0182.M0557)
  gv0182.M01D2 = Format$(gv0182.M0515, "hh:mm:ss")
  gv0182.M052B = Format$(gv0182.M0515, "dd.mm.yyyy")
End Sub

Sub sub019C (p00D4 As Variant)
  gv0182.M0515 = p00D4
  gv0182.M0533 = Year(p00D4)
  gv0182.M053B = Month(p00D4)
  gv0182.M0542 = Day(p00D4)
  gv0182.M0549 = Hour(p00D4)
  gv0182.M0550 = Minute(p00D4)
  gv0182.M0557 = Second(p00D4)
  gv0182.M051A = (fn01C0(gv0182.M0549, &H800) Or gv0182.M0550 * &H20 Or gv0182.M0557 \ 2)
  gv0182.M0522 = (gv0182.M0533 - 1980) * &H200 + gv0182.M053B * &H20 + gv0182.M0542
  gv0182.M01D2 = Format$(gv0182.M0515, "hh:mm:ss")
  gv0182.M052B = Format$(gv0182.M0515, "dd.mm.yyyy")
End Sub

Sub sub01C8 (p00F2 As Variant)
  gv0182.M0515 = p00F2
End Sub

