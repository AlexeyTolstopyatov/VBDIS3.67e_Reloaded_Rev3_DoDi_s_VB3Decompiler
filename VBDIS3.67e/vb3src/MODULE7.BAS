'Attribute VB_Name = "MODULE7"
' Module7
Option Explicit

Sub sub06F3(p0010 As Integer)
Dim l0012 As Integer

11000
  l0012 = p0010 - Len(gDisOutputTxt)
  If l0012 > 0 Then gDisOutputTxt = gDisOutputTxt & Space$(l0012)
 End Sub

Sub sub08B9(p0014 As Integer, p0016 As Long, p0018%, p001A%)
11020
  sub082B p0014, p0016, Uint32(p0018), Uint32(p001A)
  p0018 = p001A
End Sub

Sub sub0845(p001E%, p0020&, p0022 As Long, p0024 As Long)
11040
  If (p0022 Xor p0024) And 1 Then
    sub082B p001E, p0020, p0022, p0024
  Else
    sub0853 p001E, p0020, p0022, p0024
  End If
End Sub

Sub sub0853(hfile_para As Integer, Item As Long, p002A As Long, p002C As Long)
Dim l002E%
Dim l0030 As Integer
Dim l0032 As Long
Dim lfile_buff%
Dim l0036 As String

11050
  l0036 = Hex$(p002C): If Len(l0036) < 4 Then l0036 = Space$(4)
  l0032 = p002C - p002A
  If hFile_Out = 0 And l0032 > 512 Then l0032 = 512
  Seek hfile_para, Item + p002A
  l0030 = 16
  While l0032 > 0
    RSet l0036 = Hex$(p002A): gDisOutputTxt = l0036
    If l0032 < l0030 Then l0030 = l0032
    For l002E = 1 To l0030 Step 2
      Get hfile_para, , lfile_buff
      DisOutput fn02AC(lfile_buff)
    Next l002E
    p002A = p002A + l0030
    l0032 = l0032 - l0030
    WriteOutputTxt
  Wend
  If p002A < p002C Then
    gDisOutputTxt = "...": WriteOutputTxt
  End If
  p002A = p002C
  Seek hfile_para, Item + p002A
End Sub

Sub sub0863(p003A As String, p003C As Integer)
Dim l003E As Integer

11060
  l003E = InStr(p003C, p003A, Chr$(0)) - p003C
  DisOutput " '" & Mid$(p003A, p003C, l003E) & "'"
  WriteOutputTxt
End Sub

Sub CloseOutputFile()
11070
  
  If hFile_Out Then
    Close hFile_Out
    hFile_Out = 0
  End If
  
  Set Listbox_File_Out = Nothing
  Screen.MousePointer = vbDefault

End Sub

Sub sub0879(hfile_para As Integer, p0046 As Long, p0048 As Long)
Dim i%
Dim l004C As Integer
Dim lfile_buff As String
Dim l0050 As Integer
Dim l0052 As Long
Dim l0054 As Long

11080
  If p0046 = 0 Then Exit Sub
  Seek hfile_para, p0046
  l0052 = p0048 - p0046
  If hFile_Out = 0 And l0052 > 512 Then l0052 = 512
  l004C = 16
  lfile_buff = Space$(l004C)
  While l0052 > 0
    If l0054 < &H1000 Then
      gDisOutputTxt = IntToHexStr(Int(l0054))
    Else
      gDisOutputTxt = Hex$(l0054)
    End If
    If l0052 < 16 Then l004C = l0052: lfile_buff = Space$(l0052)
    Get hfile_para, , lfile_buff
    For i = 1 To l004C
      l0050 = Asc(Mid$(lfile_buff, i, 1))
      DisOutput fn029A(l0050)
      If l0050 < 32 Then Mid$(lfile_buff, i, 1) = " "
    Next i
    If 54 > Len(gDisOutputTxt) Then DisOutput Space$(54 - Len(gDisOutputTxt))
    DisOutput lfile_buff
    WriteOutputTxt
    l0054 = l0054 + l004C
    l0052 = l0052 - l004C
  Wend
End Sub

Sub sub0883(hfile_para As Integer, f_offset As Long, p005E As Long)
Dim i%
Dim l0064 As Integer
Dim Disp_Lvl As Long
Dim l006A As Long
Dim lfileBuffInt%

11090
  l006A = 0
  Seek hfile_para, f_offset
  Disp_Lvl = p005E - f_offset
  If hFile_Out = 0 And Disp_Lvl > 512 Then Disp_Lvl = 512
  l0064 = 16
  While Disp_Lvl > 0
    If l006A < &H1000 Then
      gDisOutputTxt = IntToHexStr(Int(l006A))
    Else
      gDisOutputTxt = Hex$(l006A)
    End If
    If Disp_Lvl < l0064 Then l0064 = Disp_Lvl
    For i = 1 To l0064 Step 2
      Get hfile_para, , lfileBuffInt
      DisOutput fn02AC(lfileBuffInt)
    Next i
    l006A = l006A + l0064
    Disp_Lvl = Disp_Lvl - l0064
    WriteOutputTxt
  Wend
End Sub

Sub WriteOutputTxt()
On Error Resume Next
11100
  If hFile_Out Then
    Print #hFile_Out, gDisOutputTxt
  Else
    If Len(gDisOutputTxt) Then
      Listbox_File_Out.AddItem gDisOutputTxt
    End If
  End If
  gDisOutputTxt = ""
End Sub

Sub OpenOutputFile(FileName As String)
11110
  If hFile_Out Then
    MsgBox "file already open"
    Close hFile_Out
  End If
  hFile_Out = FreeFile
  
  Screen.MousePointer = vbHourglass
  
  Open FileName For Output As hFile_Out
End Sub

Sub sub06E9(p0070 As Integer)
11120
  gDisOutputTxt = Space$(p0070)
  gv0C5E = p0070
End Sub

Sub sub0896(p0074 As Control)
11130
  Set Listbox_File_Out = p0074
  p0074.Clear
End Sub

Sub sub08A0(p0076 As String, p0078 As Integer, p007A As String)
Dim i%
Dim l007E As Integer

11140
  For i = 1 To Len(p007A)
    Select Case Mid$(p007A, i, 1)
    Case "b"
      DisOutput Right$(Hex$(Asc(Mid$(p0076, p0078, 1)) + &H100), 2)
      p0078 = p0078 + 1
    Case "w"
      DisOutput Right$(Hex$(Asc(Mid$(p0076, p0078 + 1, 1)) + &H100), 2)
      DisOutput Right$(Hex$(Asc(Mid$(p0076, p0078, 1)) + &H100), 2)
      p0078 = p0078 + 2
    Case "l"
      DisOutput Right$(Hex$(Asc(Mid$(p0076, p0078 + 3, 1)) + &H100), 2)
      DisOutput Right$(Hex$(Asc(Mid$(p0076, p0078 + 2, 1)) + &H100), 2)
      DisOutput Right$(Hex$(Asc(Mid$(p0076, p0078 + 1, 1)) + &H100), 2)
      DisOutput Right$(Hex$(Asc(Mid$(p0076, p0078, 1)) + &H100), 2)
      p0078 = p0078 + 4
    Case "n"
      WriteOutputTxt
    Case "s"
      l007E = Asc(Mid$(p0076, p0078, 1))
      DisOutput "'" & Mid$(p0076, p0078 + 1, l007E - 1) & "'"
      p0078 = p0078 + l007E + 1
    Case "t"
      l007E = Len(gDisOutputTxt) Mod 14
      If l007E Then DisOutput Space$(14 - l007E)
    Case "z"
      l007E = InStr(p0078, p0076, Chr$(0)) - p0078
      DisOutput "'" & Mid$(p0076, p0078, l007E) & "'"
    Case "Z"
      l007E = InStr(p0078, p0076, Chr$(0)) - p0078
      DisOutput "'" & Mid$(p0076, p0078, l007E) & "'"
      p0078 = p0078 + l007E + 1
    Case "\"
      p0078 = p0078 + 1
    Case Else
      DisOutput Mid$(p007A, i, 1)
    End Select
  Next i
End Sub

Sub DisOutput(DisOutputTxt As String)
11150
  gDisOutputTxt = gDisOutputTxt & DisOutputTxt
End Sub

Sub sub08AA(p0082 As Long, p0084() As String)
Dim l008A As Long
Dim l008C
Dim l0090 As Integer

11160
  l008A = 1
  For l008C = 0 To UBound(p0084)
    If p0082 And l008A Then
      GoSub LFE72
      DisOutput p0084(l008C)
      p0082 = p0082 Xor l008A
    End If
    l008A = l008A * 2
  Next
  If p0082 Then
    GoSub LFE72
    DisOutput "0x" & Hex$(p0082)
  End If
Exit Sub

LFE72:
  If l0090 Then
    DisOutput " | "
  Else
    DisOutput " ": l0090 = True
  End If
Return
End Sub

Sub sub082B(p0092 As Integer, p0094 As Long, p0096 As Long, p0098 As Long)
Dim l009C%
Dim l009E As Integer
Dim l00A0 As Integer
Dim l00A2 As Long
Dim l00A6 As String
Dim l00A8 As String

11030
  l00A8 = Hex$(p0098): If Len(l00A8) < 4 Then l00A8 = Space$(4)
  Seek p0092, p0094 + p0096
  l00A2 = p0098 - p0096
  If hFile_Out = 0 And l00A2 > 512 Then l00A2 = 512
  l009E = 16
  l00A6 = Space$(l009E)
  While l00A2 > 0
    If l00A2 < l009E Then l009E = l00A2: l00A6 = Space$(l009E)
    Get p0092, , l00A6
    For l009C = 1 To l009E
      l00A0 = Asc(Mid$(l00A6, l009C, 1))
      DisOutput fn029A(l00A0)
      If l00A0 < 32 Or l00A0 > &H7E Then Mid$(l00A6, l009C, 1) = " "
    Next l009C
    If 49 > Len(gDisOutputTxt) Then DisOutput Space$(49 - Len(gDisOutputTxt))
    RSet l00A8 = Hex$(p0096)
    gDisOutputTxt = l00A8 & gDisOutputTxt & l00A6: WriteOutputTxt
    p0096 = p0096 + l009E
    l00A2 = l00A2 - l009E
  Wend
  If p0096 <> p0098 Then
    gDisOutputTxt = "...": WriteOutputTxt
    p0096 = p0098
  End If
  Seek p0092, p0094 + p0096
End Sub
