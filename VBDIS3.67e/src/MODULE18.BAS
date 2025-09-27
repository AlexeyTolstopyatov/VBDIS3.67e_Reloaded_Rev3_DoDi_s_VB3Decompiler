Attribute VB_Name = "MODULE18"
' Module18
Option Explicit
Const mc0060 = -32632 ' &H8088%
Const mc0062 = 16640 ' &H4100%
Const mc0064 = 8016 ' &H1F50%
Const vbCritical = 16 ' &H10%
Const mc0074 = 48 ' &H30%
Const mc0076 = 1 ' &H1%

Function fn0EF1() As Integer
26000
  fn0EF1 = gv7B32
End Function

Sub ShowErrMsg(p0068 As Variant)
 Const mc0070 = 16 ' &H10%
 Const mc0074 = 48 ' &H30%
 Const mc0076 = 1 ' &H1%
26010
  If gv7B36 = 0 Then
    gv7B36 = gDisOutLogDbgLineCount: gv7B38 = CurOffset
  End If
  
  If VarType(p0068) = 8 Then
    MsgBox p0068, mc0070, News_from_VB_Discompiler
    gv7B30 = -1
  Else
    If gv7B32 And p0068 Then Exit Sub
    ErrorMsg Val(p0068), mc0074, mc0076
  End If
End Sub

Function fn0F0B() As Integer
26020
  fn0F0B = gv7B30
End Function

Sub sub0418()
26030
  gv7B30 = 0: gv7B36 = 0
End Sub

Sub ErrorMsg(p007A As Integer, MsgBoxStyle As Integer, MsgBoxButtons As Integer)
Dim l0080 As Integer
Dim MsgBoxText As String

26040
  If gv7B30 = 0 Then gv7B30 = p007A
  If p007A And mc0060 Then l0080 = l0080 Or gc7B82
  If p007A And mc0062 Then l0080 = l0080 Or gc7B80
  If p007A And mc0064 Then l0080 = l0080 Or gc7B7E
  If l0080 = 0 Then l0080 = gc7B7C
  gv7B32 = gv7B32 Or p007A
  
  Select Case p007A
  Case 0
  Case 1: MsgBoxText = "Found unknown data structures!"
  Case 2: MsgBoxText = "Not a Visual Basic program"
  Case 4: MsgBoxText = "VB Version" & Str$(VBver_11) & " is not supported"
  Case 8: MsgBoxText = "Found an unknown resource!"
  Case &H10: MsgBoxText = "Found unknown fixups!"
  Case &H20: MsgBoxText = "Error in Discompiler logic!"
  Case &H40
   
   'Replace ".300" with "VBX" in MsgBoxText
    If Right$(std_p_e, 4) = ".300" Then
      MsgBoxText = Left$(std_p_e, Len(std_p_e) - 3) & "VBX"
    Else
      MsgBoxText = std_p_e
    End If
    
    MsgBoxText = "Missing description for " & MsgBoxText
    MsgBoxStyle = vbCritical
'    MsgBoxButtons = vbRetryCancel
    MsgBoxButtons = vbAbortRetryIgnore + vbDefaultButton3
    
  Case &H80: MsgBoxText = std_p_e & " contains unknown structure!"
  Case &H100: MsgBoxText = "Found an unknown token!"
  Case &H200: MsgBoxText = "Unexpected variable reference!"
  Case &H400: MsgBoxText = "Found incompatible scopes!"
  Case &H800: MsgBoxText = "Found incompatible types"
  Case &H1000: MsgBoxText = "Found an unknown collection!"
  Case &H2000: MsgBoxText = "An already known problem occured"
  Case &H4000: MsgBoxText = "File not found or wrong version: " & Chr$(10) & std_p_e
  Case &H8000: MsgBoxText = "The program is too big for this version of VB Discompiler"
  Case Else: MsgBoxText = "Error " & Format$(p007A)
  End Select
  If Len(MsgBoxText) Then
    If Right$(MsgBoxText, 1) = "!" Then MsgBoxText = MsgBoxText & Chr$(10) & "You may send this program to DoDi to improve VB Discompiler"
    MsgBoxResponse = MsgBox(MsgBoxText, MsgBoxStyle Or MsgBoxButtons, News_from_VB_Discompiler)
    Select Case MsgBoxResponse
    Case vbRetry
      Exit Sub
    Case vbCancel, vbAbort, vbNo
      GoTo QuitApp
    End Select
  End If
  
'stop
' on useless hint...
   Exit Sub
  
  l0080 = l0080 And Not UnloadReason
  UnloadReason = UnloadReason Or l0080
  Select Case l0080
  Case 0
  Case gc7B82
    MsgBox "For this program you'll need an upgraded VB Discompiler", vbCritical, ErrorSituationVBDiscompiler
    MsgBoxResponse = vbIgnore
  Case gc7B80
    MsgBoxResponse = MsgBox(gc032A, 48 + 1, ErrorSituationVBDiscompiler)
  Case gc7B7E
    MsgBoxResponse = MsgBox(gc032E, 32 + 1, ErrorSituationVBDiscompiler)
  Case Else
    MsgBoxResponse = MsgBox(gc0332, 64 + 1, ErrorSituationVBDiscompiler)
  End Select
  
  Select Case MsgBoxResponse
   Case vbCancel, vbAbort, vbNo
  
  Case Else
    Exit Sub
  
  End Select

QuitApp:
  UnloadReason = UnloadReason Or gc7B84
  Unload FrmMain
End Sub
