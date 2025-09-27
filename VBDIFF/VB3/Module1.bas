' Module1
Option Explicit

Sub sub0081 ()
  Erase gv0066, gv007E
  gv0094 = 0
  Do
    gv0094 = (gv0094 + 1) And &HF
  Loop Until fn009A(gv0066(gv0094), gv007E(gv0094))
End Sub

Function fn008E (p0032 As String, p0034 As String) As Integer
  If  gv0098 Then
    fn008E = StrComp(Trim$(p0032), Trim$(p0034), gv0096)
  Else
    fn008E = StrComp(p0032, p0034, gv0096)
  End If
 End Function

Function fn009A (p003A As T045C, p003E As T045C) As Integer
  p003A.M0467 = Seek(gv0038)
  p003E.M0467 = Seek(gv003A)
  If  EOF(gv0038) Then
    If  EOF(gv003A) Then
      gv0094 = -1
      fn009A = True
      Exit Function
    End If
  Else
    Line Input #gv0038, p003A.M046E
  End If
  If  EOF(gv003A) Then fn009A = True: Exit Function
  Line Input #gv003A, p003E.M046E
  fn009A = fn008E(p003A.M046E, p003E.M046E)
End Function
