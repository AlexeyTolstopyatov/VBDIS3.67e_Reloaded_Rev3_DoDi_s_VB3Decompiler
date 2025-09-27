'Attribute VB_Name = "ErrorHandle_MODULE15"
' Module15
Option Explicit

Sub sub0394(p0006 As Integer, p0008 As Integer, p000A As Integer)
20010
  If p000A Then
    gv0FF6(p0006).id_5 = -p0008
  Else
    gv0FF6(p0006).id_5 = p0008
  End If
End Sub

Sub ShowRuntimeErr()
20020
  MsgBox Error$ & Chr$(13) & Chr$(10) & "in line" & Erl
End Sub

Sub sub0D9D(p0014 As Integer)
 Const mc0016 = -32768 ' &H8000%
20000
  If p0014 Then Exit Sub
  ShowErrMsg mc0016
  End
End Sub
