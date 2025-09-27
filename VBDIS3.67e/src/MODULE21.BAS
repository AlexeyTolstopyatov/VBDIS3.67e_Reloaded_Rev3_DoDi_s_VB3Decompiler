Attribute VB_Name = "MODULE21"
' Module21
Option Explicit

Sub sub0FC1()
Dim l001A As Integer
Dim l001C As String

29000
  l001C = gv7B86
  l001A = ExtractPath(l001C, "\")
  If l001A Then
    gv7B8A = Left(l001C, l001A)
    l001C = Mid$(l001C, l001A + 1)
  End If
  l001A = InStr(l001C, ".")
  If l001A Then l001C = Left$(l001C, l001A - 1)
  NewPrjName = l001C
End Sub

Function PathAddExt(lPath As String, ByVal lExt As String) As String
Dim Dot_Pos As Integer

29010
  If gWorkDir = "" Then FrmPrjDir.Show 1
  
  If Len(lPath$) Then
    
    Dot_Pos = InStr(lPath, ".")
    If Dot_Pos Then
      PathAddExt = gWorkDir & Left$(lPath, Dot_Pos) & lExt
    Else
      PathAddExt = gWorkDir & lPath & "." & lExt
    End If
  
  Else
    
    PathAddExt = gWorkDir & NewPrjName & "." & lExt
  
  End If
  
End Function

Sub sub0FDA()
Dim l002C As Integer
Dim l002E As Integer
Dim l0030 As String
Dim l0032 As String

29020
  gv7BC8 = 0
  ReDim gv7B96(20)
29021
  l002E = FreeFile
  gCurPath = gv7B8A
  Open gv7B86 For Input As #l002E
29022
  While Not EOF(l002E)
29023
    Line Input #l002E, l0030
    l002C = InStr(l0030, "=")
    If l002C Then
29024
      If Mid$(l0030, l002C + 1, 1) = """" Then
        l0032 = Mid$(l0030, l002C + 2)
        l0032 = Left$(l0032, Len(l0032) - 1)
29025
        Select Case Left$(l0030, l002C - 1)
        Case "ExeName"
          gIn_FileNameAndExt = l0032
        Case "Path"
          gCurPath = l0032
        End Select
      End If
    Else
29026
      If Right$(l0030, 1) <> "X" And Right$(l0030, 3) <> "DLL" Then
        gv7BC8 = gv7BC8 + 1
29027
        If gv7BC8 >= UBound(gv7B96) Then ReDim Preserve gv7B96(gv7BC8 + 20)
29028
        gv7B96(gv7BC8) = l0030
      End If
    End If
  Wend
  Close l002E
  l002C = InStr(gIn_FileNameAndExt, ".")
  If l002C Then gIn_FileNameAndExt = Left$(gIn_FileNameAndExt, l002C - 1)
29029
  If Right$(gCurPath, 1) <> "\" Then gCurPath = gCurPath & "\"
  gIn_FileFulPath = gCurPath & gIn_FileNameAndExt & ".EXE"
End Sub
