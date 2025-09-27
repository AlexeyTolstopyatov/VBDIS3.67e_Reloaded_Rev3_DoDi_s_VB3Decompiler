'Attribute VB_Name = "NE_RES_MODULE8"
' Module8
Option Explicit

Sub SeekToResID(ResID As Integer)
Dim l0052%

12000
  curResID = 0
  For l0052 = Res_ID_gv0DC0 To ResObjectsCount
    If ResObjects1(l0052).ResourceID = ResID Then
      ResSeek l0052
      Exit Sub
    End If
  Next l0052
End Sub

Sub ResSeek(ResID As Integer)
Dim tmpsize As Long

12010
  If ResID = curResID Then Exit Sub
  curResObjects1 = ResObjects1(ResID)
  
  If curResObjects1.Offset = 0 Then Exit Sub
  curResID = ResID
  CurResOffset = 1 + PageToOffset(curResObjects1.Offset, ResAlignment)
  
  tmpsize = PageToOffset(curResObjects1.Size, ResAlignment)
  
  ResSize = CurResOffset + tmpsize
  
  Seek hVBFile, CurResOffset
  
End Sub

Function ReadFixedString(p0062 As Integer) As String
Dim l0064 As String

12020
  l0064 = Space$(p0062)
  Get hVBFile, , l0064
  ReadFixedString = strNullCut(l0064)
End Function
