Attribute VB_Name = "NE_RES_MODULE8"
' Module8
Option Explicit

Sub SeekToRCData(ResID As Integer)
Dim i%

12000
  curResID = 0
  For i = ResID_RCData To ResID_RCData_Count
    If ResChilds(i).ResourceID = ResID Then
      ResSeek i
      Exit Sub
    End If
  Next i
End Sub

Sub ResSeek(ResID As Integer)
Dim CurResSize As Long

12010
 'Check if already there
  If ResID = curResID Then Exit Sub
  
 'Exit if Resource is empty (has of offset)
  curResChild = ResChilds(ResID)
  If curResChild.Offset = 0 Then Exit Sub
  
  curResID = ResID
  
 'Seek to Res Start in File
  CurResOffsetStart = PageToOffset(curResChild.Offset, ResAlignment) + 1
  Seek hVBFile, CurResOffsetStart
  
 'Calc Res End Offset
  CurResSize = PageToOffset(curResChild.size, ResAlignment)
  CurResOffsetEnd = CurResOffsetStart + CurResSize

End Sub

Function ReadFixedString(Chars As Integer) As String
Dim Buff As String

12020
  Buff = Space$(Chars)
  Get hVBFile, , Buff
  ReadFixedString = strNullCut(Buff)
End Function
