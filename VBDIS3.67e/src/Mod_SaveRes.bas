Attribute VB_Name = "Mod_SaveRes"
Option Explicit

Type TYPE_RES_INFO_Str
  Type As String * 1
  TypeAsInt As Integer
End Type

Type TYPE_RES_INFO
  ResID As TYPE_RES_INFO_Str
  IndexID As TYPE_RES_INFO_Str
  Version As Integer
  ResSize As Long
End Type


Sub CreateResFile()
'Find Version Res in Resource RCData
   If ResID_VerInfomation Then
      
    
     'Make *.res Filename
      Dim Res_Filename$
      Res_Filename = PathAddExt(gVBTitle, "RES")
      
     'Seek To VerInfoRes
      ResSeek (ResID_VerInfomation)
          
     'copy VerInfoRes to *.res
      Dim ReadBuff$, ResSize%
      Get hVBFile, CurResOffsetStart, ResSize
      
      Dim ResHeader As TYPE_RES_INFO
      ResHeader.ResID.Type = Chr(&HFF)
      ResHeader.ResID.TypeAsInt = &H10
      
      ResHeader.IndexID.Type = Chr(&HFF)
      ResHeader.IndexID.TypeAsInt = &H1
      
      ResHeader.Version = &H30
      
      ResHeader.ResSize = ResSize
      
     'Create *.res
      Dim hFileRes%
      hFileRes = FreeFile
      Open Res_Filename For Binary As hFileRes
      Put hFileRes, , ResHeader
      
      ReadBuff = Space(ResSize)
      Get hVBFile, CurResOffsetStart, ReadBuff
      Put hFileRes, , ReadBuff
      
      Close hFileRes

   End If

End Sub

