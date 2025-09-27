' Module4
Option Explicit
Dim m0032_offset As Long
Dim m0032_size As Long
Dim OffsetResourceTable As Long

Function fn02C8 (p018A As Integer) As Long
  fn02C8 = IntToULng(Segs(p018A).M0B77) * NE_Flags
End Function

Sub LoadEntryTable ()
   Dim EntryType As Integer
   Dim EntryTableOffset As Long
   Dim EntryTableRecord As EntryTableStruct
   Dim EntryTableRecord2 As EntryTableStruct2
     
     EntryTableOffset = MZ_FileHeader.Offset_NE_Header + NE_FILEHEADER.ENTRYTABLE
     Objects(ENTRYTABLE).Offset_Start = EntryTableOffset
     
     Get hFile, EntryTableOffset + 1, FIXED_ONE_BYTE_STRING
     
     EntryType = Asc(FIXED_ONE_BYTE_STRING)
     
     While EntryType
       
       Get hFile, , FIXED_ONE_BYTE_STRING
       
       Select Case Asc(FIXED_ONE_BYTE_STRING)
       
          Case 0
          
          Case &HFF
            Seek hFile, Seek(hFile) + EntryType * Len(EntryTableRecord)
          
          Case Else
            Seek hFile, Seek(hFile) + EntryType * Len(EntryTableRecord2)
          
       End Select
       
       Get hFile, , FIXED_ONE_BYTE_STRING
       
       EntryType = Asc(FIXED_ONE_BYTE_STRING)
     
     Wend
     
     Objects(ENTRYTABLE).Offset_NE_Header = Loc(hFile) - EntryTableOffset

End Sub

Sub LoadImportTable ()
   Dim ImportTbl_Offset As Long
   Dim ModRefTbl_Offset As Long
   Dim ModRefData As Integer'
   Dim iModTbl As Integer'
   Dim MaxModRefData As Integer
       
   ModRefTbl_Offset = MZ_FileHeader.Offset_NE_Header + NE_FILEHEADER.ModuleReferenceTable + 1
   ImportTbl_Offset = MZ_FileHeader.Offset_NE_Header + NE_FILEHEADER.IMPORTTABLE + 1

   Objects(IMPORTTABLE).Offset_Start = ImportTbl_Offset - 1

   For iModTbl = 1 To NE_FILEHEADER.ModuleTableEntryCount

      Get hFile, ModRefTbl_Offset, ModRefData
      ModRefTbl_Offset = ModRefTbl_Offset + 2
      
      Get hFile, ImportTbl_Offset + ModRefData, FIXED_ONE_BYTE_STRING
      ModRefData = ModRefData + Asc(FIXED_ONE_BYTE_STRING)
   
      If ModRefData > MaxModRefData Then MaxModRefData = ModRefData
   
   Next iModTbl

   Objects(IMPORTTABLE).Offset_NE_Header = MaxModRefData
   
End Sub

Sub LoadModuleReferenceTable ()
    Objects(MODULE_REF_TABLE).Offset_Start = MZ_FileHeader.Offset_NE_Header + NE_FILEHEADER.ModuleReferenceTable
    Objects(MODULE_REF_TABLE).Offset_NE_Header = NE_FILEHEADER.ModuleTableEntryCount * 2
End Sub

Sub LoadNonResidentNameTable ()
  
  Dim size As Integer
  
  Dim Offset As Long
  Offset = NE_FILEHEADER.NonResidentNameTable
  Objects(NonResidentNameTable).Offset_Start = Offset

  Get hFile, Offset + 1, FIXED_ONE_BYTE_STRING
  
  size = Asc(FIXED_ONE_BYTE_STRING)
  
  While size
    Get hFile, Seek(hFile) + size + 2, FIXED_ONE_BYTE_STRING
    size = Asc(FIXED_ONE_BYTE_STRING)
  Wend
  
  Objects(NonResidentNameTable).Offset_NE_Header = Loc(hFile) - Offset

End Sub

Sub LoadResData2 ()
Dim i%
Dim l0174 As Long
Dim Filepos As Long
  For i = 1 To gv07D8
    ResData2_sub02BC i
    
    If m0032_size Then
      
      If l0174 = 0 Then
        l0174 = m0032_offset: Filepos = m0032_offset + m0032_size
      
      ElseIf m0032_offset < l0174 Then
        l0174 = m0032_offset
      
      ElseIf m0032_offset >= Filepos Then
        Filepos = m0032_offset + m0032_size
      End If

    End If
  Next
  
  Objects(RES_DATA2).Offset_Start = l0174 - 1
  Objects(RES_DATA2).Offset_NE_Header = Filepos - l0174
  
  If Filepos > LOF(hFile) Then
     Filepos = LOF(hFile) - Filepos + 1
    Objects(RES_DATA2).RelSize = Filepos
  End If

End Sub

Sub LoadResidentNameTable ()
Dim ResidentNameTable As Long
Dim Readbuff As String
  
  ResidentNameTable = MZ_FileHeader.Offset_NE_Header + NE_FILEHEADER.ResidentNameTable

  Objects(gc0818).Offset_Start = ResidentNameTable

  Get hFile, ResidentNameTable + 1, FIXED_ONE_BYTE_STRING
  
  While Asc(FIXED_ONE_BYTE_STRING)
    
    Readbuff = Space$(Asc(FIXED_ONE_BYTE_STRING) + 2)
    Get hFile, , Readbuff
    Get hFile, , FIXED_ONE_BYTE_STRING
  Wend
  Objects(gc0818).Offset_NE_Header = Loc(hFile) - ResidentNameTable
End Sub

Private Sub LoadResourceTable ()
Dim ResourceTableBuff As ResourceTableType
Dim l0146 As Integer
Dim l0148 As Integer
Dim Filepos As Long
Dim l014C As Integer
  Objects(RESOURCETABLE).Offset_Start = MZ_FileHeader.Offset_NE_Header + NE_FILEHEADER.RESOURCETABLE
  
  If NE_FILEHEADER.RESOURCETABLE < NE_FILEHEADER.ResidentNameTable Then
    OffsetResourceTable = MZ_FileHeader.Offset_NE_Header + NE_FILEHEADER.RESOURCETABLE + 1
    
    Get hFile, OffsetResourceTable, gv05FC
    
    gv05FE = 2 ^ gv05FC
    
    ReDim gv073E(gc073A), gv0770(gc073A): gv07A2 = 0
    ReDim gv07A6(0): gv07D8 = 0
    Do
      Filepos = Loc(hFile) + 2
      Get hFile, , ResourceTableBuff

      If ResourceTableBuff.M07E0 = 0 Then Exit Do
      
      gv07A2 = gv07A2 + 1
      
      l0148 = ResourceTableBuff.M07E0 Xor &H8000
      
      If l0148 < 0 Then
        BeepSound
      Else
        gv0770(l0148) = gv07D8
        gv073E(l0148) = ResourceTableBuff
        l0146 = ResourceTableBuff.M0C00
        If l0146 Then
          ReDim Preserve gv07A6(gv07D8 + l0146)
          While l0146 > 0
            gv07D8 = gv07D8 + 1

            Get hFile, , gv07A6(gv07D8)
            
            If gv07A6(gv07D8).M07E7 > 0 Then
              
              If l014C = 0 Then
                l014C = gv07A6(gv07D8).M07E7
              
              ElseIf gv07A6(gv07D8).M07E7 > l014C Then
                l014C = gv07A6(gv07D8).M07E7
              
              End If

            ElseIf gv07A6(gv07D8).M07E7 = 0 Then
              
              gv073E(l0148).M0C00 = gv073E(l0148).M0C00 - 1
              gv07D8 = gv07D8 - 1
            
            End If
            
            l0146 = l0146 - 1
          Wend
        End If
      End If
    Loop

  End If

  Objects(RESOURCETABLE).Offset_NE_Header = Filepos - OffsetResourceTable - 1
  Objects(gc0816).Offset_Start = Filepos
  
  If l014C Then
    Get hFile, OffsetResourceTable + l014C, FIXED_ONE_BYTE_STRING
    l014C = l014C + 2 + Asc(FIXED_ONE_BYTE_STRING)
  End If
  
  Objects(gc0816).Offset_NE_Header = l014C

End Sub

Sub LoadSegs ()
Dim i
  ReDim Segs(Segments)
  
  Objects(SEGMENTTABLE).Offset_Start = MZ_FileHeader.Offset_NE_Header + NE_FILEHEADER.SegmentTableOffset
  Objects(SEGMENTTABLE).Offset_NE_Header = Segments * Len(Segs(1))
  
  Seek hFile, MZ_FileHeader.Offset_NE_Header + NE_FILEHEADER.SegmentTableOffset + 1
  
  For i = 1 To Segments
    Get hFile, , Segs(i)
  Next i

End Sub

Sub LoadVBCODE ()
   Dim iSegs%
   Dim l017C As Long
   Dim Filepos As Long
   Dim Buff%
   Dim l0186 As VBCODEStruct

     For iSegs = 1 To Segments
       
       LoadVBCODE_sub02AD iSegs
       
       If gv05F0 Then
         
         If l017C = 0 Then
           l017C = gv05EC: Filepos = gv05EC + gv05F0
         
         ElseIf gv05EC < l017C Then
           l017C = gv05EC
         
         ElseIf gv05EC >= Filepos Then
           Filepos = gv05EC + gv05F0
           
           Get hFile, Filepos, Buff
           
           If Buff > 0 Then
             Filepos = Loc(hFile) + Buff * Len(l0186)
           
           End If

         End If
       End If
     Next

   Objects(VBCODE).Offset_Start = l017C - 1
   Objects(VBCODE).Offset_NE_Header = Filepos - l017C - 1

End Sub

Sub LoadVBCODE_sub02AD (p013C As Integer)
  gv05F4 = p013C

  If gv05F4 Then
    
    gv05EC = 1 + fn02C8(p013C)
    gv05F0 = CLng(Segs(p013C).M0B81) And &HFFFF&
    
    If gv05F0 = 0 Then
      If Segs(p013C).M0B77 Then gv05F0 = &H10000
    End If
    
    Seek hFile, gv05EC
  
  Else
    
    gv05EC = 0
    gv05F0 = 0

  End If
End Sub

Sub ReadFileData (FileName As String)
Dim iObj
  NE_FILEHEADER.Signature = 0
  
  If hFile Then
    Close hFile
  Else
    hFile = FreeFile
  End If

  Open FileName For Binary As hFile
  
  Erase Objects
  
 'Read NE-Header
  Get hFile, , MZ_FileHeader
  
  Objects(MZ_HEADER).Offset_NE_Header = MZ_FileHeader.Offset_NE_Header
  
 'Check MZ-signature
  If (MZ_FileHeader.Signature <> MZ_Magic) Or (MZ_FileHeader.RelocTable < &H40) Then Exit Sub
  
  Objects(NE_HEADER).Offset_Start = MZ_FileHeader.Offset_NE_Header
  Objects(NE_HEADER).Offset_NE_Header = Len(NE_FILEHEADER)
  
 'Read NE-Header
  Get hFile, MZ_FileHeader.Offset_NE_Header + 1, NE_FILEHEADER
  
 'Check NE-signature
  If NE_FILEHEADER.Signature <> Offset_NE_Header Then
    NE_FILEHEADER.Signature = 0
    Close hFile
    Exit Sub
  End If
  
  NE_Flags = 2 ^ NE_FILEHEADER.MiscFlags
  Segments = NE_FILEHEADER.SegmentTableEntryCount
  
  LoadSegs
  LoadResourceTable
  LoadResidentNameTable
  LoadModuleReferenceTable
  LoadImportTable
  LoadEntryTable
  LoadNonResidentNameTable

  LoadVBCODE
  LoadResData2
  
  For iObj = 1 To VBCODE - 1

    If Objects(iObj + 1).Offset_Start > Objects(iObj).Offset_Start + Objects(iObj).Offset_NE_Header Then
      
      Objects(iObj).RelSize = Objects(iObj).Offset_Start + Objects(iObj).Offset_NE_Header - Objects(iObj + 1).Offset_Start
    
    End If

  Next

End Sub

Sub Rebuild_NEHeader (FileHeader As NE_FileHeader_Struct, ResData As String)
   Dim item As Integer '
   Dim l00F6 As Integer
   Dim itemBase As Long
   Dim itemStart As Integer
   Dim itemRelSize As Integer
   Dim itemSize As Long
   Dim ResDataSize As Integer
   Dim l0102 As Long
  
  ResDataSize = Len(ResData)
  Randomize
  For item = MZ_HEADER To RES_DATA2
    
    itemStart = itemBase - Objects(item).Offset_Start
    itemSize = Objects(item).Offset_NE_Header
    
    Select Case item
    
    Case SEGMENTTABLE      '3
      FileHeader.SegmentTableOffset = FileHeader.SegmentTableOffset + itemStart
    
    Case RESOURCETABLE    '4
      FileHeader.RESOURCETABLE = FileHeader.RESOURCETABLE + itemStart
      
      If ResDataSize Then
        If gv073E(gc0738).M07E0 = 0 Then
          gv073E(gc0738).M07E0 = gc0738 Or &H8000
          gv07A2 = gv07A2 + 1
        End If
        If gv073E(gc0738).M0C00 = 0 Then
          gv073E(gc0738).M0C00 = 1
          gv0770(gc0738) = gv07D8
          gv07D8 = gv07D8 + 1
          If gv07D8 > UBound(gv07A6) Then
            ReDim Preserve gv07A6(gv07D8)
          End If
        End If
      End If

      itemSize = gv07A2 * Len(gv073E(1)) + gv07D8 * Len(gv07A6(1)) + 4
      gv083E = itemSize - Objects(item).Offset_NE_Header
    
    Case gc0816              '5
      If itemSize = 0 Then itemSize = 1
    
    Case gc0818              '6
      FileHeader.ResidentNameTable = FileHeader.ResidentNameTable + itemStart
    
    Case MODULE_REF_TABLE    '7
      FileHeader.ModuleReferenceTable = FileHeader.ModuleReferenceTable + itemStart
    
    Case IMPORTTABLE         '8
      FileHeader.IMPORTTABLE = FileHeader.IMPORTTABLE + itemStart
    
    Case ENTRYTABLE          '9
      FileHeader.ENTRYTABLE = FileHeader.ENTRYTABLE + itemStart
    
    Case NonResidentNameTable '10
      FileHeader.NonResidentNameTable = FileHeader.NonResidentNameTable + itemStart
    
    Case VBCODE    '11
      gv0840 = (itemStart + NE_Flags - 1) \ NE_Flags
      itemBase = Objects(item).Offset_Start + gv0840 * NE_Flags
    
    Case RES_DATA2  '12
      gv0842 = (itemStart + gv05FE - 1) \ gv05FE
      itemBase = Objects(item).Offset_Start + gv0842 * gv05FE
      l0102 = Rnd * &H10000
      FileHeader.M08E4 = fn00DD(l0102)
      l0102 = FileHeader.M08E4 And &H7FFF
      If ResDataSize Then
        FileHeader.M08B6 = FileHeader.M08B6 + &H100
        If gv073E(gc0738).M0C00 Then
          l00F6 = gv0770(gc0738) + 1
          itemSize = itemSize - gv07A6(l00F6).Offset_NE_Header * gv05FE
          gv07A6(l00F6).Offset_Start = 0
          gv07A6(l00F6).Offset_NE_Header = (ResDataSize + gv05FE - 1) \ gv05FE
          gv07A6(l00F6).M07ED = &H1C30
          gv07A6(l00F6).M07E7 = &H8001
        End If
        itemSize = itemSize + ResDataSize
        For l00F6 = ResDataSize To ResDataSize \ 2 Step -1
          l0102 = (l0102 * 3 + Asc(Mid$(ResData, l00F6))) And &H7FFF&
        Next
      End If
      FileHeader.M08EC = fn00DD(Rnd(-l0102) * &H10000)
    End Select
    
    
    itemStart = itemBase - Objects(item).Offset_Start
    itemSize = itemSize + (itemSize And 1)
    itemRelSize = itemSize - Objects(item).Offset_NE_Header

    Objects(item).RelStart = itemStart
    Objects(item).RelSize = itemRelSize

    itemBase = itemBase + itemSize

  Next

End Sub

Sub ResData2_sub02BC (index As Integer)
  
  m0032_offset = 1 + IntToULng(gv07A6(index).Offset_Start) * gv05FE
  
  m0032_size = IntToULng(gv07A6(index).Offset_NE_Header) * gv05FE
  
  Seek hFile, m0032_offset

End Sub

Sub sub01E8 (p00D4 As Integer, p00D6 As Long)
Dim Buff As String
Dim l00DA As Integer
Dim l00DC
Dim l00E2 As T0E25
Dim l00E6 As T0E6C
Dim l00E8 As Long
  l00E8 = Seek(p00D4)
  Get hFile, , l00E2: Put p00D4, , l00E2
  Buff = Space$(l00E2.M0E57)
  Get hFile, , Buff: Put p00D4, , Buff
  Buff = Space$(l00E2.M0E62)
  Get hFile, , Buff: Put p00D4, , Buff
  Do
    Get hFile, , l00E6
    l00DA = Asc(l00E6.M0E7C)
    Select Case l00DA
    Case 0
      Exit Do
    Case gc08C6, gc08C8
      Buff = Space$(Asc(l00E6.M07D1))
      Get hFile, , Buff
      If l00DA = gc08C8 Then
        l00E6.M07D1 = Chr$(0)
        Buff = ""
      End If
    Case gc08CA
      If Asc(l00E6.M07D1) = &HFF Then
        Get hFile, , FIXED_ONE_BYTE_STRING
        Buff = Space$(Asc(FIXED_ONE_BYTE_STRING))
        Get hFile, , Buff
        Buff = FIXED_ONE_BYTE_STRING & Buff
      Else
        Buff = ""
      End If
    Case Else
      Exit Do
    End Select
    Put p00D4, , l00E6
    Put p00D4, , Buff
    Select Case l00DA
    Case gc08C8, gc08CA
      Get hFile, , l00DC: Put p00D4, , l00DC
      If l00DC > 0 Then
        Buff = Space$(l00DC * 4)
        Get hFile, , Buff: Put p00D4, , Buff
      End If
    End Select
  Loop
  p00D6 = p00D6 - (Seek(p00D4) - l00E8)
  Buff = String$(p00D6, 0)
  Put p00D4, , Buff
End Sub

Sub WriteFileData (FileName As String, FileNameRes As String)
Dim l010E As Integer
Dim Filebuff As String
Dim Buff2% '
Dim iSeg%
Dim iRes_data2 As Long
Dim l0118 As Integer
Dim i
Dim FileHeaderOutBuff As NE_FileHeader_Struct
Dim Seg_Item As SegmentStruct
Dim l0124 As Integer
Dim l0128 As VBCODEStruct
Dim l012C As ResourceTableType
Dim l0130 As T0C07
Dim l0132() As Integer
Dim iMax As Integer
Dim l013A As Integer
  
  If hFile = 0 Then BeepSound: Exit Sub

On Error GoTo 0
  FileHeaderOutBuff = NE_FILEHEADER

  Rebuild_NEHeader FileHeaderOutBuff, FileNameRes
  
  l010E = FreeFile
  Open FileName For Binary As l010E
  
 'Fill MZ-Header with Spaces
  Filebuff = Space$(Objects(MZ_HEADER).Offset_NE_Header)
  Get hFile, 1, Filebuff
  Put l010E, 1, Filebuff

  
  Put l010E, , FileHeaderOutBuff
  
  ReDim l0132%(Segments)
  
  l0124 = Objects(VBCODE).Offset_Start \ NE_Flags + gv0840
  
  For iSeg = 1 To Segments
    
    Seg_Item = Segs(iSeg)
    
    LoadVBCODE_sub02AD iSeg
    
    If gv05F0 Then
      
      Seg_Item.M0B77 = l0124
      
      If Segs(iSeg).M07ED And gc06CA Then
        
        Get hFile, gv05EC + gv05F0, Buff2

        If Buff2 > 0 Then gv05F0 = gv05F0 + Buff2 * Len(l0128)

        gv05F0 = gv05F0 + 2

      End If

      Buff2 = (gv05F0 + NE_Flags - 1) \ NE_Flags

      l0124 = l0124 + Buff2

    Else
      Buff2 = 0

    End If

  ' write Object
    Put l010E, , Seg_Item

    l0132(iSeg) = Buff2

  Next
  
  Put l010E, Objects(RESOURCETABLE).Offset_Start + Objects(RESOURCETABLE).RelStart + 1, gv05FC

  
 'Delete FormName Pointer
  If Started Then
    iSeg = gv0770(gc0730)
    
    iMax = gv073E(gc0730).M0C00
    
    For i = 1 To iMax
      
      l013A = gv07A6(iSeg + i).M07E7
      
      If l013A And 1 Then

        If (l013A And &H7FFF) > 4 Then l013A = 0

      End If

      If l013A = 0 Then
        
        gv07A6(iSeg + i).M09B7 = -1

        iMax = iMax - 1

      End If

    Next

  End If

'Save Res_data2
  iRes_data2 = Objects(RES_DATA2).Offset_Start \ gv05FE + gv0842
  For iSeg = 1 To gv07D8
    If gv07A6(iSeg).M09B7 = 0 Then
      
      If gv07A6(iSeg).Offset_NE_Header Then
        gv07A6(iSeg).M09B7 = iRes_data2
        iRes_data2 = iRes_data2 + gv07A6(iSeg).Offset_NE_Header
      Else
        Stop
      End If
    Else
      gv07A6(iSeg).M09B7 = 0
    End If
  Next

'Save ?
  For iSeg = 1 To gc073A
    
    If gv073E(iSeg).M0C00 Then
      l012C = gv073E(iSeg)
      If l012C.M07E0 > 0 Then l012C.M07E0 = l012C.M07E0 + gv083E
      Put l010E, , l012C
      l0118 = gv0770(iSeg)
      For i = 1 To gv073E(iSeg).M0C00
        l0130 = gv07A6(l0118 + i)

        If l0130.M07E7 > 0 Then l0130.M07E7 = l0130.M07E7 + gv083E
        l0130.Offset_Start = l0130.M09B7
        l0130.M09B7 = 0
        Put l010E, , l0130
      Next
    End If

  Next

  iSeg = 0: Put l010E, , iSeg

  Buff2 = Objects(gc0816).Offset_NE_Header

  If Buff2 Then
    Filebuff = Space$(Buff2)
    Get hFile, Objects(gc0816).Offset_Start + 1, Filebuff
    Put l010E, , Filebuff
  End If
  
  Buff2 = Objects(gc0816).RelSize
  Filebuff = String$(Buff2, 0)
  Put l010E, , Filebuff
  For iSeg = gc0818 To NonResidentNameTable
    Buff2 = Objects(iSeg + 1).Offset_Start - Objects(iSeg).Offset_Start
    Buff2 = Buff2 + Objects(iSeg + 1).RelStart - Objects(iSeg).RelStart
    Filebuff = Space$(Buff2)
    Get hFile, Objects(iSeg).Offset_Start + 1, Filebuff: Put l010E, , Filebuff
  Next
  l0124 = Objects(VBCODE).Offset_Start \ NE_Flags + gv0840
  For iSeg = 1 To Segments
    LoadVBCODE_sub02AD iSeg
    If gv05F0 Then
      gv05F0 = l0132(iSeg) * NE_Flags
      Filebuff = Space$(&H7000)
      While gv05F0 > Len(Filebuff)
        Get hFile, , Filebuff: Put l010E, , Filebuff
        gv05F0 = gv05F0 - Len(Filebuff)
      Wend
      Filebuff = Space$(gv05F0)
      Get hFile, , Filebuff: Put l010E, , Filebuff
    End If
    l0124 = l0124 + l0132(iSeg)
  Next
  
  For iSeg = 1 To gv07D8
    ResData2_sub02BC iSeg
    If gv07A6(iSeg).M09B7 Then
      If iSeg = gv0770(gc0730) + 1 Then
        sub01E8 l010E, m0032_size
      ElseIf m0032_offset > 1 Then
        Filebuff = Space$(&H7000)
        While m0032_size > Len(Filebuff)
          Get hFile, , Filebuff: Put l010E, , Filebuff
          m0032_size = m0032_size - Len(Filebuff)
        Wend
        Filebuff = Space$(m0032_size)
        Get hFile, , Filebuff
        If iSeg = gv0770(gc0730) + 1 Then
          Do
            i = InStr(Filebuff, ".FRM")
            If i = 0 Then Exit Do
            Mid$(Filebuff, i, 4) = "   ?"
          Loop
        End If
        Put l010E, , Filebuff
      Else
        If Len(FileNameRes) Then Put l010E, , FileNameRes
      End If
    End If
  Next
  Close hFile
End Sub

