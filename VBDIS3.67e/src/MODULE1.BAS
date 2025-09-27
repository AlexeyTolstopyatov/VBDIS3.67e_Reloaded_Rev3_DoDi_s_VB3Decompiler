Attribute VB_Name = "MODULE1"
' Module1
Option Explicit
Const TypeCURSOR = 1 ' &H1%
Const Type_BITMAP = 2 ' &H2%
Const Type_ICON = 3 ' &H3%
Const Type_MENU = 4 ' &H4%
Const Type_DIALOG = 5 ' &H5%
Const Type_STRING = 6 ' &H6%
Const Type_FONTDIR = 7 ' &H7%
Const Type_FONT = 8 ' &H8%
Const Type_ACCELERATOR = 9 ' &H9%
Const Type_ResData = 10 ' &HA%
Const Type_GROUPCURSOR = 12 ' &HC%
Const Type_GROUPICON = 14 ' &HE%
Const Type_NAMETABLE = 15 ' &HF%
Const Type_Version_Information = 16 ' &H10%

Sub InitResTypeNameStringsArray()
2000
  gResTypeNameStrings(TypeCURSOR) = "CURSOR"
  gResTypeNameStrings(Type_BITMAP) = "BITMAP"
  gResTypeNameStrings(Type_ICON) = "ICON"
  gResTypeNameStrings(Type_MENU) = "MENU"
  gResTypeNameStrings(Type_DIALOG) = "DIALOG"
  gResTypeNameStrings(Type_STRING) = "STRING"
  gResTypeNameStrings(Type_FONTDIR) = "FONTDIR"
  gResTypeNameStrings(Type_FONT) = "FONT"
  gResTypeNameStrings(Type_ACCELERATOR) = "ACCELERATOR"
  gResTypeNameStrings(Type_ResData) = "RCDATA"
  gResTypeNameStrings(Type_GROUPCURSOR) = "GROUP CURSOR"
  gResTypeNameStrings(Type_GROUPICON) = "GROUP ICON"
  gResTypeNameStrings(Type_NAMETABLE) = "NAMETABLE"
  gResTypeNameStrings(Type_Version_Information) = "VERSIONINFO"
End Sub

Function IsInRange(ByVal Value&, ByVal Max&, Optional ByVal Min& = 0) As Integer
   IsInRange = ((Min <= Value) And (Value <= Max))
End Function

Sub SeekToEntrySegment()

  For curSegID = 1 To NE.SegmentTableEntryCount

    curSegOffs = 1 + GetSegOffset(curSegID)
    curSegSize = Uint32(Segs(curSegID).size)
    
    Dim SegPageOffs%
    SegPageOffs = Segs(curSegID).SegOffset
    If IsInRange(NE.Initial_IP, SegPageOffs + curSegSize, SegPageOffs) Then
       Seek hVBFile, curSegOffs
       Exit Sub
    End If
  Next

   curSegOffs = 0
   curSegSize = 0

End Sub


Sub SeekToSegment(SegID As Integer)
2010
  curSegID = SegID
  
  If curSegID Then
    
    curSegOffs = 1 + GetSegOffset(SegID)
    curSegSize = Uint32(Segs(SegID).size)
    
    Seek hVBFile, curSegOffs
  
  Else
    curSegOffs = 0
    curSegSize = 0
  End If
End Sub


Sub ReadFileData()
2020
  NE.Signature = 0
  VBver_11 = 0
  
' Reopen
  If hVBFile Then
    Close hVBFile
  Else
    hVBFile = FreeFile
  End If
  
  Open gIn_FileFulPath For Binary Access Read As hVBFile
  Get hVBFile, , MZ
  
  If MZ.Signature <> MZ_MAGIC Or MZ.RelocTable < &H40 Then Exit Sub
  
  Get hVBFile, MZ.OffsetToNE + 1, NE
  If NE.Signature <> NE_MAGIC Then
    NE.Signature = 0
    Exit Sub
  End If
  
  SegmentAlign = 2 ^ NE.MiscFlags
  LoadEntryTable
  LoadSegmInformation
  LoadResourceTable
  LoadResidentNameTable
  LoadModAndImpTbl
  LoadNonResTbl
End Sub

Sub ApplyReloc(Id As Integer)
Dim i%
Dim VBCODEStruct_ As VBCODEStruct
Dim l00B6 As T1CBD

2030
  If Id = curSegID Then Exit Sub
  
  SeekToSegment Id
  
  If Segs(Id).Flags And RELOCINFO Then
  
    Get hVBFile, curSegOffs + curSegSize, gv0758
    If gv0758 < 0 Then
      gv0758 = 0
    End If
    
    ReDim gv0726(gv0758)
    For i = 1 To gv0758
    
      Get hVBFile, , VBCODEStruct_
      l00B6.Size_M13F4 = VBCODEStruct_.Size_M13F4
      l00B6.M1CC7 = Asc(VBCODEStruct_.M1CA2)
      l00B6.iToken = Asc(VBCODEStruct_.iToken)
      l00B6.M1CB1 = VBCODEStruct_.M1CB1
      l00B6.M1CB7 = VBCODEStruct_.M1CB7
      gv0726(i) = l00B6
    Next i
    
  Else
    gv0758 = 0
    
  End If
End Sub

Sub LoadEntryTable()
Dim EntryTableOffset As Long
Dim EntryType As Integer
Dim ReadBuff1 As Integer
Dim l00BE As Integer
Dim EntryTableRecord As EntryTableStruct
Dim EntryTableRecord2 As EntryTableStruct2

2040
  EntryTableOffset = MZ.OffsetToNE + NE.ENTRYTABLE + 1
  gv067E = 0
  
  Get hVBFile, EntryTableOffset, FixedByteBuff: EntryType = Asc(FixedByteBuff)
  
  While EntryType
    gv067E = gv067E + EntryType
    Get hVBFile, , FixedByteBuff: ReadBuff1 = Asc(FixedByteBuff)
    Select Case ReadBuff1
    
    Case 0
    
    Case &HFF
      
      While EntryType
        Get hVBFile, , EntryTableRecord
        EntryType = EntryType - 1
      Wend
      
    Case Else
      While EntryType
        Get hVBFile, , EntryTableRecord2
        EntryType = EntryType - 1
      Wend
    End Select
    
    Get hVBFile, , FixedByteBuff: EntryType = Asc(FixedByteBuff)
  Wend
  
  
  ReDim gv064C(gv067E)
  
  Get hVBFile, EntryTableOffset, FixedByteBuff: EntryType = Asc(FixedByteBuff)
  
  
  l00BE = 0
  While EntryType
    Get hVBFile, , FixedByteBuff: ReadBuff1 = Asc(FixedByteBuff)
    Select Case ReadBuff1
    
    Case 0
      l00BE = l00BE + EntryType
    
    Case &HFF
      While EntryType
        Get hVBFile, , EntryTableRecord
        gv064C(l00BE).M_4_sub = Asc(EntryTableRecord.M_4_sub)
        gv064C(l00BE).segm0 = Asc(EntryTableRecord.M1C09)
        gv064C(l00BE).Size_M13F4 = EntryTableRecord.M1BEB
        EntryType = EntryType - 1
        l00BE = l00BE + 1
      Wend
    
    Case Else
      While EntryType
        Get hVBFile, , EntryTableRecord2
        gv064C(l00BE).M_4_sub = Asc(EntryTableRecord2.M_4_sub)
        gv064C(l00BE).segm0 = ReadBuff1
        gv064C(l00BE).Size_M13F4 = EntryTableRecord2.M1BEB
        EntryType = EntryType - 1
        l00BE = l00BE + 1
      Wend
    End Select
    
    Get hVBFile, , FixedByteBuff: EntryType = Asc(FixedByteBuff)
    
  Wend
  
End Sub

Sub LoadModAndImpTbl()
Dim ImpTbl As Long
Dim ModulTbl As Long
Dim lFileBuff%
Dim i%

2050
  ModulTbl = MZ.OffsetToNE + NE.ModuleReferenceTable + 1
  ImpTbl = MZ.OffsetToNE + NE.IMPORTTABLE + 1
  
  ReDim VBRUNs(NE.ModuleTableEntryCount)
  
  For i = 1 To NE.ModuleTableEntryCount
    
    Get hVBFile, ModulTbl, lFileBuff: ModulTbl = ModulTbl + 2
    
    Get hVBFile, ImpTbl + lFileBuff, FixedByteBuff
    
    VBRUNs(i) = Space$(Asc(FixedByteBuff))
    Get hVBFile, , VBRUNs(i)
    
    If UCase$(Left$(VBRUNs(i), 5)) = "VBRUN" Then
      VBver_11 = Val(Mid$(VBRUNs(i), 6, 1))
      
    ElseIf Left$(VBRUNs(i), 5) = "VB400" Then
      VBver_11 = 4
    End If
    
  Next i
End Sub

Sub LoadNonResTbl()
Dim Offset As Long
Dim filebuf As String
Dim i%
Dim NullPos As Integer

2060
  Offset = NE.NonResidentNameTable + 1
  
  NonResCount = 0
  Get hVBFile, Offset, FixedByteBuff
  
  While Asc(FixedByteBuff)
    
    NonResCount = NonResCount + 1
    
    filebuf = Space$(Asc(FixedByteBuff) + 2)
    Get hVBFile, , filebuf
    
    Get hVBFile, , FixedByteBuff
    
  Wend
  ReDim ResidentNames2(NonResCount)
  
  Seek hVBFile, Offset
  
  For i = 1 To NonResCount
    
    Get hVBFile, , FixedByteBuff
    
    filebuf = Space$(Asc(FixedByteBuff))
    Get hVBFile, , filebuf
    
   'Nullcut
    NullPos = InStr(filebuf, Chr$(0))
    If NullPos Then filebuf = Left$(filebuf, NullPos - 1)
    
    ResidentNames2(i).Name = filebuf
    
    Get hVBFile, , ResidentNames2(i).AddData
    
  Next i
End Sub


Sub LoadResidentNameTable()
Dim Offset As Long
Dim Filebuff As String
Dim Dummy%
Dim i%

2070
  ResidentCount = 0
  
  Offset = MZ.OffsetToNE + NE.ResidentNameTable + 1
  Get hVBFile, Offset, FixedByteBuff
  
  While Asc(FixedByteBuff)
    ResidentCount = ResidentCount + 1
    
    Filebuff = Space$(Asc(FixedByteBuff))
    
    Get hVBFile, , Filebuff
    Get hVBFile, , Dummy
    Get hVBFile, , FixedByteBuff
  Wend
  
  ReDim ResidentNames(ResidentCount)
  
  Seek hVBFile, Offset
  
  For i = 1 To ResidentCount
    
    Get hVBFile, , FixedByteBuff
    
    Filebuff = Space$(Asc(FixedByteBuff))
    Get hVBFile, , Filebuff
    
   'NullCut
    Dummy = InStr(Filebuff, Chr$(0))
    If Dummy Then Filebuff = Left$(Filebuff, Dummy - 1)
    
    ResidentNames(i).Name = Filebuff
    Get hVBFile, , ResidentNames(i).AddData
    
  Next i
  
End Sub

Sub LoadResourceTable()
Dim ResRootTblItem As ResRootTblItemStruc
Dim lResItem As ResourceChildType
Dim Offset As Long
Dim i%
Dim Childs As Integer
Dim ResIndex As Integer
Dim iChild%
Dim ResObjType As Integer

2080
  If NE.RESOURCETABLE = NE.ResidentNameTable Then Exit Sub
  
  Offset = MZ.OffsetToNE + NE.RESOURCETABLE + 1
  
' Get Alignment shift count for resource data.
  Get hVBFile, Offset, ResAlignment
  ResAlignment = 2 ^ ResAlignment
  
  ReDim ResRootTbl(0): ResRootCount = 0
  ReDim ResChilds(0): ResChildsCount = 0
  
  ResID_VerInfomation = 0
  ResID_RCData = 0
  ResID_RCData_Count = 0

  
'Read ResRoot + Dirs
  Do
    Get hVBFile, , ResRootTblItem
    
  ' Reached End?
    If ResRootTblItem.Type_3 = 0 Then Exit Do
    
    ResRootCount = ResRootCount + 1
    ReDim Preserve ResRootTbl(ResRootCount)
    
    ResRootTbl(ResRootCount).Type_3 = ResRootTblItem.Type_3
    ResRootTbl(ResRootCount).Childs = ResRootTblItem.Childs
    ResRootTbl(ResRootCount).index = ResChildsCount
    
    Childs = ResRootTblItem.Childs
    
   'Is RCData
    Select Case (ResRootTblItem.Type_3 Xor &H8000)
      Case Type_ResData
         ResID_RCData = ResChildsCount + 1
         ResID_RCData_Count = ResChildsCount + Childs
      
      Case Type_Version_Information
         ResID_VerInfomation = ResChildsCount + 1
    
    End Select
    

    
   'GetChilds
    ReDim Preserve ResChilds(ResChildsCount + Childs)
    While Childs > 0
      
      ResChildsCount = ResChildsCount + 1
      
      Get hVBFile, , lResItem
      
      ResChilds(ResChildsCount).Type_3 = ResRootTblItem.Type_3
      ResChilds(ResChildsCount).size = lResItem.size
      ResChilds(ResChildsCount).FlagWord = lResItem.FlagWord
      ResChilds(ResChildsCount).ResourceID = lResItem.ResourceID
      ResChilds(ResChildsCount).Offset = lResItem.Offset
      
      Childs = Childs - 1
    Wend
    
  Loop
  
  
'
  For i = 1 To ResRootCount
    ResObjType = ResRootTbl(i).Type_3
    
   'Get ResDirNames
    If ResObjType < 0 Then
    
    Else
    
      Get hVBFile, Offset + ResObjType, FixedByteBuff
      ResRootTbl(i).Name_4 = Space$(Asc(FixedByteBuff))
      Get hVBFile, , ResRootTbl(i).Name_4
      
    End If
    
    
   'Get ResChildNames
    ResIndex = ResRootTbl(i).index
    For iChild = ResIndex + 1 To ResIndex + ResRootTbl(i).Childs
      
      ResObjType = ResChilds(iChild).ResourceID
      If ResObjType < 0 Then
       'Generate Name (for ex. "800a")
        ResChilds(iChild).Name_4 = Hex$(ResObjType)
      Else
       'Get Name From File
        Get hVBFile, Offset + ResObjType, FixedByteBuff
        ResChilds(iChild).Name_4 = Space$(Asc(FixedByteBuff))
        Get hVBFile, , ResChilds(iChild).Name_4
        
      End If
    
    Next iChild
  
  Next i
  
End Sub

Sub LoadSegmInformation()
Dim i%

2090
  Segments = NE.SegmentTableEntryCount
  ReDim Segs(Segments)
  
  Seek hVBFile, MZ.OffsetToNE + NE.SegmentTableOffset + 1
  
 'Load SegmentTable Data
  For i = 1 To Segments
    Get hVBFile, , Segs(i)
  Next i
  
End Sub

Sub sub011F(p0112 As Integer)
Dim l0114%
Dim l0118 As Integer
Dim l011A As Long
Dim l011E As T1CBD
Dim l0120%
Dim l0122 As Integer
Dim l0124 As String
Dim l0126() As Long

2100
  curSegID = 0: ApplyReloc p0112
  Erase gv0680: gv06B2 = 0
  ReDim gv0680(Segments, 255)
  ReDim l0126(255) As Long
  For l0114 = 1 To gv0758
    l011E = gv0726(l0114)
    If l011E.iToken <> 0 Or l011E.M1CC7 <> 2 Then
    Else
      l011A = Uint32(l011E.Size_M13F4)
      gv06B2 = 1: l0124 = Chr$(1)
      Do
        gv06B2 = gv06B2 + 1
        Get hVBFile, curSegOffs + l011A - 2, l0118: l011A = Uint32(l0118)
        l0126(gv06B2) = l011A
        For l0120 = 1 To Len(l0124)
          If l0126(Asc(Mid$(l0124, l0120, 1))) > l011A Then Exit For
        Next l0120
        l0124 = Left$(l0124, l0120 - 1) & Chr$(gv06B2) & Mid$(l0124, l0120)
      Get hVBFile, , l0118
      l011A = Uint32(l0118)
      Loop Until l0118 = &HFFFF
      For l0120 = 1 To Len(l0124)
        l0122 = Asc(Mid$(l0124, l0120, 1))
        gv0680(l011E.M1CB1, l0120) = ToSInt16(l0126(l0122))
      Next l0120
      gv0680(l011E.M1CB1, 0) = Len(l0124)
    End If
  Next l0114
End Sub

Function GetSegOffset(p0130 As Integer) As Long
2110
  GetSegOffset = PageToOffset(Segs(p0130).SegOffset, SegmentAlign)
End Function
