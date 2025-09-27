Attribute VB_Name = "MODULE41"
' Module4
Option Explicit
Dim ResData_offset As Long
Dim ResData_size As Long
Dim OffsetResourceTable As Long

Const DELETE& = -1

Dim ControlNames As Variant

Sub ArrayAdd(ByRef DynArray, Data)
    Dim size
    size = UBound(DynArray)
   
   'Extent Array
    ReDim Preserve ControlNames(size + 1)
   
   'Set data
    ControlNames(size) = Data

End Sub
Function GetSegmentOffset(Segment As Integer) As Long
  GetSegmentOffset = IntToULng(Segs(Segment).Offset) * NE_Alignment
End Function

Sub LoadEntryTable()
   Dim EntryType As Integer
   Dim EntryTableOffset As Long
   Dim EntryTableRecord As EntryTableStruct
   Dim EntryTableRecord2 As EntryTableStruct2
     
     EntryTableOffset = MZ.NE_Hdr + NE.ENTRYTABLE
     Objects(ENTRYTABLE).Offset = EntryTableOffset
     
     Get hInFile, EntryTableOffset + 1, FIXED_ONE_BYTE_STRING
     
     EntryType = Asc(FIXED_ONE_BYTE_STRING)
     
     While EntryType
       
       Get hInFile, , FIXED_ONE_BYTE_STRING
       
       Select Case Asc(FIXED_ONE_BYTE_STRING)
       
          Case 0
          
          Case &HFF
            Seek hInFile, Seek(hInFile) + EntryType * Len(EntryTableRecord)
          
          Case Else
            Seek hInFile, Seek(hInFile) + EntryType * Len(EntryTableRecord2)
          
       End Select
       
       Get hInFile, , FIXED_ONE_BYTE_STRING
       
       EntryType = Asc(FIXED_ONE_BYTE_STRING)
     
     Wend
     
     Objects(ENTRYTABLE).size = Loc(hInFile) - EntryTableOffset

End Sub

Sub LoadImportTable()
   Dim ImportTbl_Offset As Long
   Dim ModRefTbl_Offset As Long
   Dim ModRefData As Integer '
   Dim iModTbl As Integer '
   Dim MaxModRefData As Integer
       
   ModRefTbl_Offset = MZ.NE_Hdr + NE.ModuleReferenceTable + 1
   ImportTbl_Offset = MZ.NE_Hdr + NE.IMPORTTABLE + 1

   Objects(IMPORTTABLE).Offset = ImportTbl_Offset - 1

   For iModTbl = 1 To NE.ModuleTableEntryCount

      Get hInFile, ModRefTbl_Offset, ModRefData
      ModRefTbl_Offset = ModRefTbl_Offset + 2
      
      Get hInFile, ImportTbl_Offset + ModRefData, FIXED_ONE_BYTE_STRING
    ' Skip over namestring
      ModRefData = ModRefData + Asc(FIXED_ONE_BYTE_STRING)
   
    ' Max MaxModRefData
      If ModRefData > MaxModRefData Then MaxModRefData = ModRefData
   
   Next iModTbl

   Objects(IMPORTTABLE).size = MaxModRefData
   
End Sub

Sub LoadModuleReferenceTable()
    Objects(MODULE_REF_TABLE).Offset = MZ.NE_Hdr + NE.ModuleReferenceTable
    Objects(MODULE_REF_TABLE).size = NE.ModuleTableEntryCount * 2
End Sub

Sub LoadNonResidentNameTable()
  
  Dim size As Integer
  
  Dim Offset As Long
  Offset = NE.NonResidentNameTable
  Objects(NonResidentNameTable).Offset = Offset

  Get hInFile, Offset + 1, FIXED_ONE_BYTE_STRING
  
  size = Asc(FIXED_ONE_BYTE_STRING)
  
  While size
    Get hInFile, Seek(hInFile) + size + 2, FIXED_ONE_BYTE_STRING
    size = Asc(FIXED_ONE_BYTE_STRING)
  Wend
  
  Objects(NonResidentNameTable).size = Loc(hInFile) - Offset

End Sub

Sub LoadResData()
Dim i%
Dim MaxOffset As Long
Dim Filepos As Long
  For i = 1 To ResObjsIndex
    SeekToResData i
    
    If ResData_size Then
      
      If MaxOffset = 0 Then
        MaxOffset = ResData_offset: Filepos = ResData_offset + ResData_size
      
      ElseIf ResData_offset < MaxOffset Then
        MaxOffset = ResData_offset
      
      ElseIf ResData_offset >= Filepos Then
        Filepos = ResData_offset + ResData_size
      End If

    End If
  Next
  
  Objects(RES_DATA2).Offset = MaxOffset - 1
  Objects(RES_DATA2).size = Filepos - MaxOffset
  
  If Filepos > LOF(hInFile) Then
     Filepos = LOF(hInFile) - Filepos + 1
    Objects(RES_DATA2).Gap = Filepos
  End If

End Sub

Sub LoadResidentNameTable()
Dim Offset As Long
Dim Readbuff As String
  
  Offset = MZ.NE_Hdr + NE.ResidentNameTable

  Objects(ResidentNameTable).Offset = Offset

  Get hInFile, Offset + 1, FIXED_ONE_BYTE_STRING
  
  While Asc(FIXED_ONE_BYTE_STRING)
    
   'Dummy read ResidentNames
    Readbuff = Space$(Asc(FIXED_ONE_BYTE_STRING) + 2)
    Get hInFile, , Readbuff
    Get hInFile, , FIXED_ONE_BYTE_STRING
    
  Wend
  Objects(ResidentNameTable).size = Loc(hInFile) - Offset
  
End Sub

Private Sub LoadResourceTable()
Dim ResourceTableBuff As ResourceTableRootType
Dim Childs As Integer
Dim ResType_ As Integer
Dim Filepos As Long
Dim TypeNameStrings As Integer
  Objects(RESOURCETABLE).Offset = MZ.NE_Hdr + NE.RESOURCETABLE
  
' RESOURCETABLE comes before ResidentNameTable ?...
  If NE.RESOURCETABLE < NE.ResidentNameTable Then
    
  ' Seek to ResTable Offset
    OffsetResourceTable = MZ.NE_Hdr + NE.RESOURCETABLE + 1
   
  ' Get Alignment shift count for resource data.
    Get hInFile, OffsetResourceTable, Res_Align_Raw
    Res_Align = 2 ^ Res_Align_Raw
    
    
    ReDim ResRootTree(RES_TYPES), ResTypesCount(RES_TYPES): ResRootCount = 0
    ReDim ResObjs(0): ResObjsIndex = 0
    Do
      
      Filepos = Loc(hInFile) + 2
      Get hInFile, , ResourceTableBuff

    ' Reached End?
      If ResourceTableBuff.Type_ID_and_Offset = 0 Then Exit Do
      
    ' Count resource
      ResRootCount = ResRootCount + 1
    
    ' offset relative to the beginning of the resource table (if the high-order bit is set (8000h))
      ResType_ = ResourceTableBuff.Type_ID_and_Offset Xor &H8000
      
      If ResType_ < 0 Then
      ' ...it is an offset to the type string
      ' STOP Not Handled
        BeepSound
      Else
      
        ResTypesCount(ResType_) = ResObjsIndex
        ResRootTree(ResType_) = ResourceTableBuff
        
        Childs = ResourceTableBuff.Childs
        If Childs Then
          
         'Enlarge ResObjs to be big enough for Childs
          ReDim Preserve ResObjs(ResObjsIndex + Childs)
          
          While Childs > 0
            
           'Count absolute data
            ResObjsIndex = ResObjsIndex + 1

            Get hInFile, , ResObjs(ResObjsIndex)
            
           
          ' String Type?
            If ResObjs(ResObjsIndex).ResourceID > 0 Then
              
              If TypeNameStrings = 0 Then
                TypeNameStrings = ResObjs(ResObjsIndex).ResourceID
              
              ElseIf ResObjs(ResObjsIndex).ResourceID > TypeNameStrings Then
                TypeNameStrings = ResObjs(ResObjsIndex).ResourceID
              
              End If

          ' Is this the Last?
            ElseIf ResObjs(ResObjsIndex).ResourceID = 0 Then
            
            ' Rewind Childscounter and ResObjsIndex by 1
              ResRootTree(ResType_).Childs = ResRootTree(ResType_).Childs - 1
              ResObjsIndex = ResObjsIndex - 1
            
            End If
            
            Childs = Childs - 1
          Wend
        End If 'of handling Res Childs
      
      End If 'of the case ResType_ is a offset to the type string
     
    Loop 'Childiteration Loop

  End If

' Store size of ResTable
  Objects(RESOURCETABLE).size = Filepos - OffsetResourceTable - 1
  Objects(EXTRA_RES).Offset = Filepos
  
  If TypeNameStrings Then
    Get hInFile, OffsetResourceTable + TypeNameStrings, FIXED_ONE_BYTE_STRING
    TypeNameStrings = TypeNameStrings + 2 + Asc(FIXED_ONE_BYTE_STRING)
  End If
  
  Objects(EXTRA_RES).size = TypeNameStrings

End Sub

Sub LoadSegmInformation()
Dim i
  ReDim Segs(Segments)
  
  Objects(SEGMENTTABLE).Offset = MZ.NE_Hdr + NE.SegmentTableOffset
  Objects(SEGMENTTABLE).size = Segments * Len(Segs(1))
  
  Seek hInFile, MZ.NE_Hdr + NE.SegmentTableOffset + 1
  
'Load Segment Data
  For i = 1 To Segments
    Get hInFile, , Segs(i)
  Next i

End Sub

Sub LoadVBCODE()
   Dim iSegs%
   Dim MaxOffset As Long
   Dim Filepos As Long
   Dim Buff%
   Dim l0186 As VBCODEStruct

     For iSegs = 1 To Segments
       
       SeekToSeg iSegs
       
       If CurrentSegmentSize Then
         
         If MaxOffset = 0 Then
           MaxOffset = CurrentSegmentOffset: Filepos = CurrentSegmentOffset + CurrentSegmentSize
         
         ElseIf CurrentSegmentOffset < MaxOffset Then
           MaxOffset = CurrentSegmentOffset
         
         ElseIf CurrentSegmentOffset >= Filepos Then
           Filepos = CurrentSegmentOffset + CurrentSegmentSize
           
           Get hInFile, Filepos, Buff
           
           If Buff > 0 Then
             Filepos = Loc(hInFile) + Buff * Len(l0186)
           
           End If

         End If
       End If
     Next

   Objects(VBCODE).Offset = MaxOffset - 1
   Objects(VBCODE).size = Filepos - MaxOffset - 1

End Sub

Sub SeekToSeg(Segm As Integer)
  
  CurrentSegment = Segm

  If CurrentSegment Then
    
    CurrentSegmentOffset = 1 + GetSegmentOffset(Segm)
    CurrentSegmentSize = CLng(Segs(Segm).size) And &HFFFF&
    
    If CurrentSegmentSize = 0 Then
      If Segs(Segm).Offset Then CurrentSegmentSize = &H10000
    End If
    
    Seek hInFile, CurrentSegmentOffset
  
  Else
    
    CurrentSegmentOffset = 0
    CurrentSegmentSize = 0

  End If
End Sub

Sub ReadFileData(FileName As String)
Dim iObj
  NE.Signature = 0
  
  If hInFile Then
    Close hInFile
  Else
    hInFile = FreeFile
  End If

  Open FileName For Binary As hInFile
  
  Erase Objects
  
 'Read NE-Header
  Get hInFile, , MZ
  
  Objects(MZ_HEADER).size = MZ.NE_Hdr
  
 'Check MZ-signature
  If (MZ.Signature <> MZ_MAGIC) Or (MZ.RelocTable < &H40) Then Exit Sub
  
  Objects(NE_Header).Offset = MZ.NE_Hdr
  Objects(NE_Header).size = Len(NE)
  
 'Read NE-Header
  Get hInFile, MZ.NE_Hdr + 1, NE
  
 'Check NE-signature
  If NE.Signature <> NE_Hdr Then
    NE.Signature = 0
    Close hInFile
    Exit Sub
  End If
  
  NE_Alignment = 2 ^ NE.MiscFlags
  Segments = NE.SegmentTableEntryCount
  
  LoadSegmInformation
  LoadResourceTable
  LoadResidentNameTable
  LoadModuleReferenceTable
  LoadImportTable
  LoadEntryTable
  LoadNonResidentNameTable

  LoadVBCODE
  LoadResData
  
  For iObj = 1 To VBCODE - 1
  ' Estimate overlapsing(+) or Gap(-)
  ' Calc only gaps [ Next > (Current + Current.size) ]
    If Objects(iObj + 1).Offset > (Objects(iObj).Offset + Objects(iObj).size) Then
      
    ' Gap = (Current + Current.size) - Next
      Objects(iObj).Gap = (Objects(iObj).Offset + Objects(iObj).size) - Objects(iObj + 1).Offset
    
    End If

  Next

End Sub

Sub Rebuild_NEHeader(FileHeader As NE_Struct, ResData As String)
   Dim item As Integer '
   Dim l00F6 As Integer
   Dim itemOffset As Long
   Dim itemOverlaps As Long
   Dim itemGap As Integer
   Dim itemSize As Long
   Dim ResDataSize As Integer
   Dim Rand_UInt As Long
  
  ResDataSize = Len(ResData)
  Randomize

' Iterate through all NE-HeaderObjects
  For item = MZ_HEADER To RES_DATA2
    
    itemOverlaps = itemOffset - Objects(item).Offset
    itemSize = Objects(item).size
    
    Select Case item
    
    Case SEGMENTTABLE      '3
      FileHeader.SegmentTableOffset = FileHeader.SegmentTableOffset + itemOverlaps
    
    Case RESOURCETABLE    '4
      FileHeader.RESOURCETABLE = FileHeader.RESOURCETABLE + itemOverlaps
      
      If ResDataSize Then
      ' Reached end?
        If ResRootTree(RES_TABLE_TYPE).Type_ID_and_Offset = 0 Then
          ResRootTree(RES_TABLE_TYPE).Type_ID_and_Offset = RES_TABLE_TYPE Or &H8000
          ResRootCount = ResRootCount + 1
        End If
      
      ' If there are no Resource of this Type...
        If ResRootTree(RES_TABLE_TYPE).Childs = 0 Then
          ResRootTree(RES_TABLE_TYPE).Childs = 1
          ResTypesCount(RES_TABLE_TYPE) = ResObjsIndex
          ResObjsIndex = ResObjsIndex + 1
          If ResObjsIndex > UBound(ResObjs) Then
            ReDim Preserve ResObjs(ResObjsIndex)
          End If
        End If
      End If

      itemSize = ResRootCount * Len(ResRootTree(1)) + ResObjsIndex * Len(ResObjs(1)) + 4
      ResTblSizeDiff = itemSize - Objects(item).size
    
    Case EXTRA_RES              '5
      If itemSize = 0 Then itemSize = 1
    
    Case ResidentNameTable              '6
      FileHeader.ResidentNameTable = FileHeader.ResidentNameTable + itemOverlaps
    
    Case MODULE_REF_TABLE    '7
      FileHeader.ModuleReferenceTable = FileHeader.ModuleReferenceTable + itemOverlaps
    
    Case IMPORTTABLE         '8
      FileHeader.IMPORTTABLE = FileHeader.IMPORTTABLE + itemOverlaps
    
    Case ENTRYTABLE          '9
      FileHeader.ENTRYTABLE = FileHeader.ENTRYTABLE + itemOverlaps
    
    Case NonResidentNameTable '10
      FileHeader.NonResidentNameTable = FileHeader.NonResidentNameTable + itemOverlaps
    
    Case VBCODE    '11
      VBCodeOverlaps = (itemOverlaps + NE_Alignment - 1) \ NE_Alignment
      itemOffset = Objects(item).Offset + VBCodeOverlaps * NE_Alignment
    
    Case RES_DATA2  '12
      
      'AlignOffset
      RES_DATAOverlaps = (itemOverlaps + Res_Align - 1) \ Res_Align
      itemOffset = Objects(item).Offset + RES_DATAOverlaps * Res_Align
      
     'Create Rand
      Rand_UInt = Rnd * &H10000
      
      FileHeader.CRC32_L = UintToInt(Rand_UInt)
      Rand_UInt = FileHeader.CRC32_L And &H7FFF
      
      If ResDataSize Then
        FileHeader.LinkerVer = FileHeader.LinkerVer + &H100
        
      ' If there are Resource of this Type...
        If ResRootTree(RES_TABLE_TYPE).Childs Then
          l00F6 = ResTypesCount(RES_TABLE_TYPE) + 1
          itemSize = itemSize - ResObjs(l00F6).size * Res_Align
          ResObjs(l00F6).Offset = 0
          ResObjs(l00F6).size = (ResDataSize + Res_Align - 1) \ Res_Align
          '0010h = MOVEABLE  Resource is not fixed.
          '0020h = PURE      Resource can be shared.
          '0040h = PRELOAD   Resource is preloaded.
          
          ResObjs(l00F6).FlagWord = &H1C30
          ResObjs(l00F6).ResourceID = &H8001
        End If
        itemSize = itemSize + ResDataSize
        
        For l00F6 = ResDataSize To ResDataSize \ 2 Step -1
          Rand_UInt = (Rand_UInt * 3 + Asc(Mid$(ResData, l00F6))) And &H7FFF&
        Next
        
      End If
      
      FileHeader.CRC32_H = UintToInt(Rnd(-Rand_UInt) * &H10000)
    
    End Select
    
    
    itemOverlaps = itemOffset - Objects(item).Offset
    itemSize = itemSize + (itemSize And 1)
    itemGap = itemSize - Objects(item).size

    Objects(item).Overlaps = itemOverlaps
    Objects(item).Gap = itemGap

  ' Next
    itemOffset = itemOffset + itemSize

  Next

End Sub

Sub SeekToResData(ResItem As Integer)
  
  ResData_offset = 1 + IntToULng(ResObjs(ResItem).Offset) * Res_Align
  
  ResData_size = IntToULng(ResObjs(ResItem).size) * Res_Align
  
  Seek hInFile, ResData_offset

End Sub

Sub RebuildControls(ResID)

''Store cur filepos
'   Dim oldFileLoc&
'   oldFileLoc = Loc(hInFile)
'
'     ' For All RC_DATA Resources...
'       iSeg = ResTypesCount(RES_TYPE_RC_DATA)
'       iMax = ResRootTree(RES_TYPE_RC_DATA).Childs
'       For i = 1 To iMax
'
'         lRes_ID = ResObjs(iSeg + i).ResourceID
'
'         If (lRes_ID = ResID) Then
'           Exit For
'         End If
'
'       Next
'       SeekToResData (i)
'
'   Debug.Print Hex(Loc(hInFile))
'
''   ReDim ControlNames(0)
''         ' Log Form/ControllNames for deprotecting(recreating FromResource)
''           If (Typ = VB_Form) Then
''              ArrayAdd ControlNames, Buff
''
''           ElseIf (Typ = VBX) Then
''              Dim ControlNamesCount%
''              ControlNamesCount = ControlNamesCount + 1
''              ArrayAdd ControlNames, "iControl" & Format(ControlNamesCount, "00")
''           End If
''
'
'  'restore filepos
'   Seek hInFile, oldFileLoc + 1

End Sub



Sub DelFrmData(hOutFile As Integer, OutFileDataSize As Long)
   Dim Buff As String
   Dim Typ As Integer
   Dim l00DC%
   Dim VB_Dir As VB_Dir_Struct
   Dim VB_Main As VB_Main_Struct
   Dim CurFilePos As Long
   
    'Copy ProjectnameStruct
     CurFilePos = Seek(hOutFile)
     Get hInFile, , VB_Dir: Put hOutFile, , VB_Dir
     
    'Copy program name
     Buff = Space$(VB_Dir.NameSize_M0E57)
     Get hInFile, , Buff: Put hOutFile, , Buff
     
    'Copy application title
     Buff = Space$(VB_Dir.otherSize_M0E62)
     Get hInFile, , Buff: Put hOutFile, , Buff
     
     
     Do
       
       Get hInFile, , VB_Main
                   
                   Typ = Asc(VB_Main.Type_M0E7C)
       Select Case Typ
       
       Case 0
       ' Write Record
         Put hOutFile, , VB_Main
         Exit Do
       
       '67 'C', 70 'F'
       Case VBX, VB_Form
        
       ' Read Name
         Buff = Space$(Asc(VB_Main.length_M07D1))
         Get hInFile, , Buff
         
         If ProtectionEnabled Then
         ' Delete FormName
           If Typ = VB_Form Then
             VB_Main.length_M07D1 = Chr$(0)
             Buff = ""
           End If
         
         Else
         
         ' Restore FormName
           If (Typ = VB_Form) And (VB_Main.length_M07D1 = Chr$(0)) Then
           ' TODO:Rebuild Controls
             RebuildControls (VB_Main.ResIDAssoc)
           
           ' Create new FormName
             Dim FormCount%
             Buff = "Form" & Format(FormCount, "00") & ".FRM" & Chr$(0)
             FormCount = FormCount + 1
             
           ' Set Len for new FormName
             VB_Main.length_M07D1 = Chr$(Len(Buff))
             
           End If
         
         End If
         
       ' 88 X
       Case VB_Control_Type
         
        'If it's a usercontroll read name into buff
         If Asc(VB_Main.length_M07D1) = &HFF Then
           Get hInFile, , FIXED_ONE_BYTE_STRING
           Buff = Space$(Asc(FIXED_ONE_BYTE_STRING))
           Get hInFile, , Buff
           Buff = FIXED_ONE_BYTE_STRING & Buff
           
         Else
           
           Buff = ""
           
         End If
       
       Case Else
         Exit Do
       End Select
       
     ' Write Record
       Put hOutFile, , VB_Main
       Put hOutFile, , Buff
       
       
       Select Case Typ
       
       '70 F ,88 X
       Case VB_Form, VB_Control_Type
         Get hInFile, , l00DC: Put hOutFile, , l00DC
         If l00DC > 0 Then
           Buff = Space$(l00DC * 4)
           Get hInFile, , Buff: Put hOutFile, , Buff
         End If
       End Select
       
       
     Loop
     
     OutFileDataSize = OutFileDataSize - (Seek(hOutFile) - CurFilePos)
     
     Buff = String$(OutFileDataSize, 0)
     Put hOutFile, , Buff
End Sub

Sub WriteFileData(FileName As String, FileNameRes As String)
   Dim hOutFile As Integer
   Dim Filebuff As String
   Dim CopyBuff% '
   Dim iSeg%
   Dim iRes_data_Offset As Long
   Dim iRes_Objs_index As Integer
   Dim i
   Dim FileHeaderOutBuff As NE_Struct
   Dim Seg_Item As SegmentStruct
   Dim Offset As Integer
   Dim l0128 As VBCODEStruct
   Dim lResRoot As ResourceTableRootType
   Dim lResChilds As ResourceChildType
   Dim lSegs() As Integer
   Dim iMax As Integer
   Dim lRes_ID As Integer
     
     If hInFile = 0 Then BeepSound: Exit Sub
   
   On Error GoTo 0
   
  ' Prepare undo Segmentswitch(VBPRO Protection)
    Dim PlaceSeg3AtBeginning As Boolean
    If ProtectionEnabled Then
      If (NE.Initial_CS = 1) Then
'        PlaceSeg3AtBeginning = True
'        NE.Initial_CS = 3
      End If
    Else
     'No Protection (Protection should be removed) and
      If (NE.Initial_CS = 3) Then
       '...changed order detected(StartSegment is 3)
         If vbYes = MsgBox("(This will move NE.Segment 3 to the beginning.)" & vbCrLf & "If you select 'yes' the file will not be execuable/runable anymore - but therefor you will be able to decompile it with VBDis.", vbQuestion + vbYesNo, "Undo VBPro-Protection?") Then
            PlaceSeg3AtBeginning = True
            NE.Initial_CS = 1
         End If
      End If
    End If
   
   
   
   ' Recreate NE-Header
     FileHeaderOutBuff = NE
     Rebuild_NEHeader FileHeaderOutBuff, FileNameRes
     
     
   ' Open out File
     hOutFile = FreeFile
     Open FileName For Binary As hOutFile
     
    'Fill MZ-Header with space
     Filebuff = Space$(Objects(MZ_HEADER).size)
     Get hInFile, 1, Filebuff
     Put hOutFile, 1, Filebuff
   
   
   ' Write NE-Header
     Put hOutFile, , FileHeaderOutBuff
     
   
   ' Write Segment Table
     ReDim lSegs%(Segments)
     
     Offset = Objects(VBCODE).Offset \ NE_Alignment + VBCodeOverlaps
     
   ' Create order on how to Process Segment
   ' Depending on PlaceSeg3AtBeginning it's "1,2,3,4,5..." or "3,1,2,4,5..."
     Dim item
     Dim SegmentToProcess As New Collection
     For iSeg = 1 To Segments
      item = iSeg
      If PlaceSeg3AtBeginning Then
         Select Case iSeg
            Case 1:    item = 3
            Case 2:    item = 1
            Case 3:    item = 2
         End Select
      End If
      SegmentToProcess.Add item
     Next
     
     For Each item In SegmentToProcess
       iSeg = item
'     For iSeg = 1 To Segments
     ' set & seek to Segment
       Seg_Item = Segs(iSeg)
       SeekToSeg iSeg
       
     ' is there data inside this seg
       If CurrentSegmentSize Then
         
       ' set New Offset
         Seg_Item.Offset = Offset
         
        'Process Relocationinfo
         If Seg_Item.Flags And RELOCINFO Then
'        If Segs(iSeg).Flags And RELOCINFO Then
           
           Get hInFile, CurrentSegmentOffset + CurrentSegmentSize, CopyBuff
   
           If CopyBuff > 0 Then CurrentSegmentSize = CurrentSegmentSize + CopyBuff * Len(l0128)
   
           CurrentSegmentSize = CurrentSegmentSize + 2
   
         End If
   
         CopyBuff = (CurrentSegmentSize + NE_Alignment - 1) \ NE_Alignment
   
        'Calc new offset
         Offset = Offset + CopyBuff
   
       Else
        
       ' Segment is empty
         CopyBuff = 0
   
       End If
   
     ' write Segment into file
       Put hOutFile, , Seg_Item
   
       lSegs(iSeg) = CopyBuff
   
     Next
     
     
     
   ' Write ResHeader
     Put hOutFile, Objects(RESOURCETABLE).Offset + Objects(RESOURCETABLE).Overlaps + 1, Res_Align_Raw
   
     
     
    ' Delete FormName RC_DATA Resource

'8001 VBMain
'8002 VBMain_add
'8004 Form1
'8005 Form1_Controlname <- will get deleted
'8006 Form1
'8007 Form1_Controlname <- will get deleted
'8008 ...

     ' For All RC_DATA Resources...
       iSeg = ResTypesCount(RES_TYPE_RC_DATA)
       iMax = ResRootTree(RES_TYPE_RC_DATA).Childs
       For i = 1 To iMax
         
         lRes_ID = ResObjs(iSeg + i).ResourceID
         
        'Cut all uneven Res bigger than 4 like 5,7,9...
         '
         If ProtectionEnabled And (lRes_ID And 1) Then
           If (lRes_ID And &H7FFF) > 4 Then lRes_ID = 0
         End If
   
       ' ResID=0 marks the end
         If lRes_ID = 0 Then
           
         ' Mark Resource for delete
           ResObjs(iSeg + i).Reserved = DELETE
           iMax = iMax - 1
   
         End If
   
       Next
   
   
   
   ' Set start offset of resourcesdata item for saving (store it in .Reserved)
     iRes_data_Offset = Objects(RES_DATA2).Offset \ Res_Align + RES_DATAOverlaps
     For iSeg = 1 To ResObjsIndex
      
      'Add all not delete (Reserved = -1 means deleted) and
      'with Offset not 0 <- because they were delete in a prev. Session
       If (ResObjs(iSeg).Reserved = 0) Then
         
        'If resource has a size
         If ResObjs(iSeg).size Then
          
         ' write startoffset in .Reserved [for later use]
           ResObjs(iSeg).Reserved = iRes_data_Offset
'           If (ResObjs(iSeg).Offset > 0) Then
           iRes_data_Offset = iRes_data_Offset + ResObjs(iSeg).size
 '          Else
  '           Stop
   '        End If
           
         Else
           Stop
         End If
       Else
       ' Set Reserved back to 0
         ResObjs(iSeg).Reserved = 0
       End If
     Next
   
   
   
   ' Save resdata
   
    'Go through all Res_Types
     For iSeg = 1 To RES_TYPES
       
     ' If there are Resource of this Type...
       If ResRootTree(iSeg).Childs Then
         
       ' Reached the end of Resources?
            lResRoot = ResRootTree(iSeg)
         If lResRoot.Type_ID_and_Offset > 0 Then
            lResRoot.Type_ID_and_Offset = lResRoot.Type_ID_and_Offset + ResTblSizeDiff
         End If
         
        'write Res_Rootenty
         Put hOutFile, , lResRoot
         
         iRes_Objs_index = ResTypesCount(iSeg)
         
       ' Do for all Resource of this Type...
         For i = 1 To ResRootTree(iSeg).Childs
           
         ' Reached the end of Resources?
              lResChilds = ResObjs(iRes_Objs_index + i)
           If lResChilds.ResourceID > 0 Then
              lResChilds.ResourceID = lResChilds.ResourceID + ResTblSizeDiff
           End If
         
         ' Set Offset
           lResChilds.Offset = lResChilds.Reserved
         
         ' set '.Reserved' back to 0
           lResChilds.Reserved = 0
           
         ' Save Res Child Table
           Put hOutFile, , lResChilds
         
         Next
         
       End If
   
     Next
   
   
   
   ' write 00 00
     iSeg = 0: Put hOutFile, , iSeg
   
   ' Copy Extra Res (like ResID-stringnames...) from source into dest
     CopyBuff = Objects(EXTRA_RES).size
     If CopyBuff Then
       Filebuff = Space$(CopyBuff)
       Get hInFile, Objects(EXTRA_RES).Offset + 1, Filebuff
       Put hOutFile, , Filebuff
     End If
     
   ' copy/Write ResidentNameTable
                                        CopyBuff = Objects(EXTRA_RES).Gap
                     Filebuff = String$(CopyBuff, 0)
     Put hOutFile, , Filebuff
     
     For iSeg = ResidentNameTable To NonResidentNameTable
     ' Calc Offset difference between segments
                  CopyBuff = Objects(iSeg + 1).Offset - Objects(iSeg).Offset
     ' Calc Overlaps difference between segments
       CopyBuff = CopyBuff + Objects(iSeg + 1).Overlaps - Objects(iSeg).Overlaps
     
     ' Copy 'CopyBuff'-Bytes from Source to Dest File
       Filebuff = Space$(CopyBuff)
       Get hInFile, Objects(iSeg).Offset + 1, Filebuff
       
       Dim vbRunPos%
       vbRunPos = InStr(1, Filebuff, "vBrUn300", vbTextCompare)
       If vbRunPos Then
         If ProtectionEnabled Then
           Mid(Filebuff, vbRunPos) = "vbrun300"
         Else
           Mid(Filebuff, vbRunPos) = "VBRUN300"
         End If
       End If
       
       Put hOutFile, , Filebuff
     Next
     
   
   
   ' copy/write Segments
     Offset = Objects(VBCODE).Offset \ NE_Alignment + VBCodeOverlaps
     
'     For iSeg = 1 To Segments
     For Each item In SegmentToProcess
       iSeg = item

       
       SeekToSeg iSeg
       
       If CurrentSegmentSize Then
         CurrentSegmentSize = CLng(lSegs(iSeg)) * NE_Alignment
         
        'Copy loop if Segment is bigger than &H7000
         If CurrentSegmentSize > &H7000 Then
            Filebuff = Space$(&H7000)
            While CurrentSegmentSize > &H7000
              Get hInFile, , Filebuff: Put hOutFile, , Filebuff
              CurrentSegmentSize = CurrentSegmentSize - Len(Filebuff)
            Wend
         End If
         
        'Copy data with size of CurrentSegmentSize
         Filebuff = Space$(CurrentSegmentSize)
         Get hInFile, , Filebuff: Put hOutFile, , Filebuff
       
       End If
       Offset = Offset + lSegs(iSeg)
     Next
     
     
     
   ' copy/write Resdata  and Filterout form names
     For iSeg = 1 To ResObjsIndex
       SeekToResData iSeg
      
      'is Offset set
       If ResObjs(iSeg).Reserved Then
       
       ' is last segment?
         If iSeg = ResTypesCount(RES_TYPE_RC_DATA) + 1 Then
           DelFrmData hOutFile, ResData_size
         
       ' does it have a size
         ElseIf ResData_offset > 1 Then
           
'         ' Copy loop if Segment is bigger than &H7000
'           Filebuff = Space$(&H7000)
'           While ResData_size > Len(Filebuff)
'             Get hInFile, , Filebuff: Put hOutFile, , Filebuff
'             ResData_size = ResData_size - Len(Filebuff)
'           Wend
           
         ' Copy data with size of ResData_size
           Filebuff = Space$(ResData_size)
           Get hInFile, , Filebuff
            ' is last segment?
              If iSeg = ResTypesCount(RES_TYPE_RC_DATA) + 1 Then
                If ProtectionEnabled Then
                  Do
                    i = InStr(Filebuff, ".FRM")
                    If i = 0 Then Exit Do
                    Mid$(Filebuff, i, 4) = "   ?"
                  Loop
                 End If
                  
              End If
           Put hOutFile, , Filebuff
           
         Else
           
           If Len(FileNameRes) Then
              Put hOutFile, , FileNameRes
           Else
           
              'Create buffer
              Filebuff = String$(ResData_size, Chr(0))
              
              Dim frmcounter&
              'Dim item
              Dim Fillstr$
              Fillstr = ""
              For Each item In Split("dum1 dum2 dum3 dum4 dum5 dum6 dum7 dum8 dum9 duma dumb dumc dumq dumw dume dumr a1 a2 a3 a4 a5 a6 a7 a8 a9 aa aq aw e ar")
               Fillstr = Fillstr & Chr(Len(item) + 2) & item & Format(frmcounter, "00")
              Next
              
              Mid$(Filebuff, 1) = Fillstr
              
              frmcounter = frmcounter + 1
                            
              Put hOutFile, , Filebuff
              
              
           End If
           
           
         End If
         
       End If
     Next
     
     Close hOutFile
     
End Sub

