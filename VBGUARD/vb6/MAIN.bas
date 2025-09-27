Attribute VB_Name = "MAIN"
' main.txt - global definitions
Global Const VERSION_TYPE = "e"
Global Const ERR_FileNotFound = "File not found"
Global Const ERR_NOT_A_RES_FILE = "Not a RES file"
Global Const ERR_NO_VER_INFO_FOUND = "Found no Version Information"
Global Const SAVE_AS = "Save As ..."
Global Const PROGRAM_EXECUTABLES = "Program|*.exe"
Global Const PROGRAM_RESOURCES = "Resource|*.res"
Global Const ERR_CannotCopytoSameFile = "Cannot copy to same file"
Global Const VERSIONINFO_RES = "VersionInfo Resource"
Global Const PROGRAM_TO_PROTECT = "Program to protect"
Global Const PROGRAM_PROTECED = "Program protected"
Global FIXED_ONE_BYTE_STRING As String * 1


Global CommandlineMode%

Type T0446
  M044F As String * 2
End Type

Global gv004C As T0446
Type T045A
  M0463 As Integer
End Type

Global gv0060 As T045A
Type T046E
  M044F As String * 4
End Type

Type T0477
  M0480 As Long
End Type

Type T0485
  M048E As Integer
  M0494 As Integer
End Type

Global gv009E As T046E
Global gv00A4 As T0477
Type T04AC
  M04B8 As Variant
  M04BD As Integer
  M04C5 As Integer
  M015D As String * 8
  M04CE As String * 10
  M04D6 As Integer
  M04DE As Integer
  M04E5 As Integer
  M04EC As Integer
  M04F3 As Integer
  M04FA As Integer
End Type

Global gv011C As T04AC
Global Const Version = "1.0"
Type T0648
  M065C As Long
  M0668 As Long
  M0678(3) As Integer
  M0687(3) As Integer
  M0696 As Long
  M06A7 As Long
  M06B4 As Integer
  M06C0 As Integer
  M06CC As Long
  M06D8 As Long
  M06E7 As Long
  M06F5 As Long
End Type

Type T0703
  M044F As String * 52
End Type

Type T070D
  ResSize As Integer
  M071F As Integer
  M0727 As String * 16
  M072F As T0648
End Type

Type T073B
  M0480 As Integer
  M0744 As Integer
  M044F(12) As String
End Type

Type T0770
  M072F As T0648
  M077B As Integer
  M0789 As Integer
  M0792(1 To 2) As T073B
  M079A As Integer
  M07A5 As String
  M07B4 As Integer
End Type

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

Type MZ_Struct
  Signature As Integer
  Extra_Bytes As Integer
  Pages As Integer
  Reloc_Items As Integer
  Header_Size As Integer
  Min_Alloc As Integer
  Max_Alloc As Integer
  Initial_SS As Integer
  Initial_SP As Integer
  Check_Sum As Integer
  Initial_IP As Integer
  Initial_CS As Integer
  RelocTable As Integer
  Overlay As Integer
  Fill(15) As Integer
  NE_Hdr As Long
End Type

Type NE_Struct
  Signature As Integer
  LinkerVer As Integer
  ENTRYTABLE As Integer
  M08D2 As Integer
  CRC32_L As Integer
  CRC32_H As Integer
  M07ED As Integer
  M08F4 As Integer
  M08FE As Integer
  M090A As Integer
  Initial_IP As Integer
  Initial_CS As Integer
  M0915 As Integer
  M091F As Integer
  SegmentTableEntryCount As Integer
  ModuleTableEntryCount As Integer
  M093B As Integer
  SegmentTableOffset As Integer
  RESOURCETABLE As Integer
  ResidentNameTable As Integer
  ModuleReferenceTable As Integer
  IMPORTTABLE As Integer
  NonResidentNameTable As Long
  M0977 As Integer
  MiscFlags As Integer
  M098C As Integer
  M0999 As Integer
  M09A5 As Integer
  M09AE As Integer
  OffsetLng As Integer
  M09C0 As Integer
End Type

'Signature_NE                 Signature As Integer
'LinkerVerRev                LinkerVer As Integer
'EntryTable                             EntryTable As Integer
'EntryTableSize                   M08D2 As Integer
'M08EC                        CRC32_L As Integer
'                             M08EC As Integer
'Type                         M07ED As Integer
'AutoDataSegNumber            M08F4 As Integer
'LocalHeapSize                M08FE As Integer
'StackSize                    M090A As Integer
'Initial_IP                   Initial_IP As Integer
'Initial_CS                   Initial_CS As Integer
'Initial_SP                   M0915 As Integer
'Initial_SS                   M091F As Integer
'SegmentTableEntryCount       SegmentTableEntryCount As Integer
'ModuleTableEntryCount        ModuleTableEntryCount As Integer
'Non-ResidentNameTableSize    M093B As Integer
'SegmentTable                 SegmentTableOffset As Integer
'ResourceTable                ResourceTable As Integer
'ResidentNameTable            ResidentNameTable As Integer
'ModuleReferenceTable         ModuleReferenceTable As Integer
'ImportTable                  ImportTable As Integer
'Non-residentNameTable        NonResidentNameTable As Long
'EntryPointCountMoveable      M0977 As Integer
'Alignment                    MiscFlags As Integer
'NumberReservedSegment        M098C As Integer
'TargetOS                     M0999 As Integer
'MiscFlags                    M09A5 As Integer
'FastLoadOffset               M09AE As Integer
'FastLoadSize                 OffsetLng As Integer
'Reserved                     M09C0 As Integer
'WindowsRevision
'WindowsVersion




Global hInFile As Integer
Global MZ As MZ_Struct
Global NE As NE_Struct
Global Const MZ_MAGIC = 23117 ' &H5A4D%
Global Const NE_Hdr = 17742 ' &H454E%
Type EntryTableStruct2
  M07ED As String * 1
  M0A4A As Integer
End Type

Type EntryTableStruct
  M07ED As String * 1
  M0A5F As Integer
  M0A68 As String * 1
  M0A4A As Integer
End Type

Global CurrentSegmentOffset As Long
Global CurrentSegmentSize&
Global CurrentSegment As Integer
Global Segments As Integer
Global NE_Alignment As Integer
Global Res_Align_Raw As Integer
Global Res_Align As Integer
Type T0AD2
  Offset As Integer
  M0AE6 As Integer
  M0AED As Integer
End Type

Type VBCODEStruct
  M0AE6 As String * 1
  M0AED As String * 1
  Offset As Integer
  M0B01 As Integer
  M0B07 As Integer
End Type


Type SegmentStruct
  Offset As Integer
  size As Integer
  Flags As Integer
  Mem As Integer
End Type

Global Segs() As SegmentStruct

Global Const RELOCINFO = &H100

'Enum SegmentStruct_Type
'    TYPE_MASK = &H7    'Segment-type field.
'    CODE = &H0         'Code-segment type.
'    Data = &H1         'Data-segment type.
'    Moveable = &H10    'Segment is not fixed.
'    PRELOAD = &H40     'Segment will be preloaded; read-only if
'                       'this is a data segment.
'    RELOCINFO = &H100  'Set if segment has relocation records.
'    DISCARD = &HF000   'Discard priority.
'End Enum


Type ResourceTableRootType
  Type_ID_and_Offset As Integer    ' This is an integer type if the high-order bit is set (8000h);
                        ' otherwise, it is an offset to the type string,
                        '   the offset is relative to the beginning of the resource table.
                        ' A zero type ID marks the end of the resource type information blocks.
  Childs As Integer
  Reserved As Long
End Type

Type ResourceChildType
  Offset As Integer 'relative to beginning of file.
  size As Integer 'in bytes
  FlagWord As Integer
  ResourceID As Integer
  Reserved As Long
End Type

Global Const RES_TYPE_RC_DATA = 10 ' &HA%
Global Const RES_TABLE_TYPE = 16 ' &H10%
Global Const RES_TYPES = 16 ' &H10%
Global ResRootTree() As ResourceTableRootType
Global ResTypesCount() As Integer
Global ResRootCount As Integer
Global ResObjs() As ResourceChildType
Global ResObjsIndex As Integer

Type NE_Object_Type
  Offset As Long
  size As Long
  Overlaps As Long
  Gap As Integer
End Type

Global Const MZ_HEADER = 1 ' &H1%
Global Const NE_Header = 2 ' &H2%
Global Const SEGMENTTABLE = 3 ' &H3%
Global Const RESOURCETABLE = 4 ' &H4%
Global Const EXTRA_RES = 5 ' &H5%
Global Const ResidentNameTable = 6 ' &H6%
Global Const MODULE_REF_TABLE = 7 ' &H7%
Global Const IMPORTTABLE = 8 ' &H8%
Global Const ENTRYTABLE = 9 ' &H9%
Global Const NonResidentNameTable = 10 ' &HA%
Global Const VBCODE = 11 ' &HB%
Global Const RES_DATA2 = 12 ' &HC%

Global Objects(12) As NE_Object_Type
Global ResTblSizeDiff As Integer
Global VBCodeOverlaps As Integer
Global RES_DATAOverlaps As Long
Global Started As Integer

Type VB_Dir_Struct
  M0E33 As Integer
  M0E3C As Integer
  M0E45 As Integer
  PrjName As String * 9
  NameSize_M0E57 As Integer
  otherSize_M0E62 As Integer
End Type

Type VB_Main_Struct
  Type_M0E7C As String * 1
  length_M07D1 As String * 1
  ResIDAssoc  As Integer
  M_2_curr As Integer
  M_3_next As Integer
  M_4_sub As String * 1
  M_5_Size As Integer
End Type

Global Const VBX = 67 ' &H43%
Global Const VB_Form = 70 ' &H46%
Global Const VB_Control_Type = 88 ' &H58%

Global ProtectionEnabled As Integer
