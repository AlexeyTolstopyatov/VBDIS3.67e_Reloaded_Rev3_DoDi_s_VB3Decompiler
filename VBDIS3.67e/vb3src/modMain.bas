'Attribute VB_Name = "modMain"
Option Explicit
'Option Base 1

' main.txt - global definitions
Type T1267
    M1275 As Integer
    M127E As Integer
    M1287 As Integer
    M128E As String * 9
    M1299 As Integer
    M12A4 As Integer
End Type

Type VB_Main_Struct
    Type As String * 1
    length As String * 1

    ResIDAssoc  As Integer
    M_2_curr As Integer
    M_3_next As Integer
    M_4_sub As String * 1
    M_5_Size As Integer
End Type

Type TControlItem
    Type_3 As Integer
    memEater1 As String * 1

    ResIDAssoc As Integer

    M_2_curr As Integer
    M_3_next As Integer
    M_4_sub As String * 1
    M_5_Size As Integer
    M2548 As Integer
    M2551 As Integer
End Type

Type T1379
    frmOffset As Integer
    M1390 As Integer
    M139C As Integer
    M13A7 As Integer
End Type

Type T13D2
    frmOffset As Integer
    M13DF As Integer
    Type_3 As Integer
    Size_M13F4 As Integer
End Type

Type T1425
    M1432 As Integer
    M143D As Integer
    M1445 As Integer
    M144F As Integer
    M1458 As Integer
    Type_3 As Integer
End Type

Type T1477
    M1432 As Integer
    M1481 As Integer
End Type

Type T149E
    M14AE As String * 3
    Type_3 As String * 1
    M14B9 As String * 1
    M14C1 As Long
End Type

Type T14C9
    mSize6 As Long
    M14E1 As String * 1
End Type

Type T14EB
    Type_3 As Integer
    M14F8 As Integer
    frmOffset As Integer
    M1502 As Integer
    M150C As Integer
    M1516 As Integer
    SubOrFunc As Integer
    PubOrPrivate As Integer
    M1531 As Integer
    M1537 As Integer
    M_5_Size As Integer
    M153D As Integer
    Offs As Integer
    M1551 As Integer
    M155C As Integer
    Reserved(15 To 17) As Integer
    Size As Integer
    M157B As Integer
    M1587 As Integer
    M1592(21 To 22) As Integer
    M159D As Integer
    M15A9 As Integer
    M15B4 As Integer
    M15BE As Long
End Type

Type T15C8
    Type_3 As Integer
    M14F8 As Integer
    frmOffset As Integer
    M1502(3 To 6) As Integer
    M15D5 As Integer
    M15E0(8 To 12) As Integer
    M15E6 As Integer
    M15EF As Integer
    M15F5 As Integer
    M1600(16 To 28) As Integer
    M_4_sub As Integer
    M1607 As Long
End Type

Global Const vbRetryCancel = 5  ' &H5%
Global Const vbAbortRetryIgnore = 2
Global Const vbRetry = 4 ' &H4%
Global Const vbCancel = 2 ' &H2%
Global Const vbAbort = 3 ' &H3%
Global Const vbNo = 7 ' &H7%

Global Const gVersionStr= "3.67"
Global Const gRevisionStr = "e"
Global Const gErrorSituationStr = "Error Situation VB Discompiler"
Global Const gSevereErrorsStr = "Severe errors may cause the Discompiler to crash"
Global Const gInternalErrorsBuggyCodeStr = "Internal problems, the code created may be buggy"
Global Const gLatestVersionAdStr = "Do you have the latest edition of VB Discompiler?"
Global Const gRuntimeErrorStr = "Runtime Error in VB Discompiler"
Global Const gDecompilerNewsStr = "News from VB Discompiler"
Global Const gProEditionAdStr = "This option is available only in the Professional version" '
Global Const gLongNameStr = "Name too long"
'Global Const gc0326 = "For this program you'll need an upgraded VB Discompiler"
'Global Const gc033E = "You may send this program to DoDi to improve VB Discompiler"
'Global Const gc0342 = "Error "
'Global Const gc034A = "Found unknown data structures!"
'Global Const gc034E = "Not a Visual Basic program"
'Global Const gc0356 = "Found an unknown resource!"
'Global Const gc035A = "Found unknown fixups!"
'Global Const gc035E = "Error in Discompiler logic!"
'Global Const gc0366 = "Missing description for "
'Global Const gc036E = "Found an unknown token!"
'Global Const gc0372 = "Unexpected variable reference!"
'Global Const gc0376 = "Found incompatible scopes!"
'Global Const gc037A = "Found incompatible types"
'Global Const gc037E = "Found an unknown collection!"
'Global Const gc0382 = "An already known problem occured"
'Global Const gc0386 = "File not found or wrong version: "


Global Const sInitializing = "Initializing"

Global Const sCombiningCodeAndForms = "Combining forms und code"
Global Const sDecompilingFinish = "Discompilation finished"
Global Const sOpen = "Open "
Global Const sScanning = "Scanning "
Global Const sForms = "Forms"

Global Const gc03E6 = "Segments"
Global Const gc03EA = "Scopes"

Global Const gc03F2 = "Tokens"
Global Const Fixups = "Fixups"
Global Const gc03FA = "Data"
Global Const C_LOCAL = "local "
Global Const C_GLOBAL = "global "
Global Const CurFrmName = "Subroutine calls"
Global Const gc040A = "Global Declarations"
Global Const gc040E = "Declarations in "
Global Const gc0412 = "Error module offset"
Global Const gc0416 = "Rename"
Global Const gc041A = "Startform not found"
Global Const gc041E = "No more errors"
Global Const gc0422 = "Specify type!"
Global Const gc0426 = "Unknown type"
Global Const gc042A = "Different modules selected"
Global Const gc042E = "Unexpected EXE fixup"
Global Const gc0432 = "Token uses no variable"
Global Const gc0436 = "Show Variable"
Global Const gc043A = "Source is not saved in binary format"

' DOS Header - documented structure -> All fields are known
Type TMZHeader
  Signature As Integer
  M19E5 As Integer
  M19EE As Integer
  M19F7 As Integer
  M1A01 As Integer
  M1A0C As Integer
  M1A17 As Integer
  M1A22 As Integer
  M1A2C As Integer
  CRC32 As Integer
  Initial_IP As Integer
  Initial_CS As Integer
  RelocTable As Integer
  M1A5E As Integer
  M1A68(15) As Integer
  OffsetToNE As Long
End Type
' NE Header - documented data structure -> all fields are known
' Mind: INVALIDCASE may tell about something what extremely needed
' or for else -> hard to parse, just because entry table has
' nested variable sized types
' Import table (see https://alexeytolstopyatov.github.io/notes/2025/09/23/ne-imptab.html)
' might have unexpected paddings or run-time calls (function names) instead of DLL names  
Type TNEHeader
  Signature As Integer
  LinkerVer As Integer
  ENTRYTABLE As Integer
  M1AA6 As Integer
  CRC32 As Long
  Flags As Integer
  DataSegmentNumber As Integer
  Initial_heap As Integer
  Initial_stack As Integer
  Initial_IP As Integer
  Initial_CS As Integer
  
  Initial_SP As Integer
  Initial_SS As Integer
  
  SegmentTableEntryCount As Integer
  ModuleTableEntryCount As Integer
  M1AFF As Integer
  SegmentTableOffset As Integer
  RESOURCETABLE As Integer
  ResidentNameTable As Integer
  ModuleReferenceTable As Integer
  IMPORTTABLE As Integer
  NonResidentNameTable As Long
  M1B3B As Integer
  MiscFlags As Integer
  M1B50 As Integer
  M1B5D As Integer
  M1B69 As Integer
  M1B72 As Integer
  Reserved As Integer
  M1B7B As Integer
End Type

Global Const MZ_MAGIC = &H5A4D
Global Const NE_MAGIC = &H454E
Type EntryTableStruct2
  M_4_sub As String * 1
  M1BEB As Integer
End Type

Type EntryTableStruct
  M_4_sub As String * 1
  M1C00 As Integer
  M1C09 As String * 1
  M1BEB As Integer
End Type

Type T1C2C
  M_4_sub As Integer
  segm0 As Integer
  Size_M13F4 As Integer
End Type

Global gv064C() As T1C2C
Global gv067E As Integer
Global gv0680() As Integer
Global gv06B2 As Integer
Global curSegOffs As Long
Global curSegSize As Long
Global curSegID As Integer
Global Segments As Integer
Type VBCODEStruct
  M1CA2 As String * 1
  iToken As String * 1
  Size_M13F4 As Integer ' M13F4's taken from EntryTable 
  M1CB1 As Integer
  M1CB7 As Integer
End Type

Type T1CBD
  Size_M13F4 As Integer
  M1CC7 As Integer
  iToken As Integer
  M1CB1 As Integer
  M1CB7 As Integer
End Type

Global gv0726() As T1CBD
Global gv0758 As Integer
Type TSegment
  SegOffset As Integer
  Size As Integer
  Flags As Integer
  Alloc As Integer
End Type

Global Segs() As TSegment
Global Const RELOCINFO = &H100

' Resident Name record contains Pascal String, 
' that's why TResidentName missing first BYTE
Type TResidentName
  Name As String
  AddData As Integer
End Type

Global ResidentNames() As TResidentName
Global ResidentCount As Integer
Global ResidentNames2() As TResidentName
Global NonResCount As Integer
Global VBRUNs() As String
Global gIn_FileFulPath As String
Global gCurPath As String
Global gIn_FileNameAndExt As String
Global hVBFile As Integer
Global VBver_11 As Integer
Global MZ As TMZHeader
Global NE As TNEHeader
Global SegmentAlign As Integer
Global ResAlignment%
Global bPrjFilesCreated As Integer
Global gWriteOutPut As Integer
Global Tk_ForwardBytes%
Global VBCodeStart&
Global CurOffset As Long
Global VBCodeEnd&
Type T1E64
  Offs As Long
  M1E72 As String
End Type

Global gv094E() As T1E64
Global gv0980 As Integer
Global Cancel As Integer
Global SomeDataBuff() As Integer
Global gv09B6 As Long
Global gv09BA As Integer
Global gVBCodeTokenWordsExE() As Integer
Global gRawVBCodeSize As Long
Global gVBCodeSizeWords As Integer
Global gVBCodeTokenWordsFinal() As Integer
Global FixedByteBuff As String * 1
Type T1EE0
  M1EE9 As String * 2
End Type

Global gv0A3E As T1EE0
Type T1EF4
  M1EFD As Integer
End Type

Global gv0A52 As T1EF4
Type T1F08
  M1EE9 As String * 4
End Type

Type T1F11
  M1F1A As Long
End Type

Type T1F1F
  M1F28 As Integer
  M1F2E As Integer
End Type

Global gv0A90 As T1F08
Global gv0A96 As T1F11
Type T1F46
  M1F52 As Variant
  M1F57 As Integer
  M1F5F As Integer
  M034F As String * 8
  M1F68 As String * 10
  M1F70 As Integer
  M1F78 As Integer
  M1F7F As Integer
  M1F86 As Integer
  M1F8D As Integer
  M1F94 As Integer
End Type

Global gv0B0E As T1F46
Global gv0B68 As Integer
Global gFrmCount2 As Integer
Type T20E6
  M12C8 As String * 1
  M1EFD As Integer
End Type

Global Base1 As Long
Global gv0B8A As Long
Global gv0B8E As Long
Global gv0B92
Global gfrmOffset As Long
Global gCurModStructOffset As Long
Global gv0B9E As String
Global gWindowsDir As String
Global gSystemDir As String
Global VB3Path_2 As String
Global gVB16Location As String
Global VB3Path As String
Global AppPathwithSlash As String
Global CurForm_gv0BC0 As Form
Global gv0BC4 As Integer
Global gv0BC6 As Integer
Global gv0BC8 As Integer
Global Const gc0BCC = 1 ' &H1%
Global gv0BCE As Integer
Global Const gc0BD2 = 1 ' &H1%
Global gv0BD8 As Integer
Global FileOff_PrevToken As Long
Global DisOut As String
Type T2219
  Type_3 As Integer
  M2226 As Integer
End Type

Global gv0BFC As T2219
Type T2237
  M2241 As Integer
  M12C8 As String * 1
  M2249 As String * 1
End Type

Type T2250
  M225B As String
  M2262 As String
End Type

Global gv0C4A As T2250
Global hFile_Out As Integer
Global Listbox_File_Out As Control
Global gDisOutputTxt As String
Global gv0C5E As Integer

Type ResourceTableRootType
  Type_3 As Integer
  Childs As Integer
  Reserved As Long
End Type

Type TResObject
  Type_3 As Integer
  Childs As Integer
  index As Integer
  Name_4 As String
End Type

Global ResObjects0() As TResObject
Global ResRootCount As Integer
Global gv0D04(16) As String
Type ResourceChildType
  Offset As Integer
  Size As Integer
  FlagWord As Integer
  ResourceID As Integer
  Reserved As Long
End Type

Type TResObject1
  Type_3 As Integer
  Size As Integer
  FlagWord As Integer
  ResourceID As Integer
  Offset As Integer
  Name_4 As String
End Type

Global ResObjects1() As TResObject1
Global ResObjsIndex As Integer
Global Res_ID_gv0DC0 As Integer
Global ResObjectsCount As Integer
Global curResObjects1 As TResObject1
Global curResID As Integer
Global CurResOffset As Long
Global ResSize As Long
Global ModulCount As Integer
Global gv0DE4 As Integer
Global NamesLookUpTbl() As Integer
Global Const gc0E18 = 1 ' &H1%
Global Const gc0E1A = 2 ' &H2%
Global Const gc0E1C = 3 ' &H3%
Global Const gc0E20 = 1 ' &H1%
Global Const gc0E22 = 6 ' &H6%
Global Const gc0E24 = 1 ' &H1%
Global Const gc0E26 = 2 ' &H2%
Global Const gc0E28 = 3 ' &H3%
Global Const gc0E2A = 5 ' &H5%
Global Const gc0E2C = 6 ' &H6%
Type TVBForm
  mSubroutinesNameLen As Integer
  frmOffset As Integer
  M24D0 As Integer
  M24DA As Integer
  M_2_curr As Integer
  ResIDAssoc As Integer
  mfrmOffset As Long
End Type

Type T24EF
  TKFrmRes As Integer
  mSubroutinesNameLen As Integer
  Size_M13F4 As Integer
  segm0 As Integer
  frmOffset As Integer
  M2503 As Integer
  M250C As Integer
  M2515 As Integer
  M1CC7 As String
  M251F As Integer
  mCodeSize As Integer
  M2533 As String
End Type

Global gv0EE2(7) As Integer
'Type TControlItem
'  Type_3 As Integer
'  ResIDAssoc As Integer
'  M_2_curr As Integer
'  M_3_next As Integer
'  M_4_sub As String * 1
'  M_5_Size As Integer
'  M2548 As Integer
'  M2551 As Integer
'End Type

Global VBControlNames(256) As String
Global gControlItems(256) As TControlItem
Global VBX_Names_Count As Integer
Global gFormsCount As Integer
Global gControlItemsCount As Integer
Global gVBTitle As String
Global Proj_Title As String
Global Proj_HelpFile As String
Global Start_Form_Index As Integer
Global Proj_IconForm As Integer
Global Start_Form_Res As Integer
Global gv0F88 As Integer
Type TVBDefObject
  id_5 As Integer
  mFrmCount As Integer
  id_add_5 As Integer
  Name_4 As String
End Type

Global gVBObjsDefs(512) As TVBDefObject
Global gv0FD2 As Integer
Type T2619
  id_5 As Integer
  Name_4 As String
End Type

Global gv0FF6() As T2619
Global gControlCount1 As Integer
Type T2643
  M2651 As Integer
  M2658 As Integer
End Type

Global gv1042() As Integer
Global gv1076() As T2643
Type T2679
  Offs As Long
  iTokens As Integer
  M268E As Integer
  M2694 As Integer
  M269A As Integer
  M26A0 As Integer
  M26A6 As String
  M26AC As Integer
  M26B2 As String
  TokenIndex As Integer
End Type

Global ScanedTokens() As T2679
Global ScanedTokensCount As Integer
Global VBDIS_FlagToken2$
Global Const gMASK_0000_0010 = 2 ' &H2%
Global VerDepSig As Integer
Global Const gc114A = "?pmlgcfOTas"
Global Const gc114E = 1 ' &H1%
Global Const gc1150 = 2 ' &H2%
Global Const gc1152 = 3 ' &H3%
Global Const gc1154 = 4 ' &H4%
Global Const gc1156 = 5 ' &H5%
Global Const gc1158 = 6 ' &H6%
Global Const gc115A = 7 ' &H7%
Global Const gc115E = 9 ' &H9%
Global Const gc1164 = "~.[asf14|c"
Global Const gc1168 = 1 ' &H1%
Global Const gc116A = 2 ' &H2%
Global Const gc116C = 3 ' &H3%
Global Const gc116E = 4 ' &H4%
Global Const gc1170 = 5 ' &H5%
Global Const gc1172 = 6 ' &H6%
Global Const gc1174 = 7 ' &H7%
Global Const gc1176 = 8 ' &H8%
Global Const gc1178 = 9 ' &H9%
Global Const gc117A = 10 ' &HA%
Global Const gc117C = "t%&!#@vOT*A$4|"
Global Const gc1180 = "t%&!#@vOT*A$4|"
Global Const gc1184 = "t%&!#@vOT*A$4|1"
Global Const gc1188 = "t%&!#@vOT*A$4|1U"
Global Const gc118C = 1 ' &H1%
Global Const gc119C = 9 ' &H9%
Global Const gc11A0 = 11 ' &HB%
Global Const gc11A4 = 13 ' &HD%
Global Const gc11A6 = 14 ' &HE%
Global Const gc11A8 = 15 ' &HF%
Global gvBase2(31) As Long
Global Const gc11C4 = 128 ' &H80%
Global Const gc11C6 = 64 ' &H40%
Global Const gc11C8 = 32 ' &H20%
Global Const gc11CA = 16 ' &H10%
Global Const gc11D0 = 15 ' &HF%
Global Const gc11D2 = 31 ' &H1F%
Global Const gc11D6 = 128 ' &H80%
Global Const gc11DA = 17 ' &H11%
Global gvConv_bw(15) As Integer
Global gvConv_ff(15) As Integer

Global gv120C As Integer
Type T2977
  M1DF4 As Integer
  M_4_sub As Long
  M2985 As Long
  M2990 As Integer
  M299E As Long
  M29AA As Integer
  M29BA As Integer
  M29C5 As Integer
  ControlTypeName As Integer
  M29E0 As Integer
  M29EF As Integer
  M29F8 As Integer
  M2A02 As String * 1
  M2A0D As String * 1
  M2A19 As String * 1
  M2A26 As Integer
End Type

Global Const gc129C = 256 ' &H100%
Global Const gc129E = 31 ' &H1F%
Global Const gc12BC = 32 ' &H20%
Global gv12BE(32) As String
Global gv12D4(8) As String
Type T2B18
  Name_4 As Integer
  M_4_sub As Long
  M2B27 As String * 1
  M2B30 As String * 1
  M2B3C As Long
  M2B4B As Integer
  M2B57 As Integer
End Type

Type T2B62
  Name_4 As Integer
  M2B72 As Integer
  M2B7B As Integer
  M2B88 As Integer
  M2B95 As Integer
  M_4_sub As Long
End Type

Type TControlData
  M23B8 As Integer
  M1DF4 As Integer
  M2A26 As Integer
  M_4_sub As Long
  ControlTypeName As String
  M2BAC As String
  M29E0 As String
  M2BBB As Integer
  M2BC4 As Integer
  M2A02 As Integer
  M2BCF As Integer
  M2BDA As Integer
  M2BE4 As Integer
  M2A0D As Integer
End Type

Global gKnowControlsData() As TControlData
Global gKnowControlsCount As Integer
Type T2BFB
  Type_3 As Integer
  Name_4 As String
End Type

Global gv1444() As T2BFB
Global gv1476 As Integer
Global gv1478() As Integer
Global gv14AA As Integer
Type TEventName
  Name_4 As String
  M2C33 As String
  M1CC7 As String
End Type

Global g_EventNames() As TEventName
Global gv150E As Integer
Global gv1510() As Integer
Global gv1542 As Integer
Type T2C56
  Type_3 As Integer
  M1D80 As Integer
  Name_4 As String
End Type

Global Const VBDIS_FlagTokenSize = 10837 ' &H2A55%
Global gVB_a_TypeConv(1 To 7) As String
Global gVB_a_DataTypes(1 To 8) As String
Type T2CC5
  TK_Case1_8bit As Integer
  mTokenFlags As Integer
  M2CE2 As Integer
  ForwardWords As Integer
  M2CF2 As Integer
End Type

Global VBdis_String As String
Global ATK9B_VBDat_1 As T2CC5
Global AltToken_9Bit As Integer

' I don't know what it means...
Type TVB9BitToken
  M2D1C As Integer
  M2D22 As Integer
  M2CE2 As Integer
End Type

Type TVBTokens
  Token9Bit(511) As TVB9BitToken
  M2D3C(96) As Integer
End Type

Global VBdis_Struct As TVBTokens
Type T2F11
  AltToken As Integer
  iToken As Integer
  iToken1 As Integer
  iToken2 As Integer
End Type

Global tkMain As T2F11
Global intTokens As Integer


'10837 *3 => 32511 (7EFF)
Type VBDIS_FileType
  FileData(VBDIS_FlagTokenSize) As Integer  '10837 => 0x2A55
'  FileData(1 To VBDIS_FlagTokenSize + 1) As Integer '10837 => 0x2A55
End Type

Global VBDIS_FlagToken As VBDIS_FileType
Global VBDIS_ControlToken As String
Global Writing_main_txt As Integer
Global gFormNames() As String
Global gv78AA() As Integer
Global StrTokens As String
Global Const InvName_1 = 1 ' &H1%
Global Const gc78E4 = 2 ' &H2%
Global Const gc78E6 = 3 ' &H3%
Global Const gc78E8 = 4 ' &H4%
Global Const gc78EA = 5 ' &H5%
Global Const gc78F6 = 11 ' &HB%
Global Const gc78F8 = 12 ' &HC%
Global Const gc78FA = 13 ' &HD%
Global Const gc78FC = 14 ' &HE%
Global Const gc78FE = 15 ' &HF%
Global Const gc7900 = 16 ' &H10%
Global Const gc7902 = 17 ' &H11%
Global Const gc7904 = 18 ' &H12%
Global Const gc7906 = 19 ' &H13%
Global Const gc7908 = 20 ' &H14%
Global gFrmNames2() As Integer
Global fnPreFix(21) As String
Global gIntTokenStrBuff As String
Global gVBCodeTokenAddData() As String
Global gVBCodeTokenWords() As Integer
Global Const gc79BC = -1 ' &HFFFF%
Global Const gc79BE = -2 ' &HFFFE%
Global Const gc79C2 = 224 ' &HE0%
Global gv79C4 As Integer
Global Const gc79C6 = 8 ' &H8%
Global Const gc79C8 = 8 ' &H8%
Global Const gc79CA = 9 ' &H9%
Global Const gc79CC = 10 ' &HA%
Global Const gc79CE = 11 ' &HB%
Global Const gc79D0 = 16 ' &H10%
Global Const gc79D2 = 16 ' &H10%
Global Const gc79D4 = 17 ' &H11%
Global Const gc79D8 = 32 ' &H20%
Global Const gc79DA = 64 ' &H40%
Global Const gc79DE = 96 ' &H60%
Global Const gc79E0 = 128 ' &H80%
Global Const gc79E2 = 160 ' &HA0%
Global Const gc79E4 = 160 ' &HA0%
Global Const gc79E6 = 192 ' &HC0%
Global Const gc79E8 = 224 ' &HE0%
Global Const gc79EE = 3 ' &H3%
Global Const gc7A0C = 225 ' &HE1%
Global Const gc7A0E = 128 ' &H80%
Global Const gc7A10 = 226 ' &HE2%
Global Const gc7A16 = 233 ' &HE9%
Global Const gc7A18 = 231 ' &HE7%
Global Const gc7A1A = 32 ' &H20%
Global Const gc7A1C = 227 ' &HE3%
Global gDecompileStatus As Integer

Global Const b0001_FileNamesSet = 1 ' &H1%
Global Const b0010_ExeLoaded = 2 ' &H2%
Global Const b0100_ScanDone = 4 ' &H4%
Global Const b1000_SubroutinesScanned = 8 ' &H8%
Global Const b0001_0000_DecompilationDone = 16 ' &H10%
Global Const b0010_0000_FormsCombined = 32 ' &H20%

Global ScanOnly As Integer
Global gFrmStruct() As TVBForm
Global FrmFilesCount As Integer
Global curFrmStruct As TVBForm
Global NewFileExt As Integer
Global gSubroutinesNameLen As Integer
Global g_TKStruct() As T24EF
Global gSubroutines As Integer
Global TokenStruct1 As T14EB
Global g_CurSub As Integer
Global gTKFrmRes0 As Integer
Global gv7B30 As Integer
Global gv7B32 As Integer
Global UnloadReason As Integer
Global gv7B36 As Integer
Global gv7B38 As Long
Global MsgBoxResponse As Integer
Global std_p_e As String
'Global Const gc7B50 = 1 ' &H1%
'Global Const gc7B52 = 2 ' &H2%
'Global Const gc7B54 = 4 ' &H4%
'Global Const gc7B56 = 8 ' &H8%
'Global Const gc7B58 = 16 ' &H10%
'Global Const gc7B5A = 32 ' &H20%
'Global Const gc7B5C = 64 ' &H40%
'Global Const gc7B5E = 128 ' &H80%
'Global Const gc7B60 = 256 ' &H100%
'Global Const gc7B62 = 512 ' &H200%
'Global Const gc7B64 = 1024 ' &H400%
'Global Const gc7B66 = 2048 ' &H800%
'Global Const gc7B68 = 4096 ' &H1000%
'Global Const gc7B6A = 8192 ' &H2000%
'Global Const gc7B6C = 16384 ' &H4000%
'Global Const gc7B6E = -32768 ' &H8000%
Global Const gc7B7C = 1 ' &H1%
Global Const gc7B7E = 2 ' &H2%
Global Const gc7B80 = 4 ' &H4%
Global Const gc7B82 = 8 ' &H8%
Global Const gc7B84 = 16 ' &H10%
Global gv7B86
Global gv7B8A As String
Global gWorkDir As String
Global NewPrjName As String
Global gv7B96() As String
Global gv7BC8 As Integer
Global Const EOS_TOKEN = &HE5  '

Global VBStartSeg%
Global CommandlineMode%
Global Const vbSystem = 4
Global Const vbHidden = 2
Global Const vbNormal = 0
Global Const vbDefault = 0
Global Const vbHourglass = 11
Global Const vbDefaultButton3 = 512' (&H200)
Global Const vbIgnore = 5

