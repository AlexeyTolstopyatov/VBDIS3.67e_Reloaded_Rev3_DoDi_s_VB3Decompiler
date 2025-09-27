Attribute VB_Name = "MODULE2"
' Module2
Option Explicit
'Const mc0008 = 8 ' &H8%

Const mc00A4 = 32 ' &H20%

Sub ScanModuleToken1(Modul As Integer)
4000
  
  If gWriteOutPut Then Save1File
  
  curFrmStruct = gFrmStruct(Modul)
  NewFileExt = Modul
  gSubroutinesNameLen = curFrmStruct.mSubroutinesNameLen
  
  load1File
  
End Sub

'3Refs (Calls from elsewhere)
Sub DecompileSubroutine(Subroutine As Integer)
4010
  g_CurSub = Subroutine
  gTKFrmRes0 = g_TKStruct(Subroutine).TKFrmRes
  
  Get hVBFile, VBAbsOffs(gTKFrmRes0), TokenStruct1
  
  If g_TKStruct(Subroutine).segm0 Then
    
    SeekToSegment g_TKStruct(Subroutine).segm0
    
    DecompSubroutine
    
  End If
  
End Sub


Sub setPrjFilesCreatedTrue()
4020
  bPrjFilesCreated = True
End Sub

Sub sub0178()
4030
End Sub

Sub load1File()
Dim i%
Dim frmFileN As String
Dim Disp_Lvl As Integer
Dim IntBuff%
Dim File1_CodeSizeWords%
Dim l006E%
 
Dim l0096 As Integer
4040
  If NewFileExt <= 0 Then Exit Sub
  
  
  frmFileN = ModulFileNames(NewFileExt)
  
                  'Press Commandbutton
  If ScanOnly Then FrmMain.dum500.Value = 1
  
 'Get gRawVBCodeSize and gLocalVarsCount
  gCurModStructOffset = Base1 + curFrmStruct.mfrmOffset
  Get hVBFile, gCurModStructOffset, IntBuff:
  gRawVBCodeSize = Uint32(IntBuff)
  gLocalVarsCount = (gRawVBCodeSize \ 2) - 1
  gCurModStructOffset = gCurModStructOffset + 2
  
  If gLocalVarsCount < 0 Then gLocalVarsCount = 0
  
  ReDim gVBCodeTokenAddData(0), gLocalVarsTypeArr(gLocalVarsCount)
  
  gLocalVarsTypeBuff = String$(gLocalVarsCount, 0) & Chr$(EOS_TOKEN) ' &HE5%
  
  ReDim gLocalVarsTypeArrFinal(gLocalVarsCount)
  
  gWriteOutPut = True
  
  'Less than 3 bytes?
  If gRawVBCodeSize < 6 Then Exit Sub

' load raw fromdata
  DispStatus l006E, "Loading " & frmFileN
  
 'Read VBDeclarationData
  ReDim gLocalVarsTypeArrExE(gLocalVarsCount)
  For i = 0 To gLocalVarsCount
    Get hVBFile, , gLocalVarsTypeArrExE(i)
  Next i
  
  Disp_Lvl = FreeFile
  frmFileN = PathAddExt("", Hex$(NewFileExt))

'Stop
On Error Resume Next
  
  If FileLen(frmFileN) = 0 Then Error 1000
  If Err Then
    On Error GoTo 0
   
   'Create "frm.1" file
    ScanFixupsAndTokens
    
    ClearLogUntil l006E
    Exit Sub
  
  End If

On Error GoTo 0
  
  Open frmFileN For Binary Access Read As Disp_Lvl
 
 'CheckVer (normaly 8)
  Get Disp_Lvl, , File1_CodeSizeWords
  If File1_CodeSizeWords < 6 Then
    Close Disp_Lvl
    
    std_p_e = frmFileN
    
   'File not Found or wrong Version
    ShowErrMsg &H4000
    
    ClearLogUntil l006E
    Exit Sub
  
  End If
  
  gWriteOutPut = False
  Get Disp_Lvl, , IntBuff
  
  If IntBuff Then
    If IntBuff <> gLocalVarsCount Then MakeErrBeep
    Get Disp_Lvl, , gLocalVarsTypeBuff
  End If
  
' Read gVBCodeTokens
  If File1_CodeSizeWords >= 8 Then
    For i = 1 To gLocalVarsCount
      Get Disp_Lvl, , gLocalVarsTypeArr(i)
    Next
    
  ' Read Additional data
    Get Disp_Lvl, , IntBuff
    ReDim gVBCodeTokenAddData(IntBuff)
    
    For i = 1 To IntBuff
      Get Disp_Lvl, , IntBuff
      If IntBuff Then
        gVBCodeTokenAddData(i) = Space$(IntBuff)
        Get Disp_Lvl, , gVBCodeTokenAddData(i)
      End If
    Next
  
  Else
    IntBuff = gLocalVarsCount
    
    For i = 1 To IntBuff
      Get Disp_Lvl, , IntBuff
      If IntBuff Then
        l0096 = UBound(gVBCodeTokenAddData) + 1
        ReDim Preserve gVBCodeTokenAddData(l0096)
        gVBCodeTokenAddData(l0096) = Space$(IntBuff)
        Get Disp_Lvl, , gVBCodeTokenAddData(l0096)
        gLocalVarsTypeArr(i) = l0096
      End If
    Next i
    gWriteOutPut = True
    
  End If
  
  If File1_CodeSizeWords = 6 Then
  'Convert tokens
    frmFileN = Space$(gLocalVarsCount)
    Get Disp_Lvl, , frmFileN
    For i = 1 To gLocalVarsCount
      gLocalVarsTypeArrFinal(i) = Asc(Mid$(frmFileN, i, 1))
    Next
    gWriteOutPut = True
  Else
    For i = 1 To gLocalVarsCount
      Get Disp_Lvl, , gLocalVarsTypeArrFinal(i)
    Next
  End If
  
  Close Disp_Lvl
  ClearLogUntil l006E
End Sub

Sub Save1File()
Dim i As Integer
Dim hFile1 As Integer
Dim OutBuff As Integer
 Const mc00A4 = 32 ' &H20%
4050
  
  If NewFileExt <= 0 Then gWriteOutPut = False
  If Not gWriteOutPut Then Exit Sub
  
  hFile1 = FreeFile
  Open PathAddExt("", Hex$(NewFileExt)) For Binary As hFile1
 
' write Version
  OutBuff = 8
  Put hFile1, 1, OutBuff
  
' write Num of CodeSizeWords
  OutBuff = gLocalVarsCount
  Put hFile1, , OutBuff

' write oldTokenStrBuff
  Put hFile1, , gLocalVarsTypeBuff
  
' write gLocalVarsTypeArr
  For i = 1 To gLocalVarsCount
    Put hFile1, , gLocalVarsTypeArr(i)
  Next
  
' write CodeTokenAddDataSize
  OutBuff = UBound(gVBCodeTokenAddData)
  Put hFile1, , OutBuff
  
' write CodeTokenAddData
  For i = 1 To OutBuff
    OutBuff = Len(gVBCodeTokenAddData(i))
    Put hFile1, , OutBuff
    If OutBuff Then Put hFile1, , gVBCodeTokenAddData(i)
  Next i
  
  For i = 1 To gLocalVarsCount
    Put hFile1, , gLocalVarsTypeArrFinal(i)
  Next

' Final 0
  i = 0: Put hFile1, , i
  Close hFile1
  
  gWriteOutPut = False
End Sub

Sub sub016C()
4060
  If bPrjFilesCreated Then sub0178
  If gWriteOutPut Then Save1File
End Sub

Sub sub0189()

4070
  ScanedTokensCount = 0
  Do Until ScanTokens() And mc00A4
  Loop
End Sub

Sub SetStartAndEndOffset(p00A6 As Long, p00A8 As Long)
4080
  If curSegOffs Then
    p00A6 = curSegOffs + Uint32(TokenStruct1.Offs)
    p00A8 = curSegOffs + Uint32(TokenStruct1.size)
  Else
    p00A6 = 0: p00A8 = 0
  End If
  
End Sub

Sub sub023C()
4090
  gFrmStruct(NewFileExt) = curFrmStruct
  bPrjFilesCreated = True
End Sub
