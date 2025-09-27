'Attribute VB_Name = "MODULE20"
' Module20
Option Explicit
Const mc004A = 2 ' &H2%
'Const mc006A = 8192 ' &H2000%

Sub CreateMAKFile()
Dim MAKFileName As String
Dim Displvl1%

28000
  
  MAKFileName = PathAddExt(gVBTitle, "MAK")
  DispStatus Displvl1, "Creating " & MAKFileName
  
  OpenOutputFile MAKFileName
  Create_MAK_File
  CloseOutputFile
  
  ClearLogUntil Displvl1
End Sub

Sub WriteFrmFile(ModulIdx As Integer)
' called (p1 As type, p2%)
Dim Bas_FileName As String
Dim i%
'Dim l0018 As Integer
Dim Displvl1%

28010
  Screen.MousePointer = vbHourglass
  
  If ModulIdx <> NewFileExt Then ScanModuleToken1 ModulIdx
  
 'Open "xyz.Bas"
  Bas_FileName = PathAddExt(ModulFileNames(ModulIdx), "bas")
  OpenOutputFile Bas_FileName
 ' l0018 = hFile_Out
  
 ' Output: "' CALC.FRM"
  Print #hFile_Out, "' "; ModulFileNames(ModulIdx)
  DispStatus Displvl1, "Creating declarations"
  
  ' Output: "Dim m003A As Integer" ...
  CreateDeclarations ModulIdx
  
  For i = 1 To gSubroutines
    If g_TKStruct(i).mSubroutinesNameLen = gSubroutinesNameLen Then
      If g_TKStruct(i).Size_M13F4 <> -1 Then
        DispLog2 Displvl1, "Creating " & getFrmNamebyOffbyIdx(i)
      
      ' Write NewLine [ between '...End Sub' and 'Sub']
        Print #hFile_Out,
        DecompileSubroutine i
      End If
    End If
  Next
  
  CloseOutputFile
  
  ClearLogUntil Displvl1
  
End Sub

Sub WritePrjFiles()
Dim i%
Dim Displvl1%

28020
  DispStatus Displvl1, "Creating project files"
    DispLog2 Displvl1, "Creating " & "main.txt"
    
    OpenOutputFile PathAddExt("main", "txt")
    Write_main_txt
    CloseOutputFile
    
  For i = 1 To FrmFilesCount
    DispLog2 Displvl1, "Creating " & ModulFileNames(i)
    WriteFrmFile i ', 0
  Next i
  ClearLogUntil Displvl1
End Sub

Sub CreatePrjFiles()
28030
  If FrmFilesCount = 0 Then
    MsgBox "Cheat!"
    Exit Sub
  End If
  DoDecompile True
  
  CreateMAKFile
  
  WritePrjFiles
  
End Sub


Sub Create_MAK_File()
Dim itmp As Integer
Dim i%
Dim tmpStr As String
 Const mc004A = 2 ' &H2%
 

28040
  If gv0F88 Then
    itmp = RotateRight3(gv0F88)
    itmp = NamesLookUpTbl(itmp + mc004A)
    If itmp = 0 Then MakeErrBeep
    
  'Write FormNames
    For i = 1 To FrmFilesCount
      If gFrmStruct(i).mSubroutinesNameLen = itmp Then
        tmpStr = ModulFileNames(i)
        If InStr(tmpStr, ".") = 0 Then tmpStr = tmpStr & ".BAS"
        DisOutput tmpStr: WriteOutputTxt
        Exit For
      End If
    Next
    
  ElseIf Start_Form_Index Then
    itmp = RotateRight3(Start_Form_Res)
    DisOutput VBControlNames(Start_Form_Index): WriteOutputTxt
  Else
  End If
  
' Write VBX
  For i = 1 To VBX_Names_Count
    DisOutput VBControlNames(i):    WriteOutputTxt
  Next
  
' Write main.txt
  DisOutput "main.txt": WriteOutputTxt
  
' Write Modules
  For i = 1 To FrmFilesCount
    
  'filter out Start Form
    If gFrmStruct(i).mSubroutinesNameLen <> itmp Then
      
      tmpStr = ModulFileNames(i)
      If InStr(tmpStr, ".") = 0 Then tmpStr = tmpStr & ".BAS"
      DisOutput tmpStr:      WriteOutputTxt
    
    End If
    
  Next

' If VB3- Is Installed...
  If Len(VB3Path_2) Then
    itmp = FreeFile

' Overtake 'ProjWinSize=152,402,248,215' from autoload.mak
' into mak file
On Error Resume Next
    Open VB3Path_2 & "autoload.mak" For Input As itmp
    If Err = 0 Then
      Do Until EOF(itmp)
      
      ' Scan Lines for PROJWIN
        Line Input #itmp, tmpStr
        If UCase$(Left$(tmpStr, 7)) = "PROJWIN" Then
        
          DisOutput tmpStr: WriteOutputTxt
          
          Exit Do
        End If
      Loop
      Close itmp
      
    End If
On Error GoTo 0
  End If
'FormIcon
  If Proj_IconForm Then
  
  ' An Already know Problem occured
    ShowErrMsg &H2000
    DisOutput "IconForm=""":    DisOutput VBControlNames(Proj_IconForm):    DisOutput """"
    WriteOutputTxt
  End If
  
  DisOutput "Title=""":  DisOutput Proj_Title:  DisOutput """": WriteOutputTxt
  
  tmpStr = gIn_FileFulPath
  i = ExtractPath(tmpStr, "\")
  If i Then tmpStr = Mid$(tmpStr, i + 1)
  
  DisOutput "ExeName=""":  DisOutput tmpStr:  DisOutput """": WriteOutputTxt
  
  If Len(Proj_HelpFile) Then
    DisOutput "HelpFile=""":    DisOutput Proj_HelpFile:    DisOutput """": WriteOutputTxt
  End If
End Sub
