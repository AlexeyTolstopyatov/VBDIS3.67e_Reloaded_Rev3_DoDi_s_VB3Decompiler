Attribute VB_Name = "MODULE11"
' Module1
Option Explicit
'Declare Function GetActiveWindow Lib "User" () As Integer
'Declare Function GetModuleHandle Lib "Kernel" (ByVal p1$) As Integer
'Declare Function GetModuleUsage Lib "Kernel" (ByVal p1%) As Integer
'Declare Function GetPrivateProfileString Lib "Kernel" (ByVal p1$, ByVal p2$, ByVal p3$, ByVal p4$, ByVal p5%, ByVal p6$) As Integer
'Declare Function GetTempDrive Lib "Kernel" (ByVal p1%) As Integer
'Declare Function GetTempFileName Lib "Kernel" (ByVal p1%, ByVal p2$, ByVal p3%, ByVal p4$) As Integer
'Declare Function GetWindowsDirectory Lib "Kernel" (ByVal p1$, ByVal p2%) As Integer
'Declare Function WritePrivateProfileString Lib "Kernel" (ByVal p1$, ByVal p2$, ByVal p3$, ByVal p4$) As Integer

Public Declare Function GetActiveWindow Lib "user32.dll" () As Long
Public Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetTempFileName Lib "kernel32.dll" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long


Public Const STATUS_PENDING As Long = &H103
Public Const STILL_ACTIVE As Long = STATUS_PENDING
Public Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long

Public Const PROCESS_QUERY_INFORMATION As Long = (&H400)
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
'Public Declare Function VDMIsModuleLoaded Lib "VDMDBG.dll" (ByVal szPath As String) As Long

Public Declare Sub ExitProcess Lib "kernel32.dll" (ByVal uExitCode As Long)
Public Const EXIT_CODE_EXIT_NORMAL = 0
Public Const EXIT_CODE_EXIT_SUCCESS = 1
Public Const EXIT_CODE_EXIT_CANCEL = 2

Const vbNormalFocus = 1

Sub AppExit(ExitCode As Long)
   #If DevMode = 1 Then
      Stop
   #Else
      ExitProcess ExitCode
   #End If
End Sub


Sub SaveToTxt()
Dim i
  AppActivate KeyCombos.VB3_Title
  DoEvents

' Bring up ProjectManager
  SendKeys "%(" & KeyCombos.PrjManager & ")", True

DoEvents
  
  For i = 0 To UBound(prjFiles)
ExitVB3:
   'Save all Forms
    If prjFiles(i).FileType <> KeyCombos.TargetType Then
      
      SendKeys "%(" & KeyCombos.FileSaveAsText & "){ENTER}%(" & KeyCombos.Yes & ")", True
      DoEvents
      
    End If
      AppActivate KeyCombos.VB3_Title
      DoEvents
    
    SendKeys "{DOWN}", True
    DoEvents
  Next i
  
  SendKeys "%(" & KeyCombos.ExitVB3 & KeyCombos.Yes & ")", True
  DoEvents
  Screen.MousePointer = 0
 'Quit if Called with CommandLine
  If Command$ <> "" Then
    AppExit EXIT_CODE_EXIT_SUCCESS
  Else
    ShowStatus "Message: Forms in " + sFile.Mak_File + " converted!"

    frm1.WindowState = 0
    frm1.Show
  End If
End Sub


Sub ConvFrmToTxt()
Dim l0046 As Integer
Dim l0048 As Integer
Dim i
  On Error GoTo Error
  sFile.TmpFileName = String(144, Chr(0))
  
  Dim TempPath$, TempPathLen&, retval&
  
 'Make TxtBuffer for TempPath
  TempPathLen = GetTempPath(0, "")
  TempPath = Space(TempPathLen)
 
 'Get TempPath
  retval = GetTempPath(TempPathLen, TempPath)
  If retval = 0 Then Err.Raise vbObjectError + 1, , "Error Getting TempfileDir"
 
' 'Cut Null at the end
'  TempPath = Left(TempPath, retval)



'  l0046 = GetTempFileName(GetTempDrive(ByVal l0048), ByVal "F2T", 0, sFile.TmpFileName)
  l0046 = GetTempFileName(TempPath, ByVal "F2T", 0, sFile.TmpFileName)
  
  FileCopy sFile.Mak_Path & sFile.Mak_File, sFile.TmpFileName
  For i = 0 To UBound(prjFiles)
    prjFiles(i).TmpFileName = String(144, Chr(0))
    
 '   l0046 = GetTempFileName(GetTempDrive(ByVal l0048), ByVal "F2T", 0, prjFiles(i).TmpFileName)
    l0046 = GetTempFileName(TempPath, ByVal "F2T", 0, prjFiles(i).TmpFileName)
    
    FileCopy prjFiles(i).Mak_Path & prjFiles(i).Mak_File, prjFiles(i).TmpFileName
    DoEvents
  Next i
  GoTo Success

Error:
  ShowStatus "Error: Cannot create backup files." & Chr$(10) & Chr$(13) & "Cause: " & Error$ & "."
  End

Success:
End Sub

Sub ClearTmpFiles()
Dim i
  On Error Resume Next
  Kill sFile.TmpFileName
  For i = 0 To UBound(prjFiles)
    Kill prjFiles(i).TmpFileName
  Next i
End Sub

Function FileExists(ByVal Filename As String) As Integer
Dim Size As Long
  On Error Resume Next
  Size = FileLen(Filename$)
  If Size Then FileExists% = True
End Function

Function GetFileType(Path As String, Filename As String) As Integer
Dim char_Buff As String * 1
Dim hFile As Integer
  
  On Error GoTo Error
  
  hFile = FreeFile
  
  Open Path & Filename For Binary As #hFile
  Get #hFile, 1, char_Buff
  
  Select Case Asc(char_Buff)
    
    Case Is = 252, Is = 255
      GetFileType = TYPE_BIN_FRM
    
    Case Else
      GetFileType = TYPE_TXT_FRM
  End Select
  
  Close #hFile
  
  GoTo Success

Error:
  Screen.MousePointer = 0
  ShowStatus "Error: Could not open " & Filename & "." & Chr$(10) & Chr$(13) & "Cause: " & Error$ & "."
  GetFileType = TYPE_ERROR
  Resume Success

Success:
End Function

Function ReadMakFile() As Integer
Dim txtLine As String
Dim Index As Integer
Dim hFile As Integer
Dim i%
Dim Forms() As String
Dim BasForms() As String
  Screen.MousePointer = 11
  On Error Resume Next
  

' Read Forms from Makefile and sort it
  Index = -1
  hFile = FreeFile
  Open sFile.Mak_Path & sFile.Mak_File For Input As #hFile
  Do Until EOF(hFile)
    Line Input #hFile, txtLine
    If txtLine Like "*.FRM" Then
      
     'Extend Array
      Index = Index + 1
      ReDim Preserve Forms(0 To Index) As String
     
     'Save Form
      Forms(Index) = ExtractFileName(txtLine)
      
    End If
  Loop
  Close hFile
  ArraySort Forms()
  
  'Add Forms to prjFiles Array
  ReDim Preserve prjFiles(0 To UBound(Forms)) As T053E
  hFile = FreeFile
  Open sFile.Mak_Path & sFile.Mak_File For Input As #hFile
  Do Until EOF(hFile)
    Line Input #hFile, txtLine
    
    If ExtractFileName(txtLine) Like "*.FRM" Then
      
      For i = 0 To UBound(Forms)
        
        If Forms(i) = CVar(ExtractFileName(txtLine)) Then
          
         'Fill in FileName
          prjFiles(i).Mak_File = ExtractFileName(txtLine)
          
         'Fill in Path
          prjFiles(i).Mak_Path = ExtractPathName(txtLine)
          If prjFiles(i).Mak_Path = "" Then
            prjFiles(i).Mak_Path = sFile.Mak_Path
          End If
         
         'Fill in FileType
          prjFiles(i).FileType = GetFileType(prjFiles(i).Mak_Path, prjFiles(i).Mak_File)
          If prjFiles(i).FileType = TYPE_ERROR Then
            ReadMakFile = False
            Exit Function
          End If
          
          Exit For
        End If
      Next i
    End If
  Loop
  Close hFile
  
  
  
' Read Modules from Makefile and sort it
  Index = -1
  hFile = FreeFile
  Open sFile.Mak_Path & sFile.Mak_File For Input As #hFile
  Do Until EOF(hFile)
    Line Input #hFile, txtLine
    If txtLine Like "*.BAS" Then
      
      Index = Index + 1
      ReDim Preserve BasForms(0 To Index) As String
      
      BasForms(Index) = ExtractFileName(txtLine)
    
    End If
  Loop
  Close hFile
  ArraySort BasForms()
  
 
 'Add Modules to prjFiles Array
  ReDim Preserve prjFiles(0 To UBound(Forms) + UBound(BasForms) + 1) As T053E
  hFile = FreeFile
  Open sFile.Mak_Path & sFile.Mak_File For Input As #hFile
  Do Until EOF(hFile)
    Line Input #hFile, txtLine
    
    If ExtractFileName(txtLine) Like "*.BAS" Then
      
     'Get Details for Modul
      For i = 0 To UBound(BasForms)
         If BasForms(i) = CVar(ExtractFileName(txtLine)) Then
            With prjFiles(i + 1 + UBound(Forms))
               
              'Filename
               .Mak_File = ExtractFileName(txtLine)
              
              'Path
               .Mak_Path = ExtractPathName(txtLine)
               If .Mak_Path = "" Then
                 .Mak_Path = sFile.Mak_Path
               End If
               
               'FileType
               .FileType = GetFileType(.Mak_Path, .Mak_File)
               If .FileType = TYPE_ERROR Then
                 ReadMakFile = False
                 Exit Function
               End If
               
            End With
         Exit For
        End If
      Next i
      
    End If
  Loop
  Close hFile
  
' If CommandLine Set convert to Txt
  If Command$ <> "" Then
    KeyCombos.TargetType = TYPE_TXT_FRM
  Else
    AskForTargetType
  End If
  
  On Error Resume Next
  Dim void%
  void = LBound(prjFiles)
  If Err Then
    Screen.MousePointer = 0
    ShowStatus "Error: No forms/modules in project." & Chr$(10) & Chr$(13) & "Cause: Non-VB file selected."
    ReadMakFile = False
  Else
    ReadMakFile = True
  End If
End Function

Sub InitVars()
Dim StrBuff As String
  Screen.MousePointer = 11
  StrBuff = String(255, Chr(0))
  KeyCombos.VB_PATH_Section = "Visual Basic"
  KeyCombos.VB_PATH = "vbpath"
  KeyCombos.VB3_Title = "Microsoft Visual Basic"
  KeyCombos.VB_EXE = "VB.EXE"
  KeyCombos.TimeToWait = 100
  
  If FileExists(App.Path & "\FRMS2TXT.INI") Then
    KeyCombos.PrjManager = LCase(Left(StrBuff, GetPrivateProfileString("Forms2Text", ByVal "WindowProject", "", StrBuff, Len(StrBuff), App.Path + "\" + "FRMS2TXT.INI")))
    KeyCombos.FileSaveAsText = LCase(Left(StrBuff, GetPrivateProfileString("Forms2Text", ByVal "FileSaveAsText", "", StrBuff, Len(StrBuff), App.Path + "\" + "FRMS2TXT.INI")))
    KeyCombos.ExitVB3 = LCase(Left(StrBuff, GetPrivateProfileString("Forms2Text", ByVal "FileExit", "", StrBuff, Len(StrBuff), App.Path + "\" + "FRMS2TXT.INI")))
    KeyCombos.Yes = LCase(Left(StrBuff, GetPrivateProfileString("Forms2Text", ByVal "Yes", "", StrBuff, Len(StrBuff), App.Path + "\" + "FRMS2TXT.INI")))
  End If
  
  Screen.MousePointer = 0
  
  If KeyCombos.PrjManager = "" Or KeyCombos.FileSaveAsText = "" Or KeyCombos.ExitVB3 = "" Or KeyCombos.Yes = "" Then
    frm3.Show 1
  End If
End Sub

Sub AskForTargetType()
Dim l0080 As Integer
Dim l0082
Dim l0084 As Integer
  l0080 = prjFiles(0).FileType
  For l0082 = 0 To UBound(prjFiles)
    Select Case prjFiles(l0082).FileType
      Case Is = l0080
        l0084 = False
      Case Else
        l0084 = True
        Exit For
    End Select
  Next l0082
  Select Case l0084
    Case True
      frm6.Show 1
    Case Else
      Select Case prjFiles(0).FileType
        Case TYPE_BIN_FRM
          KeyCombos.TargetType = TYPE_TXT_FRM
        Case Else
          KeyCombos.TargetType = TYPE_BIN_FRM
      End Select
  End Select
End Sub

Sub ArraySort(InArray() As String)
Dim Start As Integer
Dim Elements As Integer
Dim FirstElement As String
  Elements = UBound(InArray)
  Start = 1
  
' Go Forward for Start...Elements
  While (Start <= Elements)
    
    Call QuickSortA(InArray(), Start)
    
    Start = Start + 1
    globalLines = globalLines + 1
  Wend
  
  
' Go Backwards for Elements ... Start
  Start = Elements
  While (Start > 0)
    
    FirstElement = InArray(0)
    InArray(0) = InArray(Start)
    InArray(Start) = FirstElement
    
    Call QuickSortB(InArray(), Start - 1)
    
    Start = Start - 1
    globalLines = globalLines + 1
  Wend
End Sub

Sub QuickSortB(InArray() As String, Start As Integer)
Dim PosA As Integer
Dim PosB As Integer
Dim Tmp As String
  PosA = 0
  PosB = 2 * PosA
  
  Do While (PosB <= Start)
    If (PosB < Start And InArray(PosB) < InArray(PosB + 1)) Then
      PosB = PosB + 1
    End If
    If InArray(PosA) >= InArray(PosB) Then
      Exit Do
    End If
    
  ' Swap PosA & PosB
    Tmp = InArray(PosA)
    InArray(PosA) = InArray(PosB)
    InArray(PosB) = Tmp
    
    PosA = PosB
    PosB = 2 * PosA
    
    globalLines = globalLines + 1
  Loop
End Sub

Sub QuickSortA(InArray() As String, Start As Integer)
Dim iStart As Integer
Dim half As Integer
Dim ArrayData As String
  iStart = Start
  
  Do While (iStart > 0)
    half = Int(iStart / 2)
    If InArray(half) >= InArray(iStart) Then
      Exit Do
    End If
  
  ' Swap
    ArrayData = InArray(iStart)
    InArray(iStart) = InArray(half)
    InArray(half) = ArrayData
    
    iStart = half
    globalLines = globalLines + 1
  Loop
End Sub

Function IsExeRunning(p00B4 As String) As Integer

' Stop
 ' IsExeRunning = GetModuleUsage(GetModuleHandle(p00B4))
 'ToDo: http://support.microsoft.com/kb/178893/
 'VDMTerminateTaskWOW...
End Function

Sub Main()

  
  On Error GoTo L2CDC
  InitVars
  Load frm1
 
' are there Parameters set through the Commandline?
  If FileExists(Command$) Then
  ' CommandLine Mode
    sFile.Mak_Path = ExtractPathName(Command$)
    sFile.Mak_File = ExtractFileName(Command$)
    If ReadMakFile() Then DoConvention
  Else
    frm1.Show
  End If
  ClearTmpFiles
  GoTo L2D5C

L2CDC:
  ShowStatus "Error: Unhandled Exception." & Chr$(10) & Chr$(13) & "Cause: " & Error$ & "."
  Resume Next

L2D5C:
End Sub

Sub ShowStatus(ByVal p00B6 As String)
  Load frm5
  frm5.control2 = p00B6
  frm5.Show 1
End Sub

Sub sub01FF()
Dim l00BC
  On Error Resume Next
  FileCopy sFile.TmpFileName, sFile.Mak_Path & sFile.Mak_File
  For l00BC = 0 To UBound(prjFiles)
    FileCopy prjFiles(l00BC).TmpFileName, prjFiles(l00BC).Mak_Path & prjFiles(l00BC).Mak_File
  Next l00BC
End Sub

Sub DoConvention()
Dim TaskID As Integer
Dim StopWaiting As Integer
Dim LastTimer As Single
Dim hActiveWnd&
  frm1.WindowState = 1
  DoEvents
  
' is VB already running?
  If IsExeRunning(KeyCombos.VB_EXE) <> 0 Then
   '-> yes
    Screen.MousePointer = 0
    ShowStatus "Error: Forms2Text could not launch Visual Basic." & vbCrLf & "Cause: VB is already running."
    
   'Quit on Commandline else Exit this function
    If Command$ = "" Then
      frm1.WindowState = 0
      frm1.Show
      Exit Sub
    Else
      AppExit (EXIT_CODE_EXIT_CANCEL)
    End If
  End If
  
  On Error GoTo Error

 ' Convert MakeFileName to a short 8.3 FileName (Because VB3 is a 16bit App)
   Dim ShortFilePath$, retval&
   ShortFilePath = Space(128)
   retval = GetShortPathName(sFile.Mak_Path & sFile.Mak_File, ShortFilePath, 128)
   If retval = 0 Then Err.Raise vbObjectError + 2, , "Convert to 8.3-FileName failed!"
   ShortFilePath = Left(ShortFilePath, retval)
   
' StartVB3
  TaskID = Shell(GetVB3Path() & "\" & KeyCombos.VB_EXE & " " & _
                 ShortFilePath, vbNormalFocus)
  On Error GoTo 0
  
  StopWaiting = False
  LastTimer = Timer
  
  Do
    TaskID = DoEvents()
    hActiveWnd = GetActiveWindow()
    If Timer - LastTimer! > KeyCombos.TimeToWait Then StopWaiting = True
  Loop While hActiveWnd = frm1.hWnd And Not StopWaiting
  
  AppActivate KeyCombos.VB3_Title
  frm4.Show
  
  GoTo Success

Error:
  Screen.MousePointer = 0
  ShowStatus "Error: Forms2Text could not launch Visual Basic." & Chr$(10) & Chr$(13) & "Cause: " & Error$ & "."
  If (Command$) = "" Then
    frm1.WindowState = 0
    frm1.Show
    Exit Sub
  Else
    AppExit (EXIT_CODE_EXIT_CANCEL)
  End If

Success:
End Sub

Function ExtractPathName(p00CC As String) As String


Dim l00CE As Integer
Dim l00D0 As Integer
  If InStr(p00CC$, "\") = 0 Then
    ExtractPathName = ""
    Exit Function
  End If
  ExtractPathName$ = p00CC$
  l00CE% = InStr(p00CC$, "\")
  Do While l00CE%
    l00D0% = l00CE%
    l00CE% = InStr(l00D0% + 1, p00CC$, "\")
  Loop
  If l00D0% > 0 Then ExtractPathName$ = Mid$(p00CC$, 1, l00D0%)
End Function

Function ExtractFileName(FullFileName As String) As String

 ' Split FullFileName a "\" and store it in a Array
   Dim FilenameSplited
   FilenameSplited = Split(FullFileName, "\")
 
 ' The Last element is the Filename, so return it
   ExtractFileName = FilenameSplited(UBound(FilenameSplited))

End Function

Function GetVB3Path() As String
Dim Buff As String
   Buff = String(255, Chr(0))
   GetVB3Path = Left(Buff, _
      GetPrivateProfileString( _
         KeyCombos.VB_PATH_Section, ByVal KeyCombos.VB_PATH, "C:\VB", _
         Buff, Len(Buff), _
         Left$(Buff, _
            GetWindowsDirectory(Buff, Len(Buff)) _
         ) + "\" + "VB.INI"))
End Function
