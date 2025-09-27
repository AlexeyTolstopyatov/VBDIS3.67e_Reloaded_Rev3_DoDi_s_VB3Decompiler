Attribute VB_Name = "MODULE5"
' Module5
Option Explicit
'Declare Function GetPrivateProfileString Lib "Kernel" (ByVal p1$, ByVal p2$, ByVal p3$, ByVal p4$, ByVal p5%, ByVal p6$) As Integer
  Declare Function GetPrivateProfileString Lib "kernel32" _
      Alias "GetPrivateProfileStringA" (ByVal lpApplicationName _
      As String, ByVal lpKeyName As Any, ByVal lpDefault As _
      String, ByVal lpReturnedString As String, ByVal nSize As _
      Long, ByVal lpFileName As String) As Long

'Declare Function GetSystemDirectory Lib "Kernel" (ByVal p1$, ByVal p2%) As Integer
   Declare Function GetSystemDirectory Lib "kernel32" _
      Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, _
      ByVal nSize As Long) As Long


 '    #If Win32 Then
         Declare Function GetWindowsDirectory Lib "kernel32" _
            Alias "GetWindowsDirectoryA" _
            (ByVal lpBuffer As String, ByVal nSize As Long) As Long
   '   #Else
'         Declare Function GetWindowsDirectory Lib "Kernel" _
            (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
    '  #End If

Declare Sub WinHelp Lib "User" ()
Const mc0048 = 48 ' &H30%

Function GetWinIniItem(ByVal Item As String) As String
Dim RetVal As Integer
Dim buff As String

7000
  buff = Space$(128)
  RetVal = GetPrivateProfileString("Visual Basic" & Chr$(0), Item & Chr$(0), "", buff, 128, "vb.ini")
  GetWinIniItem = Left$(buff, RetVal)
End Function

Sub InitPaths(ByVal LogListCount As Integer)
Dim RetVal As Integer
Dim buff As String
Dim l0036 As Integer
 Const mc0048 = 48 ' &H30%


7010
On Error Resume Next
  AppPathwithSlash = AddPathSlash((App.Path))
  std_p_e = "std_p_e.300"
  
  If LogListCount Then DispLog2 LogListCount, "Loading " & std_p_e
  
  While Len(Dir$(AppPathwithSlash & std_p_e, vbNormal)) = 0
    l0036 = Screen.MousePointer
    Screen.MousePointer = vbDefault
      
      MsgBox AppPathwithSlash & std_p_e, mc0048, "File not found"
      
      FrmOpenFile.Tag = std_p_e
      FrmOpenFile.Show 1
      
    Screen.MousePointer = l0036
    buff = FrmOpenFile.Tag
    
    If Len(buff) = 0 Then End
    
    If Right$(buff, 3) = "300" Then
      RetVal = ExtractPath(buff, "\")
      If RetVal Then AppPathwithSlash = Left$(buff, RetVal)
    End If
  Wend
  
  buff = Space$(256)
  RetVal = GetWindowsDirectory(buff, Len(buff))
  gWindowsDir = AddPathSlash(Left$(buff, RetVal))
  
  RetVal = GetSystemDirectory(buff, Len(buff))
  gSystemDir = AddPathSlash(Left$(buff, RetVal))
  
  gVB16Location = GetWinIniItem("vb16location")
  If Len(gVB16Location) Then
    VB3Path = gVB16Location
    RetVal = ExtractPath(VB3Path, "\")
    If RetVal Then VB3Path = Left$(VB3Path, RetVal)
  Else
  
    VB3Path_2 = GetWinIniItem("vbpath")
    If Len(VB3Path_2) = 0 Then
      VB3Path_2 = AppPathwithSlash
      If Len(Dir$(VB3Path_2 & "vb.exe")) = 0 Then
        VB3Path_2 = ""
      End If
    End If
    
    If Len(VB3Path_2) Then
      
      If Right$(VB3Path_2, 1) <> "\" Then VB3Path_2 = VB3Path_2 & "\"
      VB3Path = VB3Path_2
    Else
      MsgBox "Visual Basic not found", &H40
    End If
  End If
  Err = 0
End Sub
