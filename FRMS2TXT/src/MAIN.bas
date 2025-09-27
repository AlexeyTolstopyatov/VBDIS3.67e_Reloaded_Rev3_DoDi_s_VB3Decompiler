Attribute VB_Name = "Globals"
' main.txt - global definitions
Type T0485
  PrjManager As String
  FileSaveAsText As String
  ExitVB3 As String
  Yes As String
  VB_PATH_Section As String
  VB_PATH As String
  VB3_Title As String
  VB_EXE As String
  TimeToWait As Integer
  TargetType As Integer
End Type

Type T0517
  Mak_Path As String
  Mak_File As String
  TmpFileName As String
End Type

Type T053E
  Mak_Path As String
  Mak_File As String
  TmpFileName As String
  FileType As Integer
End Type

Global sFile As T0517
Global prjFiles() As T053E
Global KeyCombos As T0485
Global TYPE_ERROR As Integer
Global Const TYPE_TXT_FRM = 1 ' &H1%
Global Const TYPE_BIN_FRM = 2 ' &H2%
Global globalLines
