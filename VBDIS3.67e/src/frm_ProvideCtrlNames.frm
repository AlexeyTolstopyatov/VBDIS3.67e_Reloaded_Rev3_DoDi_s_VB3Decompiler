VERSION 5.00
Begin VB.Form frm_CtrlNames 
   Caption         =   "Enter some Controllnames"
   ClientHeight    =   6135
   ClientLeft      =   2610
   ClientTop       =   315
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   5505
   Begin VB.CheckBox chk_HideMe 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      Caption         =   "Don't show me again for this form!  (to bring me back delete frmX.FRM.txt)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5535
   End
   Begin VB.TextBox Txt_ControlName 
      BorderStyle     =   0  'Kein
      Height          =   6135
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "frm_CtrlNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public FormName$
Public ctrls_Filename$
Public ControlsCount%
Public HideMe%
Dim ctrls_hFile%
Dim TextAsArray

Private Sub chk_HideMe_Click()
   If chk_HideMe.Value = vbChecked Then Unload Me
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   Txt_ControlName.Width = Me.Width - 150
   Txt_ControlName.Height = Me.Height - 550 - Txt_ControlName.Top
End Sub




Private Sub Form_Load()
   Me.Caption = FormName & ": Wups someone deleted Control names for "
   
   LoadData
   
End Sub
 
Private Sub MakeEntrysValid()
  
  TextAsArray = Split(Txt_ControlName, vbCrLf)

' Limit extent to num of controls
  ReDim Preserve TextAsArray(ControlsCount)
  
 ' == Find Duplicates ==
 ' Go though all lines
  Dim i
  For i = LBound(TextAsArray) To UBound(TextAsArray)
  
   'Replace empty or to short entrys
    If Len(TextAsArray(i)) <= 2 Then
      'First control is always the FormName
       If i = LBound(TextAsArray) Then
         TextAsArray(i) = Split(FormName, ".")(0)
       Else
         TextAsArray(i) = "control" & Format$(i)
       End If
    End If
  
  Next
  
 ' == Find Duplicates ==
   
   Dim DoublesFound As Boolean
   
   Do
      DoublesFound = False
      
      Dim j ',i
      For i = LBound(TextAsArray) To UBound(TextAsArray)
        For j = i + 1 To UBound(TextAsArray)
          If TextAsArray(i) = TextAsArray(j) Then
             TextAsArray(j) = TextAsArray(j) & "_2"
             
             DoublesFound = True
             
          End If
        Next
      Next
      
   Loop While DoublesFound
  
  

' Fill in new text in Textbox
  Txt_ControlName = Join(TextAsArray, vbCrLf)

End Sub
Private Sub LoadData()
 On Error Resume Next
' On Error GoTo 0
' ctrls_Filename = "M:\t\1\ControlsCount.txt"
 
 
 Txt_ControlName = ""
 ctrls_hFile = FreeFile
 
 Open ctrls_Filename For Input Shared As ctrls_hFile
 
' Wups they got deleted so Generate ControlNames
  Dim j
  For j = 0 To ControlsCount
    
    Dim tmpline$
    tmpline = ""
    
    Input #ctrls_hFile, tmpline
    Txt_ControlName = Txt_ControlName & tmpline & vbCrLf
    

  Next
  
 'Load HideMe
  HideMe = 0
  Input #ctrls_hFile, HideMe
  Close ctrls_hFile
 
  MakeEntrysValid
  
 'If HideMe set unload me
  If HideMe = vbChecked Then Unload Me
  
End Sub

Private Sub Form_Unload(gCancel As Integer)
 '  On Error Resume Next
   
   MakeEntrysValid
   If HideMe = False Then
      ctrls_hFile = FreeFile
      Open ctrls_Filename For Output Shared As ctrls_hFile
   End If
 
    ' Save TextBox
      Dim i
      For i = LBound(TextAsArray) To UBound(TextAsArray)
         gv0FF6(gControlCount1).Name_4 = TextAsArray(i)
         
         gControlCount1 = gControlCount1 + 1
         If HideMe = False Then Print #ctrls_hFile, TextAsArray(i)
      Next
      
      If HideMe = False Then
         Print #ctrls_hFile, chk_HideMe.Value
         Close #ctrls_hFile
      End If
   
End Sub

Private Sub Txt_ControlName_Click()
   MakeEntrysValid
End Sub

Private Sub Txt_ControlName_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyEscape Then Unload Me
End Sub
