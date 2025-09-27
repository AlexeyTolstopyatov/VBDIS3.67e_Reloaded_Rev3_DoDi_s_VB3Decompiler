VERSION 2.00
Begin Form frm_CtrlNames 
   Caption         =   "Enter some Controllnames"
   ClientHeight    =   6135
   ClientLeft      =   2610
   ClientTop       =   315
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   5505
   Begin TextBox Txt_ControlName 
      BorderStyle     =   0  'Kein
      Height          =   6135
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   5535
   End
End
'Attribute VB_Name = "frm_CtrlNames"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = False

Option Explicit
Public ctrls_Filename$
Public ControlsCount%

Dim ctrls_hFile%

Private Sub Form_Load()

   Reload
End Sub
 
Private Sub Reload()
 On Error Resume Next
' On Error GoTo 0
' ctrls_Filename = "M:\t\1\ControlsCount.txt"
 
 
 Txt_ControlName = ""
 ctrls_hFile = FreeFile
 
 Open ctrls_Filename For Input Shared As ctrls_hFile
 
' Wups they got deleted so Generate ControlNames
  Dim j
  For j = 0 To ControlsCount
    Dim tmpline
    tmpline = ""
    Input #ctrls_hFile, tmpline
    tmpline = Trim(tmpline)
    
    If Len(tmpline) <= 2 Then tmpline = "control" & Format$(j)
  
    Txt_ControlName = Txt_ControlName & tmpline & vbCrLf
    

  Next j
  
  Close ctrls_hFile
  
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   Txt_ControlName.Width = Me.Width - 150
   Txt_ControlName.Height = Me.Height - 550
End Sub

Private Sub Form_Unload(Cancel As Integer)
 '  On Error Resume Next
   
   Dim a
   a = Split(Txt_ControlName, vbCrLf)
   ReDim Preserve a(ControlsCount)

   ctrls_hFile = FreeFile
   Open ctrls_Filename For Output Shared As ctrls_hFile%
 
      Dim i
      For i = LBound(a) To UBound(a)
         gv0FF6(gControlCount1).Name_4 = a(i)
         
         gControlCount1 = gControlCount1 + 1
         Print #ctrls_hFile, a(i)
         On Error Resume Next
      Next
      
   Close #ctrls_hFile

End Sub

