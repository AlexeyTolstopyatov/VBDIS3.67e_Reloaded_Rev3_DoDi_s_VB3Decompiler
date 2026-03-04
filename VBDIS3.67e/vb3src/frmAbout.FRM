VERSION 2.00
Begin Form frmAbout 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "DoDi's VB Tools"
   ClientHeight    =   2190
   ClientLeft      =   4815
   ClientTop       =   1860
   ClientWidth     =   4950
   Height          =   2700
   Icon            =   0
   Left            =   4755
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4950
   Top             =   1410
   Width           =   5070
   Begin PictureBox dum203 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   408
      Left            =   240
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   5
      Top             =   960
      Width           =   408
   End
   Begin CommandButton dum303 
      Caption         =   "wow !"
      Default         =   -1  'True
      Height          =   552
      Left            =   3960
      TabIndex        =   4
      Top             =   180
      Width           =   852
   End
   Begin Label dum403 
      Caption         =   "http://members.aol.com/vbdis/"
      Height          =   195
      Index           =   1
      Left            =   1800
      TabIndex        =   12
      Top             =   1260
      Width           =   3075
   End
   Begin Label dum503 
      Caption         =   "URL"
      Height          =   195
      Index           =   1
      Left            =   960
      TabIndex        =   11
      Top             =   1260
      Width           =   735
   End
   Begin Label dum503 
      Caption         =   "e-mail"
      Height          =   195
      Index           =   0
      Left            =   960
      TabIndex        =   6
      Top             =   1020
      Width           =   735
   End
   Begin Label dum603 
      Caption         =   "<your e-mail address>"
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   1860
      Width           =   3075
   End
   Begin Label dum703 
      Caption         =   "<your name>"
      Height          =   240
      Left            =   1800
      TabIndex        =   9
      Top             =   1635
      Width           =   3030
   End
   Begin Label dum803 
      Caption         =   "Translated by"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1635
      Width           =   1335
   End
   Begin Label dum403 
      Caption         =   "DoDi@compuserve.com"
      Height          =   195
      Index           =   0
      Left            =   1800
      TabIndex        =   7
      Top             =   1020
      Width           =   3075
   End
   Begin Label dum903 
      Caption         =   "2.0"
      Height          =   192
      Left            =   1080
      TabIndex        =   3
      Top             =   180
      Width           =   2292
   End
   Begin Label duma03 
      Caption         =   "Dr. H.-P. Diettrich"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin Label dumb03 
      Caption         =   "Copyright 1997"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin Label dumc03 
      AutoSize        =   -1  'True
      Caption         =   "Version"
      Height          =   192
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   852
   End
End
'Attribute VB_Name = "frmAbout"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = False
' F0003.FRM
Option Explicit

Private Sub dum303_Click ()
44000
  Unload Me
End Sub

Private Sub Form_Load ()
   Dim lCrackedAndFucked As String

   44010
   dum203.Picture = CurForm_gv0BC0.Icon
   Caption = "DoDi's " & App.Title
   lCrackedAndFucked = " cracked Demo + Src :)"
   dum903 = gVersionStr & gRevisionStr & sCrackedAndFucked
   If Left$(dum703, 1) <> "<" Then
   Else
      Height = Height - (ScaleHeight - dum803.Top)
   End If
End Sub

