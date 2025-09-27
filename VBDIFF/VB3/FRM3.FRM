VERSION 2.00
Begin Form frm3 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "DoDi's VB Tools"
   ClientHeight    =   1650
   ClientLeft      =   1125
   ClientTop       =   2280
   ClientWidth     =   4965
   Height          =   2160
   Icon            =   0
   Left            =   1065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4965
   Top             =   1830
   Width           =   5085
   Begin PictureBox control1 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   408
      Left            =   240
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   5
      Top             =   780
      Width           =   408
   End
   Begin CommandButton control2 
      Caption         =   "wow !"
      Default         =   -1  'True
      Height          =   552
      Left            =   3993
      TabIndex        =   4
      Top             =   605
      Width           =   852
   End
   Begin Label control3 
      Height          =   253
      Left            =   1936
      TabIndex        =   9
      Top             =   1331
      Width           =   1947
   End
   Begin Label control4 
      Height          =   253
      Left            =   242
      TabIndex        =   8
      Top             =   1331
      Width           =   1342
   End
   Begin Label control5 
      Caption         =   "D-70188 Stuttgart"
      Height          =   252
      Left            =   1920
      TabIndex        =   7
      Top             =   960
      Width           =   1932
   End
   Begin Label control6 
      Caption         =   "Landhausstr. 265"
      Height          =   252
      Left            =   1920
      TabIndex        =   6
      Top             =   720
      Width           =   1932
   End
   Begin Label control7 
      Caption         =   "2.0"
      Height          =   192
      Left            =   1080
      TabIndex        =   3
      Top             =   180
      Width           =   2292
   End
   Begin Label control8 
      Caption         =   "Dr. H.-P. Diettrich"
      Height          =   252
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   1932
   End
   Begin Label control9 
      Caption         =   "Copyright 1996"
      Height          =   192
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1572
   End
   Begin Label control10 
      AutoSize        =   -1  'True
      Caption         =   "Version"
      Height          =   192
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   852
   End
End
' frm3.FRM
Option Explicit

Sub control2_Click ()
  Unload Me
End Sub

Sub Form_Load ()
Dim l001A As String
  control1.Picture = Screen.ActiveForm.Icon
  Caption = "DoDi's " & App.Title
  control7 = gc001E & gc0006 & l001A
    Height = Height - (ScaleHeight - control4.Top)
End Sub
