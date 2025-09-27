VERSION 2.00
Begin Form Form2 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "DoDi's VB Tools"
   ClientHeight    =   1635
   ClientLeft      =   2880
   ClientTop       =   2235
   ClientWidth     =   4950
   Height          =   2145
   Icon            =   0
   Left            =   2820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4950
   Top             =   1785
   Width           =   5070
   Begin PictureBox Pic1 
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
   Begin CommandButton Command1 
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
      Caption         =   "D-70180 Stuttgart"
      Height          =   252
      Left            =   1920
      TabIndex        =   7
      Top             =   960
      Width           =   1932
   End
   Begin Label control6 
      Caption         =   "Weissenburgstr. 2 c"
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
      Caption         =   "Copyright 1995,96"
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
' F0001.FRM
Option Explicit

Sub Command1_Click ()
  Unload Me
End Sub

Sub Form_Load ()
Dim l0020 As String
pic1.Picture = Screen.ActiveForm.Icon
  Caption = "DoDi's " & App.Title
  control7 = VERSION & VERSION_TYPE & l0020
    control4 = Descramble("Ztoohpzbbzf)")
    control3 = Descramble("_nt;Hv~vy")
End Sub

