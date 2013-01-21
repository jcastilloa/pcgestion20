VERSION 5.00
Object = "{21AB44A3-E9FF-4C98-8931-5DFEB230BB20}#1.2#0"; "GrdBttn.ocx"
Begin VB.Form frmTest 
   BackColor       =   &H00CC0099&
   Caption         =   "Test ActiveX Gradient Button"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7275
   Icon            =   "Test.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin PCGestion.ucGrdBttn ucGrdBttn4 
      Height          =   1335
      Left            =   4620
      TabIndex        =   4
      Top             =   3540
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   2355
      Caption         =   "GrdBttn4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "Test.frx":000C
      GradientColor1_Percent=   200
      GradientColor2_Percent=   1
   End
   Begin PCGestion.ucGrdBttn ucGrdBttn2 
      Height          =   1545
      Left            =   630
      TabIndex        =   3
      Top             =   1560
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   2725
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "Test.frx":0028
      GColor_Boost0   =   -1  'True
      GradientColor1  =   13369497
      GradientColor1_Percent=   155
      GradientColor2  =   13369497
      GradientColor2_Percent=   100
   End
   Begin PCGestion.ucGrdBttn ucGrdBttn1 
      Height          =   2505
      Left            =   4980
      TabIndex        =   2
      Top             =   180
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   4419
      Caption         =   $"Test.frx":0044
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "Test.frx":0068
      GradientColor1_Percent=   300
   End
   Begin PCGestion.ucGrdBttn ucGrdBttn3 
      Height          =   1725
      Left            =   2490
      TabIndex        =   1
      Top             =   210
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   3043
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "Test.frx":0084
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   1125
      Left            =   120
      TabIndex        =   0
      Top             =   4020
      Width           =   3465
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    'Unload Me
    ucGrdBttn3.Font.Bold = Not ucGrdBttn3.Font.Bold
    ucGrdBttn3.Refresh
End Sub

'Private Sub ucGrdBttn2_Click(Index As Integer)
'    MsgBox "&"
'End Sub

'Private Sub ucGrdBttn1_Click()
'    ucGrdBttn1.GradientColor1_Percent
'End Sub
