VERSION 5.00
Begin VB.Form AboutBox 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2820
   ClientLeft      =   3450
   ClientTop       =   2070
   ClientWidth     =   3645
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2820
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Bword is programmed by Bidan Sinha (bidan_sinha@rediffmail.com)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   3495
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "for   Windows   v1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "TM"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "BWord"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "AboutBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload AboutBox
End Sub
Private Sub Form_Load()
    Call CenterForm(Me)
End Sub


