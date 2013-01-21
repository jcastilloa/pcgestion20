VERSION 5.00
Begin VB.Form frmParagraph1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BWord Paragraph"
   ClientHeight    =   1245
   ClientLeft      =   3075
   ClientTop       =   2535
   ClientWidth     =   3450
   Icon            =   "frmParagraph1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1245
   ScaleWidth      =   3450
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCxl 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   2220
      TabIndex        =   9
      Top             =   720
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   2220
      TabIndex        =   8
      Top             =   240
      Width           =   1155
   End
   Begin VB.PictureBox picAlignArr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   2
      Left            =   1620
      Picture         =   "frmParagraph1.frx":0442
      ScaleHeight     =   585
      ScaleWidth      =   645
      TabIndex        =   7
      Top             =   1560
      Width           =   675
   End
   Begin VB.PictureBox picAlignArr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   1
      Left            =   900
      Picture         =   "frmParagraph1.frx":0544
      ScaleHeight     =   585
      ScaleWidth      =   645
      TabIndex        =   6
      Top             =   1560
      Width           =   675
   End
   Begin VB.PictureBox picAlignArr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   180
      Picture         =   "frmParagraph1.frx":0646
      ScaleHeight     =   585
      ScaleWidth      =   645
      TabIndex        =   5
      Top             =   1560
      Width           =   675
   End
   Begin VB.Frame Frame1 
      Caption         =   "Paragraph Alignment"
      Height          =   1035
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1935
      Begin VB.PictureBox picAlign 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1440
         Picture         =   "frmParagraph1.frx":0748
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   4
         Top             =   360
         Width           =   270
      End
      Begin VB.OptionButton opnAlign 
         Caption         =   "Right"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1275
      End
      Begin VB.OptionButton opnAlign 
         Caption         =   "Center"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton opnAlign 
         Caption         =   "Left"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmParagraph1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCxl_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    If opnAlign(0).Value = True Then
        frmMain.rtfText.SelAlignment = rtfLeft
    ElseIf opnAlign(1).Value = True Then
        frmMain.rtfText.SelAlignment = rtfCenter
    Else
        opnAlign(2).Value = True
        frmMain.rtfText.SelAlignment = rtfRight
    End If
    Unload Me
End Sub
Private Sub Form_Load()
    Call CenterForm(Me)
    If opnAlign(0).Value = True Then
        frmMain.rtfText.SelAlignment = rtfLeft
    ElseIf opnAlign(1).Value = True Then
        
    Else
        opnAlign(2).Value = True
    End If
End Sub
Private Sub opnAlign_Click(intIndex As Integer)
    picAlign.Picture = picAlignArr(intIndex).Picture
End Sub


