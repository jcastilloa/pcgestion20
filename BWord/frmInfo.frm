VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Information"
   ClientHeight    =   4425
   ClientLeft      =   3105
   ClientTop       =   3330
   ClientWidth     =   6465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4425
   ScaleWidth      =   6465
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Height          =   1800
      Left            =   2640
      ScaleHeight     =   1740
      ScaleWidth      =   1695
      TabIndex        =   21
      Top             =   2520
      Width           =   1755
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   20
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtsize 
      Height          =   375
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox txtname 
      Height          =   375
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   120
      Width           =   5415
   End
   Begin VB.Frame fradates 
      Caption         =   "Dates"
      Height          =   1335
      Left            =   2040
      TabIndex        =   9
      Top             =   1080
      Width           =   4335
      Begin VB.Label lbldate 
         Caption         =   "00/00/0000 00:00:00"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   15
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lbldatetxt 
         Caption         =   "Created:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lbldatetxt 
         Caption         =   "Last Modified:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lbldate 
         Caption         =   "00/00/0000 00:00:00"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   12
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label lbldatetxt 
         Caption         =   "Last Accessed:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lbldate 
         Caption         =   "00/00/0000 00:00:00"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   10
         Top             =   960
         Width           =   1695
      End
   End
   Begin VB.CommandButton ok 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Frame fraattrib 
      Caption         =   "Attributes"
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
      Begin VB.CheckBox attributes 
         Caption         =   "Normal"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CheckBox attributes 
         Caption         =   "Compressed"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CheckBox attributes 
         Caption         =   "Temporary"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox attributes 
         Caption         =   "Archive"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox attributes 
         Caption         =   "Read Only"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox attributes 
         Caption         =   "System"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox attributes 
         Caption         =   "Hidden"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog dialog 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "File Information"
      Filter          =   "All Files (*.*) | *.*"
      Flags           =   2101252
   End
   Begin VB.Label lblsize 
      Caption         =   "Size:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblname 
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub updatestats()
'Dim ftime As SYSTEMTIME
End Sub

Private Sub attributes_GotFocus(Index As Integer)
ok.SetFocus
End Sub


Private Sub cmdBrowse_Click()
On Error GoTo errhand
dialog.ShowOpen
updatestats
Me.Show
ok.SetFocus
Exit Sub

errhand:

End Sub

Private Sub ok_Click()
Unload Me
End Sub


