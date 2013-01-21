VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7665
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOkay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Okay"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "developed by"
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "badhart@hotpop.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5640
      TabIndex        =   6
      Top             =   1440
      Width           =   1830
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.badsoft.co.uk"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5880
      TabIndex        =   5
      Top             =   1200
      Width           =   1590
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   4
      Top             =   2760
      Width           =   4215
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "©2002 BadSoft Entertainment, all rights reserved."
      Height          =   195
      Left            =   3885
      TabIndex        =   3
      Top             =   1800
      Width           =   3645
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "version"
      Height          =   195
      Left            =   2160
      TabIndex        =   1
      Top             =   600
      Width           =   525
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   1680
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   0
      Picture         =   "frmAbout.frx":012D
      Top             =   0
      Width           =   1950
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Author = "Drew (aka The Bad One)"
Private Sub Form_Load()
    lblTitle = App.ProductName
    lblVersion = "Version " & App.Major & "." & App.Minor & App.Revision
    lblAuthor = "Developed by " & Author
End Sub

Private Sub cmdOkay_Click()
    Unload Me
End Sub

