VERSION 5.00
Begin VB.Form frmBarcode 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Code39"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Codabar 
      BackColor       =   &H00FF0000&
      Caption         =   "Codabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      Top             =   530
      Width           =   1485
   End
   Begin VB.OptionButton Code25 
      BackColor       =   &H00FF0000&
      Caption         =   "Code 2/5 Interleaved"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   775
      Width           =   2565
   End
   Begin VB.CommandButton cmdBMP 
      BackColor       =   &H00FFFF00&
      Caption         =   "Bitmap"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.OptionButton Code128 
      BackColor       =   &H00FF0000&
      Caption         =   "Code128"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   775
      Width           =   1240
   End
   Begin VB.OptionButton Code39 
      BackColor       =   &H00FF0000&
      Caption         =   "Code39"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   530
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFF00&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFF00&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtData 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   3480
      TabIndex        =   0
      Text            =   "Free Barcode"
      Top             =   120
      Width           =   3735
   End
   Begin VB.PictureBox picBarCode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   120
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   237
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1995
      Width           =   3555
   End
End
Attribute VB_Name = "frmBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Pass these values to Barcode Function: Code Type, What to Print, _
Where to print (Screen or Printer),Scale (PIXEL size to screen, TWIP size to printer), _
Height of Barcode, Starting X location, Starting Y location
Public Sub cmdBMP_Click()
SavePicture picBarCode.Image, App.Path & "\Barcode.bmp"
End Sub
Private Sub Codabar_Click()
txtData = "1234567890"
DrawBarCode "Codabar"
End Sub
Private Sub Code128_Click()
txtData = "0000100201002"
DrawBarCode "128"
End Sub
Private Sub Code39_Click()
txtData = "Free Barcode"
DrawBarCode "39"
End Sub
Private Sub Code25_Click()
txtData = "1234567890"
DrawBarCode "2/5"
End Sub
Private Sub Form_Load()
DrawBarCode "39"
End Sub
Private Sub txtData_Change()
DrawBarCode Switch(Code39.Value = True, "39", Code128.Value = True, "128", Code25.Value = True, "2/5", Codabar.Value = True, "Codabar")
End Sub
Private Sub cmdPrint_Click()
PrintBarCode Switch(Code39.Value = True, "39", Code128.Value = True, "128", Code25.Value = True, "2/5", Codabar.Value = True, "Codabar")
End Sub
Private Sub PrintBarCode(CodeType As String)
Barcode CodeType, txtData, Printer, 15, 1500, 1000, 500
Printer.EndDoc
End Sub
Public Sub DrawBarCode(CodeType As String)
picBarCode.Cls
Barcode CodeType, txtData, picBarCode, 2, 60, 20, 20
DoEvents
End Sub
Private Sub cmdExit_Click()
End
End Sub
