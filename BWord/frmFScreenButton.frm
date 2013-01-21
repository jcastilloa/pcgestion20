VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFScreenButton 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Full Screen"
   ClientHeight    =   375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   510
   ControlBox      =   0   'False
   Icon            =   "frmFScreenButton.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   375
   ScaleWidth      =   510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imgFull 
      Left            =   480
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFScreenButton.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrFull 
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "imgFull"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FullScreen"
            Object.ToolTipText     =   "Close Full Screen"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFScreenButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    frmFScreen.FSRTB.SelStart = 0
    frmMain.rtfText.SelStart = 0
End Sub
Private Sub tbrFull_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "FullScreen"
            Unload frmFScreen
            Unload Me
            Clipboard.Clear
            Clipboard.SetText (strSaveClipBoard)
    End Select
End Sub

