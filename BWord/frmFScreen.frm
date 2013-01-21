VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmFScreen 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   ClientHeight    =   5145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6810
   Icon            =   "frmFScreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox FSRTB 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5318
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmFScreen.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmFScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bChangeFSRTB As Boolean
Private Sub Form_Load()
    Form_Resize
    FSRTB.SelStart = 0
    frmFScreenButton.Show , Me
    frmFScreen.FSRTB.BackColor = GetSetting(App.Title, "Settings", "Background", &H80000005)
End Sub
Private Sub Form_Resize()
    Me.Left = -30
    Me.Top = -30
    Me.Height = Screen.Height + 30
    Me.Width = Screen.Width + 30
    FSRTB.Width = Me.Width - 30
    FSRTB.Height = Me.Height - 30
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If bChangeFSRTB Then
        'Select all text
        FSRTB.SelStart = 0
        FSRTB.SelLength = Len(FSRTB.Text)
        'Then copy it
        SendMessage FSRTB.hwnd, WM_COPY, 0, 0&
        'And paste it on the Textbox
        frmMain.rtfText.Text = ""
        SendMessage frmMain.rtfText.hwnd, WM_PASTE, 0, 0&
        frmMain.rtfText.SelStart = 0
        FSRTB.Text = ""
        Exit Sub
    End If
End Sub
Private Sub FSRTB_Change()
    bChangeFSRTB = True
End Sub
