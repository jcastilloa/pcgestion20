VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " BWord - Opciones"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picFont 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   360
      ScaleHeight     =   4215
      ScaleWidth      =   3855
      TabIndex        =   3
      Top             =   720
      Width           =   3855
      Begin VB.ComboBox cmbSize 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   600
         Width           =   2535
      End
      Begin VB.ComboBox cmbFont 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   2535
      End
      Begin VB.CommandButton cmdBackground 
         Caption         =   "Background"
         Height          =   495
         Left            =   0
         TabIndex        =   6
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmdTextColor 
         Caption         =   "Text Color"
         Height          =   495
         Left            =   0
         TabIndex        =   5
         Top             =   1800
         Width           =   1695
      End
      Begin RichTextLib.RichTextBox rtfSample 
         Height          =   1455
         Left            =   0
         TabIndex        =   4
         Top             =   2640
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   2566
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         TextRTF         =   $"frmOptions.frx":030A
      End
      Begin VB.Label Label1 
         Caption         =   "Default Size"
         Height          =   375
         Left            =   0
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblFont 
         Caption         =   "Default Font"
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblBackGround 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   2040
         TabIndex        =   10
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblTextColor 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   2040
         TabIndex        =   9
         Top             =   1800
         Width           =   1455
      End
   End
   Begin MSComctlLib.TabStrip tbrOptions 
      Height          =   4935
      Left            =   165
      TabIndex        =   13
      Top             =   165
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   8705
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Font"
            Key             =   "Font"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&File Path"
            Key             =   "File"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   5280
      Width           =   975
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strPath As String
Private Sub cmbFont_Click()
    rtfSample.Font.Name = cmbFont.Text
End Sub

Private Sub cmbSize_Click()
    rtfSample.Font.Size = cmbSize.Text
End Sub

Private Sub cmdApply_Click()
    SaveSetting App.Title, "Settings", "Font Name", cmbFont.Text
    SaveSetting App.Title, "Settings", "Font Size", cmbSize.Text
    SaveSetting App.Title, "Settings", "Background", lblBackGround.BackColor
    SaveSetting App.Title, "Settings", "Text Color", lblTextColor.BackColor
    SaveSetting App.Title, "Settings", "Default Dir", strPath
    lblCurrent.Caption = strPath
    frmMain.dlgCommonDialog.FileName = strPath & "\" & "Document"
End Sub

Private Sub cmdBackground_Click()
    On Error GoTo errHdlr
    frmMain.dlgCommonDialog.ShowColor
    lblBackGround.BackColor = frmMain.dlgCommonDialog.Color
    rtfSample.BackColor = lblBackGround.BackColor
    Exit Sub
errHdlr:
    End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdTextColor_Click()
    On Error GoTo errHdlr
    frmMain.dlgCommonDialog.ShowColor
    lblTextColor.BackColor = frmMain.dlgCommonDialog.Color
    rtfSample.SelStart = 0
    rtfSample.SelLength = Len(rtfSample.Text)
    rtfSample.SelColor = lblTextColor.BackColor
    rtfSample.SelStart = 0
    Exit Sub
errHdlr:
    
End Sub

Private Sub Dir1_Change()
    strPath = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub


Private Sub Form_Load()
    Dim strSize As String
' Fill the combo box with available fonts in the system
    lblBackGround.BackColor = frmMain.rtfText.BackColor
    lblTextColor.BackColor = frmMain.rtfText.SelColor
    For x = 1 To Screen.FontCount
        cmbFont.AddItem Screen.Fonts(x)
    Next
    lblCurrent.Caption = GetSetting(App.Title, "Settings", "Default Dir", "C:\")
    Drive1.Drive = Left(lblCurrent.Caption, 3)
    Dir1.Path = lblCurrent.Caption
 ' Fill the combo box with sizes 1 to 72 to choose from
    For x = 5 To 72: cmbSize.AddItem str$(x): Next
    
    cmbFont.Text = GetSetting(App.Title, "Settings", "Font Name", "Times New Roman")
    strSize = GetSetting(App.Title, "Settings", "Font Size", "10")
      
    For x = 0 To cmbSize.ListCount - 1
        If Int(Val(strSize)) = Val(cmbSize.List(x)) Then
            cmbSize.ListIndex = x
            Exit For
        End If
    Next
    
    
    
    rtfSample.BackColor = lblBackGround.BackColor
    rtfSample.SelStart = 0
    rtfSample.SelLength = Len(rtfSample.Text)
    rtfSample.SelColor = lblTextColor.BackColor
    rtfSample.SelLength = 0
End Sub



Private Sub tbrOptions_Click()
If tbrOptions.SelectedItem.Key = "Font" Then
        picFont.Visible = True
        picDefDir.Visible = False
    ElseIf tbrOptions.SelectedItem.Key = "File" Then
        picFont.Visible = False
        picDefDir.Visible = True
    End If
End Sub
