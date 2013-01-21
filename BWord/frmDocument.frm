VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDocument 
   Caption         =   "Comentario"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7950
   ControlBox      =   0   'False
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   7950
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   1995
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3519
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmDocument.frx":0442
      MouseIcon       =   "frmDocument.frx":04C4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()

' For displaying Line number
DisplayLineNumber
chkBullets
DocumentClosed = False
' for Context sensitive menu
frmMain.setMenu
If Left(Me.Caption, 8) <> "Comentario" Then
        frmMain.dlgCommonDialog.FileName = Me.Caption
End If
End Sub

Private Sub Form_Load()
Me.rtfText.Font.Name = GetSetting(App.Title, "Settings", "Font Name", "Times New Roman")
Me.rtfText.Font.Size = GetSetting(App.Title, "Settings", "Font Size", 10)
Me.rtfText.BackColor = GetSetting(App.Title, "Settings", "Background", &H80000005)
Me.rtfText.SelColor = GetSetting(App.Title, "Settings", "Text Color", &H80000008)
frmMain.cmbFontName = GetSetting(App.Title, "Settings", "Font Name", "Times New Roman")
frmMain.cmbFontSize = GetSetting(App.Title, "Settings", "Font Size", 10)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim SaveIt As Integer
    ' If the file needs to be saved then only save it
    
    If NeedSaved(Val(Me.Tag)) = True Then
        SaveIt = MsgBox("¿El comentario ha cambiado. ¿Deseas salvarlo?", 3 + 32 + 0)
        Select Case SaveIt
        Case vbCancel
            Cancel = True
            frmMain.setMenu
        Case vbYes
            ' If the fiile is already loaded which was aved before
            If file(Val(Me.Tag)) = "" Or Left(Me.Caption, 8) <> "Documento" Then
                DocumentClosed = True
                frmMain.mnuFileSave_Click
                On Error GoTo XSaveCancelled
    
XSaveCancelled:
            Else
                DocumentClosed = True
            ' For a new file
                frmMain.mnuFileSaveAs_Click
            End If
        Case vbNo
            ' De Initialising the form and other variables
            Set ChildForms(Val(Me.Tag)) = Nothing
            UnAvail(Val(Me.Tag)) = False
            frmCount = frmCount - 1
            frmMain.setMenu
            Unload Me
        End Select
    Else
        Set ChildForms(Val(Me.Tag)) = Nothing
        UnAvail(Val(Me.Tag)) = False
        frmCount = frmCount - 1
        frmMain.setMenu
        Unload Me
    End If
  
End Sub

Private Sub rtfText_Change()
  ' For an unchanged document Open is true and
  ' needsaved is false
  ' needsaved Teue and opened False means
  ' no changes were made since the last save
  If Opened = False And NeedSaved(Val(Me.Tag)) = False Then
    NeedSaved(Val(Me.Tag)) = True
  End If
  Opened = False
End Sub
Private Sub rtfText_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu frmMain.mnuPopUp
    End If
   If Button = vbLeftButton Then
        ' Format Painter
        FPainter
   End If
End Sub
Private Sub rtfText_SelChange()
    DisplayLineNumber
    chkBullets
    ' These codes are for the Buttons to show
    ' their statue for the current cursor location
    frmMain.tbToolBar.Buttons("Bold").Value = IIf(rtfText.SelBold, tbrPressed, tbrUnpressed)
    frmMain.tbToolBar.Buttons("Italic").Value = IIf(rtfText.SelItalic, tbrPressed, tbrUnpressed)
    frmMain.tbToolBar.Buttons("Underline").Value = IIf(rtfText.SelUnderline, tbrPressed, tbrUnpressed)
    frmMain.tbToolBar.Buttons("Align Left").Value = IIf(rtfText.SelAlignment = rtfLeft, tbrPressed, tbrUnpressed)
    frmMain.tbToolBar.Buttons("Center").Value = IIf(rtfText.SelAlignment = rtfCenter, tbrPressed, tbrUnpressed)
    frmMain.tbToolBar.Buttons("Align Right").Value = IIf(rtfText.SelAlignment = rtfRight, tbrPressed, tbrUnpressed)
    On Error Resume Next
    frmMain.cmbFontName.Text = rtfText.SelFontName
    frmMain.cmbFontSize.Text = rtfText.SelFontSize
    frmMain.setMenu
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    rtfText.Move 0, 0, Me.ScaleWidth - 5, Me.ScaleHeight - 5
    rtfText.RightMargin = rtfText.Width - 400
End Sub

Private Sub chkBullets()
    If frmDocument.rtfText.SelBullet = True Then
    frmMain.tbToolBar.Buttons("Bullets").Value = tbrPressed
    Else
    frmMain.tbToolBar.Buttons("Bullets").Value = tbrUnpressed
    End If
End Sub
