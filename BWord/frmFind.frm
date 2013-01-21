VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " BWord Buscar"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picBar 
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   -15
      ScaleHeight     =   915
      ScaleWidth      =   5745
      TabIndex        =   6
      Top             =   900
      Width           =   5745
      Begin VB.CommandButton cmdReplaceAll 
         Caption         =   "Reemplazar Todo"
         Height          =   315
         Left            =   4200
         TabIndex        =   11
         Top             =   525
         Width           =   1545
      End
      Begin VB.CommandButton cmdReplace 
         Caption         =   "&Reemplazar..."
         Height          =   315
         Left            =   4200
         TabIndex        =   10
         Top             =   150
         Width           =   1545
      End
      Begin VB.Frame Frame1 
         Caption         =   "Opciones de búsqueda"
         Height          =   855
         Left            =   90
         TabIndex        =   7
         Top             =   15
         Width           =   4110
         Begin VB.CheckBox chkWholeWord 
            Caption         =   "Buscar solo palabras completas"
            Height          =   240
            Left            =   170
            TabIndex        =   9
            Top             =   240
            Width           =   2595
         End
         Begin VB.CheckBox chkMatchCase 
            Caption         =   "Mayúsculas / Minúsculas"
            Height          =   240
            Left            =   170
            TabIndex        =   8
            Top             =   550
            Width           =   3060
         End
      End
   End
   Begin VB.ComboBox cboReplace 
      Height          =   315
      Left            =   1290
      TabIndex        =   5
      Top             =   450
      Width           =   3270
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   4590
      TabIndex        =   3
      Top             =   450
      Width           =   990
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Buscar"
      Height          =   315
      Left            =   4590
      TabIndex        =   1
      Top             =   75
      Width           =   990
   End
   Begin VB.ComboBox cboFind 
      Height          =   315
      Left            =   1290
      TabIndex        =   0
      Top             =   75
      Width           =   3270
   End
   Begin VB.Label lblReplace 
      Caption         =   "Reemplazar por:"
      Height          =   240
      Left            =   75
      TabIndex        =   4
      Top             =   525
      Width           =   1170
   End
   Begin VB.Label lblFind 
      Caption         =   "Buscar:"
      Height          =   240
      Left            =   75
      TabIndex        =   2
      Top             =   150
      Width           =   840
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdFind_Click()
    On Error GoTo FindError
    Dim lngResult As Long
    Dim lngPos As Long
    Dim intOptions As Integer
    ' Set search options
    If chkWholeWord.Value = 1 Then intOptions = intOptions + 2
    If chkMatchCase.Value = 1 Then intOptions = intOptions + 4

    If cmdFind.Caption = "&Find" Then 'If first time
        ' Get position of the searched word
       
        lngResult = frmMain.rtfText.Find(cboFind.Text, 0, , intOptions)

        If lngResult = -1 Then 'Text not found
            MsgBox "No se ha encontrado el texto", , App.Title
            cmdFind.Caption = "Buscar" 'Set caption
            frmMain.mnuEditFindNext.Enabled = False 'Disable Find Next menu
        Else 'Text found
            
            frmMain.rtfText.SetFocus 'Set focus to rtfText
            cmdReplace.Enabled = True 'Enable Replace button
            cmdReplaceAll.Enabled = True 'Enable ReplaceAll button
            cmdFind.Caption = "Buscar siguiente" 'Set caption
            frmMain.mnuEditFindNext.Enabled = True 'Enable Find Next menu
        End If
    Else 'Find Next

        lngPos = frmMain.rtfText.SelStart + frmMain.rtfText.SelLength
        lngResult = frmMain.rtfText.Find(cboFind.Text, lngPos, , intOptions)

        If lngResult = -1 Then 'Text not found
            MsgBox "Text not found", , App.Title
            cmdFind.Caption = "&Find" 'Set caption
            cmdReplace.Enabled = False 'Disable Replace button
            cmdReplaceAll.Enabled = False 'Disable ReplaceAll button
            frmMain.mnuEditFindNext.Enabled = False 'Disable Find Next menu
        Else 'Text found
            frmMain.rtfText.SetFocus 'Set focus to rtfText
            frmMain.mnuEditFindNext.Enabled = True 'Enable Find Next menu
        End If
    End If
    Exit Sub
FindError:
    MsgBox Err.Description
End Sub

Private Sub cmdReplace_Click()
    On Error GoTo ReplaceError
    Dim lngResult As Long
    Dim lngPos As Long
    Dim intOptions As Integer
 
    If cmdReplace.Caption = "&Reemplazar..." Then 'Show replace
        cmdReplace.Top = 150 'Set cmdReplace top
        cmdReplace.Caption = "&Reemplazar" 'Set caption
        lblReplace.Visible = True 'Show lblReplace
        cboReplace.Visible = True 'Show cboReplace
        cmdReplaceAll.Visible = True 'Show cmdReplaceAll
        Exit Sub
    End If

    ' Set search options
    If chkWholeWord.Value = 1 Then intOptions = intOptions + 2
    If chkMatchCase.Value = 1 Then intOptions = intOptions + 4
    If cboFind.Text = "" Then Exit Sub
    
    With frmMain
        .rtfText.SelText = cboReplace.Text 'Replace text
        ' Find next
        lngPos = .rtfText.SelStart + .rtfText.SelLength
        ' Get position of the searched word
        lngResult = .rtfText.Find(cboFind.Text, lngPos, , intOptions)

        If lngResult = -1 Then 'Text not found
            MsgBox "Text not found", , App.Title
            cmdFind.Caption = "&Find" 'Set caption
            cmdReplace.Enabled = False 'Disable Replace button
            cmdReplaceAll.Enabled = False 'Disable ReplaceAll button
        Else 'Text found
            .rtfText.SetFocus 'Set focus
        End If
    End With
    Exit Sub
ReplaceError:
    MsgBox Err.Description
End Sub

Private Sub cmdCancel_Click()
    
    Unload Me
End Sub

Private Sub cmdReplaceAll_Click()
    On Error GoTo ReplaceAllError
    Dim intCount As Integer
    Dim lngPos As Long
    Dim intOptions As Integer
    If cboFind.Text = "" Then Exit Sub
    ' Set search options
    If chkWholeWord.Value = 1 Then intOptions = intOptions + 2
    If chkMatchCase.Value = 1 Then intOptions = intOptions + 4

    intCount = 0
    lngPos = 0
    With frmMain
        Do
            If .rtfText.Find(cboFind.Text, lngPos, , intOptions) = -1 Then 'Text not fount
                If intCount > 0 Then 'Show how many replacments have been made
                    MsgBox "The specified region has been searched. " & vbCrLf & _
                    intCount & " replacements have been made.", , App.Title
                End If
                cmdFind.Caption = "&Find" 'Set caption
                cmdReplace.Enabled = False 'Disable Replace button
                cmdReplaceAll.Enabled = False 'Disable ReplaceAll button
                Exit Do
            Else 'Text found
                lngPos = .rtfText.SelStart + .rtfText.SelLength
                intCount = intCount + 1 'Increase counter by 1
                .rtfText.SelText = cboReplace.Text 'Replace text
            End If
        Loop
    End With
    Exit Sub
ReplaceAllError:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    cmdReplace.Top = 525 'Set cmdReplace top
    lblReplace.Visible = False 'Hide lblReplace
    cboReplace.Visible = False 'Hide cboReplace
    cmdReplaceAll.Visible = False 'Hide cmdReplaceAll
    
    cboFind.AddItem frmMain.rtfText.SelText 'Add selected text to find combobox
    cboFind.Text = frmMain.rtfText.SelText 'Set text in cbo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.mnuEditFindNext.Enabled = False
End Sub
