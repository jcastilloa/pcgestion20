VERSION 5.00
Begin VB.Form frmPageSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BWord Configurar Página"
   ClientHeight    =   1305
   ClientLeft      =   3075
   ClientTop       =   2730
   ClientWidth     =   2940
   Icon            =   "frmpagesetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1305
   ScaleWidth      =   2940
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1860
      TabIndex        =   6
      Top             =   780
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1860
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Margenes"
      Height          =   1155
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1695
      Begin VB.TextBox txtRight 
         Height          =   285
         Left            =   660
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtLeft 
         Height          =   285
         Left            =   660
         TabIndex        =   2
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Dcho:"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Izdo:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmPageSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo pgSetupErr
    If Val(txtLeft.Text) < 0 Then
        MsgBox "Left margin must be no less than zero inches.", 16, "Margin Out Of Range"
        txtLeft.SelStart = 0
        txtLeft.SelLength = Len(txtLeft.Text)
        txtLeft.SetFocus
    ElseIf Val(txtRight.Text) < 0 Then
        MsgBox "Right margin must be no less than zero inches.", 16, "Margin Out Of Range"
        txtRight.SelStart = 0
        txtRight.SelLength = Len(txtRight.Text)
        txtRight.SetFocus
    Else
        Dim lngOldStart As Long
        Dim lngOldLength As Long
        
        With frmMain.rtfText
        'With frmMain.ActiveForm.rtfText
            lngOldStart = .SelStart
            lngOldLength = .SelLength
            .SelStart = 0
            .SelLength = Len(.Text)
            .SelIndent = CInt(Val(txtLeft.Text) * 1440) ' in Inches
            .SelRightIndent = CInt(Val(txtRight.Text) * 1440)
            .SelStart = lngOldStart
            .SelLength = lngOldLength
        End With
      
        Unload Me
    End If
    Exit Sub
pgSetupErr:
    MsgBox "Unexpected Values.", , App.Title
End Sub

Private Sub Form_Load()
    ' This code will set the textBoxes to the current
    ' Margin Values .
    With frmMain.ActiveForm.rtfText
        Call CenterForm(Me)
        Dim sglLeft As Single
        Dim sglRight As Single
        sglLeft = .SelIndent
        sglLeft = sglLeft / 1440#
        sglLeft = CInt(sglLeft * 100#)
        sglLeft = sglLeft / 100#
        txtLeft.Text = Trim(str(sglLeft))
        sglRight = .SelRightIndent
        sglRight = sglRight / 1440#
        sglRight = CInt(sglRight * 100#)
        sglRight = sglRight / 100#
        txtRight.Text = Trim(str(sglRight))
    End With
End Sub


