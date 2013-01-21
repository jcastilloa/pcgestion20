VERSION 5.00
Object = "{CCA214C0-DFEB-4C91-9F0D-2665F77F6E23}#1.2#0"; "IDAutomationLinear.dll"
Begin VB.Form BarCodefrm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Código de Barras"
   ClientHeight    =   600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3330
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   40
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   222
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   465
      Left            =   1530
      ScaleHeight     =   405
      ScaleWidth      =   1185
      TabIndex        =   1
      Top             =   1005
      Width           =   1245
   End
   Begin ATLCONTROLLibCtl.BarCode BarCode1 
      Height          =   585
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   3285
      _cx             =   5794
      _cy             =   1032
      Enabled         =   -1  'True
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      SymbologyId     =   13
      DataToEncode    =   "0000100302007"
      Orientation     =   0
      BarHeight       =   0,8
      NarrowBarWidth  =   0,03
      Wide2NarrowRatio=   2
      AddCheckDigit   =   1
      AddCheckDigitToText=   1
      Code128CharSet  =   1
      UPCESystem      =   0
      EANUPCSupplement=   2
      ShowText        =   0
      CodabarStartCharacter=   "A"
      CodabarStopCharacter=   "B"
      LeftMarginCM    =   0
      TopMarginCM     =   0
      SupplementToEncode=   "07"
   End
End
Attribute VB_Name = "BarCodefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public isEAN As Boolean

Private Sub Form_DblClick()
Mainfrm.ActiveForm.Width = InputBox("Cambiar el ancho", "Nuevo ancho", Mainfrm.ActiveForm.Width)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Mainfrm.ActiveForm.isEAN = True Then
If KeyCode = vbKeyAdd Then
    Mainfrm.Text1.Text = Add(Mainfrm.Text1.Text)
    Mainfrm.ActiveForm.Width = (Options.Text1.Text)
    Mainfrm.ActiveForm.Cls
    Mainfrm.ActiveForm.ScaleMode = 3
    Mainfrm.ActiveForm.Label1.Visible = True
    PaintCode Mainfrm.ActiveForm, Mid$(Mainfrm.Text1.Text, 1, 1), Mid$(Mainfrm.Text1.Text, 2, 6), Mid$(Mainfrm.Text1.Text, 8, 6)
    Mainfrm.ActiveForm.Label1.Caption = Mid$(Mainfrm.Text1.Text, 1, 1)
    Mainfrm.ActiveForm.Label2.Caption = Mid$(Mainfrm.Text1.Text, 2, 6)
    Mainfrm.ActiveForm.Label3.Caption = Mid$(Mainfrm.Text1.Text, 8, 6)
    Mainfrm.ActiveForm.isEAN = True
    Mainfrm.ActiveForm.Refresh
ElseIf KeyCode = vbKeySubtract Then
    Mainfrm.Text1.Text = Subt(Mainfrm.Text1.Text)
    Mainfrm.ActiveForm.Width = (Options.Text1.Text)
    Mainfrm.ActiveForm.Cls
    Mainfrm.ActiveForm.ScaleMode = 3
    Mainfrm.ActiveForm.Label1.Visible = True
    PaintCode Mainfrm.ActiveForm, Mid$(Mainfrm.Text1.Text, 1, 1), Mid$(Mainfrm.Text1.Text, 2, 6), Mid$(Mainfrm.Text1.Text, 8, 6)
    Mainfrm.ActiveForm.Label1.Caption = Mid$(Mainfrm.Text1.Text, 1, 1)
    Mainfrm.ActiveForm.Label2.Caption = Mid$(Mainfrm.Text1.Text, 2, 6)
    Mainfrm.ActiveForm.Label3.Caption = Mid$(Mainfrm.Text1.Text, 8, 6)
    Mainfrm.ActiveForm.isEAN = True
    Mainfrm.ActiveForm.Refresh
ElseIf KeyCode = vbKeyReturn Then
    newBarCode
    Mainfrm.ActiveForm.Width = (Options.Text1.Text)
    Mainfrm.ActiveForm.Cls
    Mainfrm.ActiveForm.ScaleMode = 3
    Mainfrm.ActiveForm.Label1.Visible = True
    PaintCode Mainfrm.ActiveForm, Mid$(Mainfrm.Text1.Text, 1, 1), Mid$(Mainfrm.Text1.Text, 2, 6), Mid$(Mainfrm.Text1.Text, 8, 6)
    Mainfrm.ActiveForm.Label1.Caption = Mid$(Mainfrm.Text1.Text, 1, 1)
    Mainfrm.ActiveForm.Label2.Caption = Mid$(Mainfrm.Text1.Text, 2, 6)
    Mainfrm.ActiveForm.Label3.Caption = Mid$(Mainfrm.Text1.Text, 8, 6)
    Mainfrm.ActiveForm.isEAN = True
    Mainfrm.ActiveForm.Refresh
End If
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    Me.PopupMenu menues.C, , X, Y
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'FState(Me.Tag).Deleted = True
End Sub
