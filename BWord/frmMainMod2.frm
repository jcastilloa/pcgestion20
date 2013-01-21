VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H8000000C&
   ClientHeight    =   6240
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10080
   Icon            =   "frmMainMod2.frx":0000
   LinkTopic       =   "MDIForm1"
   ScaleHeight     =   6240
   ScaleWidth      =   10080
   StartUpPosition =   2  'CenterScreen
   Begin miBwordOCX.epCmDlg epCmDlg1 
      Left            =   4680
      Top             =   2595
      _ExtentX        =   661
      _ExtentY        =   635
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   741
      _CBWidth        =   10080
      _CBHeight       =   420
      _Version        =   "6.7.8988"
      MinHeight1      =   360
      Width1          =   6810
      NewRow1         =   0   'False
      MinHeight2      =   360
      Width2          =   1260
      NewRow2         =   0   'False
      MinHeight3      =   360
      NewRow3         =   0   'False
      BandStyle3      =   1
      Begin MSComctlLib.ProgressBar prgBar 
         Height          =   255
         Left            =   6960
         TabIndex        =   9
         Top             =   90
         Visible         =   0   'False
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.ComboBox cmbFontSize 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmMainMod2.frx":0442
         Left            =   2280
         List            =   "frmMainMod2.frx":0444
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "cmbFontSize"
         Top             =   40
         Width           =   735
      End
      Begin VB.ComboBox cmbFontName 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "cmbFontName"
         Top             =   45
         Width           =   1935
      End
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   22
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Preview"
            Object.ToolTipText     =   "Visualizar Impresión"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cortar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copiar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Pegar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Negrita"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Cursiva"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Subrayado"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Left"
            Object.ToolTipText     =   "Alinear Izquierda"
            ImageIndex      =   11
            Style           =   2
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Object.ToolTipText     =   "Centrar"
            ImageIndex      =   12
            Style           =   2
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Right"
            Object.ToolTipText     =   "Alinear Derecha"
            ImageIndex      =   13
            Style           =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   16
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bullets"
            Object.ToolTipText     =   "Insertar Marcas"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Time"
            Object.ToolTipText     =   "Insertar Fecha/Hora"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   5865
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10839
            MinWidth        =   4410
            Text            =   "Status"
            TextSave        =   "Status"
            Key             =   "Text"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   1693
            MinWidth        =   1499
            TextSave        =   "30/12/2003"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1508
            MinWidth        =   1499
            TextSave        =   "11:07"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1935
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   2880
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod2.frx":0446
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod2.frx":0558
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod2.frx":066A
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod2.frx":077C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod2.frx":088E
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod2.frx":09A0
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod2.frx":0AB2
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod2.frx":0BC4
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod2.frx":0CD6
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod2.frx":0DE8
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod2.frx":0EFA
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod2.frx":100C
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod2.frx":111E
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod2.frx":1230
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod2.frx":1344
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod2.frx":1888
            Key             =   "Painter"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod2.frx":1DCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod2.frx":2310
            Key             =   "Spelling"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod2.frx":2854
            Key             =   "Bullets"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod2.frx":2E3C
            Key             =   "FSreen"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod2.frx":3508
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar2 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   780
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   741
      BandCount       =   2
      _CBWidth        =   10080
      _CBHeight       =   420
      _Version        =   "6.7.8988"
      MinHeight1      =   360
      Width1          =   9000
      NewRow1         =   0   'False
      MinHeight2      =   360
      NewRow2         =   0   'False
      BandStyle2      =   1
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   480
         Picture         =   "frmMainMod2.frx":3B44
         ScaleHeight     =   270
         ScaleWidth      =   11490
         TabIndex        =   7
         Top             =   105
         Width           =   11490
      End
   End
   Begin ComCtl3.CoolBar CoolBar3 
      Align           =   3  'Align Left
      Height          =   4665
      Left            =   0
      TabIndex        =   8
      Top             =   1200
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   8229
      BandCount       =   1
      Orientation     =   1
      _CBWidth        =   390
      _CBHeight       =   4665
      _Version        =   "6.7.8988"
      Child1          =   "Toolbar1"
      MinHeight1      =   330
      Width1          =   615
      NewRow1         =   0   'False
   End
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   1995
      Left            =   390
      TabIndex        =   0
      Top             =   1215
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3519
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMainMod2.frx":DD3E
      MouseIcon       =   "frmMainMod2.frx":DDBF
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Fichero"
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Salvar"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintConf 
         Caption         =   "Configurar Impresión"
         Begin VB.Menu mnuFilePageSetup 
            Caption         =   "Configurar pagina..."
         End
         Begin VB.Menu mnuFilePrintSetup 
            Caption         =   "Confirugar impresora..."
         End
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Imprimir..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Editar"
      Begin VB.Menu mnuEditCut 
         Caption         =   "C&ortar"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Pegar"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Seleccionar &Todo"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Buscar y Reemplazar ..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditFindNext 
         Caption         =   "B&uscar Siguiente"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&Ver"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewRuler 
         Caption         =   "Re&gla"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewPrintPreview 
         Caption         =   "Previsualizar impresión"
      End
      Begin VB.Menu mnuViewFullScreen 
         Caption         =   "&Pantalla Completa"
         Shortcut        =   {F11}
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insertar"
      Begin VB.Menu mnuInsertTimeDate 
         Caption         =   "&Fecha y Hora..."
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuInsertPicture 
         Caption         =   "&Imagen desde Fichero..."
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "&Formato"
      Begin VB.Menu mnuFormatFont 
         Caption         =   "&Fuente..."
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuFormatPageSetup 
         Caption         =   "Con&figurar Página..."
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuFormatParagraph 
         Caption         =   "&Párrafo..."
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuFormatSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatBullets 
         Caption         =   "&Marcas"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFormatChangeCase 
         Caption         =   "Cambia&r Caracter"
         Begin VB.Menu mnuFormatCSLower 
            Caption         =   "&minúscula"
            Shortcut        =   ^L
         End
         Begin VB.Menu mnuFormatCSUpper 
            Caption         =   "&MAYÚSCULA"
            Shortcut        =   ^U
         End
         Begin VB.Menu mnuFormatCSSC 
            Caption         =   "&Primera letra May"
            Shortcut        =   ^J
         End
      End
      Begin VB.Menu mnuFormatScript 
         Caption         =   "&Script"
         Begin VB.Menu mnuFormatSubScript 
            Caption         =   "Su&bScript"
         End
         Begin VB.Menu mnuFormatSuperScript 
            Caption         =   "S&uperScript"
         End
         Begin VB.Menu mnuFormatNoScript 
            Caption         =   "&No Script"
         End
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Herramientas"
      Begin VB.Menu mnuToolsEncrypt 
         Caption         =   "&Cifrar texto..."
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuToolsDecrypt 
         Caption         =   "&Descifrar texto..."
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuPopCut 
         Caption         =   "Cortar"
      End
      Begin VB.Menu mnuPopCopy 
         Caption         =   "Copiar"
      End
      Begin VB.Menu mnuPopPaste 
         Caption         =   "Pegar"
      End
      Begin VB.Menu mnuPopSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopFont 
         Caption         =   "Fuente"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 

Private Sub cmbFontSize_Change()

If cmbFontSize.ListIndex > -1 Then
rtfText.SelFontSize = cmbFontSize.List(cmbFontSize.ListIndex)
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Set frmMain = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Stores back the window position and size to registry
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
        
        On Error Resume Next
        Unload frmFind
        Unload AboutBox
        Unload frmDateTime
        Unload frmPageSetup
        Unload frmParagraph1
        Unload SpellIt
        Unload frmSendImage
        Unload frmFScreen
        Unload frmFScreenButton
End Sub

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

Private Sub chkBullets()
    If frmMain.rtfText.SelBullet = True Then
    frmMain.tbToolBar.Buttons("Bullets").Value = tbrPressed
    Else
    frmMain.tbToolBar.Buttons("Bullets").Value = tbrUnpressed
    End If
End Sub


Private Sub Form_Load()

  ' Fill the combo box with available fonts in the system
    For x = 1 To Screen.FontCount
        cmbFontName.AddItem Screen.Fonts(x)
    Next
    
    ' Fill the combo box with sizes 1 to 72 to choose from
    For x = 5 To 72: cmbFontSize.AddItem Str$(x): Next
    
Me.rtfText.Font.Name = GetSetting(App.Title, "Settings", "Font Name", "Times New Roman")
Me.rtfText.Font.Size = GetSetting(App.Title, "Settings", "Font Size", 12)
Me.rtfText.BackColor = GetSetting(App.Title, "Settings", "Background", &H80000005)
Me.rtfText.SelColor = GetSetting(App.Title, "Settings", "Text Color", &H80000008)
frmMain.cmbFontName = GetSetting(App.Title, "Settings", "Font Name", "Times New Roman")
frmMain.cmbFontSize = GetSetting(App.Title, "Settings", "Font Size", 12)

cmbFontSize.AddItem "14"
cmbFontSize.AddItem "16"
cmbFontSize.AddItem "18"
cmbFontSize.AddItem "20"
cmbFontSize.AddItem "22"
cmbFontSize.AddItem "24"

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
Private Sub rtfText_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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
    
    
    'rtfText.Move 0, 0, Me.ScaleWidth - 100, Me.ScaleHeight - 100
    'rtfText.RightMargin = rtfText.Width - 400
    
   ' Picture1.Width = Me.Width - 450
    rtfText.Width = Me.Width - 540
    rtfText.Height = Me.Height - 2430
    
    
End Sub

Private Sub mnuEditFind_Click()
    frmFind.Show , Me
End Sub

Public Sub mnuEditFindNext_Click()
    ' Check to see if there are any open documents, if not show error
    'If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    On Error GoTo FindNextError
    Dim lngResult As Integer
    Dim lngPos As Integer
    Dim intOptions As Integer
    ' Set search options
    If frmFind.chkWholeWord.Value = 1 Then intOptions = intOptions + 2
    If frmFind.chkMatchCase.Value = 1 Then intOptions = intOptions + 4

    lngPos = rtfText.SelStart + rtfText.SelLength
    ' Get position of the searched word
    lngResult = rtfText.Find(frmFind.cboFind.Text, lngPos, , intOptions)
    If lngResult = -1 Then 'Text not found
        MsgBox "No se ha encontrado el texto", , App.Title
        frmFind.cmdFind.Caption = "Buscar" 'Set caption
        frmFind.cmdReplace.Enabled = False 'Disable Replace button
        frmFind.cmdReplaceAll.Enabled = False 'Disable ReplaceAll button
        mnuEditFindNext.Enabled = False 'Disable Find Next menu
    Else
        rtfText.SetFocus 'Set focus
    End If
    Exit Sub
FindNextError:
    MsgBox Err.Description
End Sub
Private Sub mnuEditSelectAll_Click()
    rtfText.SelStart = 0
    rtfText.SelLength = Len(rtfText.Text)
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileMRU_Click(Index As Integer)
    'OpenFile mnuFileMRU(Index).Caption
    storeMRU (mnuFileMRU(Index).Caption)
    storeMRUinReg
End Sub


Private Sub mnuFileOpen_Click()
    'Dim FileStr As String, FileN As String
    'Dim TempStr As String, DotPos As Integer
    On Error GoTo OpenCanceled
    
    ' The file types that BWord will open
    frmMain.dlgCommonDialog.Filter = "RTF Files (*.rtf)|*.rtf|Text files (*.txt)|*.txt|Ini Files (*.ini)|*.ini|Registry Files (*.log)|*.log|Batch File (*.bat)|*.bat|All files (*.*)|*.*"
    
    ' &H4 prevents users from opening a readonly file
    ' cdlOFNHideReadOnly can also be used instead.
    dlgCommonDialog.flags = &H4
    dlgCommonDialog.ShowOpen
    
    If dlgCommonDialog.FileName = "" Then Exit Sub
'    OpenFile dlgCommonDialog.FileName
    Opened = True
    storeMRU (dlgCommonDialog.FileName)
    storeMRUinReg
    ' These array variables come handy in Save As Routine
    file(frm) = dlgCommonDialog.FileTitle
OpenCanceled:

End Sub

Private Sub mnuFilePageSetup_Click()

epCmDlg1.ShowPageSetup 'Show Page Setup dialog
End Sub

Private Sub mnuFilePrintSetup_Click()
epCmDlg1.ShowPrinter 'Show Printer dialog
End Sub

Public Sub mnuFileSave_Click()

If rtfText.Locked = False Then

    If Trim(rtfText.Text) = "" Then
        miCampo.Value = Null
    Else
        miCampo.Value = rtfText.TextRTF
    End If

End If

    ' Setting flags for readOnly and Overwrite prompt
'    dlgCommonDialog.flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
'
'    On Error Resume Next
'
'    ' If file name is Blank or its a New Document
'    If dlgCommonDialog.FileName = "" Or Left(Me.Caption, 8) = "Document" Then
        
'        mnuFileSaveAs_Click
'
'    ' if document caption and file name are same
'    ' which means save dialog box need not appear
'    ' Save the file as it is with the same name ie. Update the file
'    ElseIf Me.Caption = dlgCommonDialog.FileName Then
        ' If RTF save it in RTF format else in text format
'        If UCase(Right(dlgCommonDialog.FileName, 3)) = "RTF" Then
'            Me.rtfText.SaveFile dlgCommonDialog.FileName, rtfRTF
'        Else
'            Me.rtfText.SaveFile dlgCommonDialog.FileName, rtfText
'        End If
'
'        ' File already saved so no need to save again
'        ' So that if we close the file now it will
'        ' not prompt the user to save again
'        If NeedSaved(frm) = True Then
'            NeedSaved(frm) = False
'            Me.Caption = Right$(Me.Caption, Len(Me.Caption))
'        End If
        
        ' If somehow user doesnt save and exits
        ' Formcount is decremented
'        ' and menus refreshed
'        If DocumentClosed = True Then
'            frmCount = frmCount - 1
'            setMenu
'        End If
'End If
End Sub

Public Sub mnuFileSaveAs_Click()
    
    ' Look at the Save code
    ' Its very much similar
       
    On Error GoTo SaveCancelled
    dlgCommonDialog.flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    frmMain.dlgCommonDialog.Filter = "RTF Files (*.rtf)|*.rtf|Text files (*.txt)|*.txt|Ini Files (*.ini)|*.ini|Registry Files (*.log)|*.log|Batch File (*.bat)|*.bat|All files (*.*)|*.*"
    
    dlgCommonDialog.ShowSave
    
    If dlgCommonDialog.FileName = "" Then Exit Sub
  
    If UCase(Right(dlgCommonDialog.FileName, 3)) = "RTF" Then
        Me.rtfText.SaveFile dlgCommonDialog.FileName, rtfRTF
        storeMRU (dlgCommonDialog.FileName)
        storeMRUinReg
    Else
        rtfText.SaveFile dlgCommonDialog.FileName, rtfText
        storeMRU (dlgCommonDialog.FileName)
        storeMRUinReg
    End If
    
    If NeedSaved(frm) = True Then
        NeedSaved(frm) = False
        Me.Caption = Right$(Me.Caption, Len(Me.Caption))
    End If
  
    Me.Caption = dlgCommonDialog.FileTitle
    'file(frm) = dlgCommonDialog.FileTitle
    
    If DocumentClosed = True Then
        frmCount = frmCount - 1
        setMenu
    End If
Exit Sub

SaveCancelled:
    If DocumentClosed = True Then
        frmCount = frmCount - 1
        setMenu
    End If
End Sub

Private Sub mnuFormatBullets_Click()
    ' Passing the current form as parameter
    ' to the function
    Menu_FormatBullet Me
End Sub
Private Sub mnuFormatCSLower_Click()
    ' LowerCase Conversion
    Clipboard.SetText rtfText.SelText
    rtfText.SelText = LCase(Clipboard.GetText)
    Clipboard.Clear
End Sub

Private Sub mnuFormatCSSC_Click()
    ' Sentence Case Conversion
    Clipboard.SetText rtfText.SelText
    rtfText.SelText = StrConv(Clipboard.GetText, vbProperCase)
    Clipboard.Clear
End Sub

Private Sub mnuFormatCSUpper_Click()
    ' UpperCase Conversion
    Clipboard.SetText rtfText.SelText
    rtfText.SelText = UCase(Clipboard.GetText)
    Clipboard.Clear
End Sub

Private Sub mnuFormatFont_Click()
    ' cdlCFBoth will Display both Screen and Printer Fonts
    ' cdlCFEffects will display all the attributes
    ' including Underline , StrikeThru and Color
    dlgCommonDialog.flags = cdlCFBoth Or cdlCFEffects
    
    
    With rtfText
        ' Update the Font Dialog Box with
        ' the Current Font Format of the selected text
        ' On Error GoTo fnterr
        dlgCommonDialog.FontName = .SelFontName
        dlgCommonDialog.FontSize = .SelFontSize
        dlgCommonDialog.FontBold = .SelBold
        dlgCommonDialog.FontItalic = .SelItalic
        dlgCommonDialog.FontUnderline = .SelUnderline
        dlgCommonDialog.Color = .SelColor
        
        On Error Resume Next
        dlgCommonDialog.ShowFont
       
        'If dlgCommonDialog.CancelError = False Then Exit Sub
        
        
        ' Apply the settings in the
        ' Font Dialog Box to the selected text
        .SelFontName = dlgCommonDialog.FontName
        .SelFontSize = dlgCommonDialog.FontSize
        .SelBold = dlgCommonDialog.FontBold
        .SelItalic = dlgCommonDialog.FontItalic
        .SelStrikeThru = dlgCommonDialog.FontStrikeThru
        .SelUnderline = dlgCommonDialog.FontUnderline
        .SelColor = dlgCommonDialog.Color
        
        cmbFontName.Text = .SelFontName
        cmbFontSize.Text = .SelFontSize
        
    End With
    Exit Sub
fnterr:
End Sub

Private Sub mnuFormatNoScript_Click()
    rtfText.SelCharOffset = 0
End Sub

Private Sub mnuFormatPageSetup_Click()
    frmPageSetup.Show
End Sub

Private Sub mnuFormatParagraph_Click()
    frmParagraph1.Show
End Sub

Private Sub mnuFormatSubScript_Click()
    rtfText.SelCharOffset = -55
End Sub

Private Sub mnuFormatSuperScript_Click()
    rtfText.SelCharOffset = 55
End Sub

Private Sub mnuHelpAbout_Click()
    AboutBox.Show
End Sub
Private Sub mnuInsertPicture_Click()
    frmSendImage.Show
End Sub

Private Sub mnuInsertTimeDate_Click()
    frmDateTime.Show
End Sub
Private Sub mnuPopCopy_Click()
    mnuEditCopy_Click
End Sub
Private Sub mnuPopCut_Click()
    mnuEditCut_Click
End Sub
Private Sub mnuPopFont_Click()
    mnuFormatFont_Click
End Sub
Private Sub mnuPopPaste_Click()
    mnuEditPaste_Click
End Sub
Private Sub cmbFontName_Click()
    ' Change the Font Name with user selection
    rtfText.SelFontName = cmbFontName.List(cmbFontName.ListIndex)
End Sub

Private Sub cmbFontSize_Click()
    ' Change the Font Size.
    rtfText.SelFontSize = cmbFontSize.List(cmbFontSize.ListIndex)
End Sub
Private Sub mnuToolsDecrypt_Click()
    Dim EncStr As String  ' Encryption String
    Dim EncKey As String  ' Encryption Key
    Dim TempEncKey As String   ' Temperory to Swap
    Dim EncLen As Integer  ' Length of String
    Dim EncPos As Integer  ' Current Position
    Dim EncKeyPos As Integer  ' Current Pos of Key
    Dim tempChar As String
    ' TC = TA Xor TB
    Dim TA As Integer, TB As Integer, TC As Integer
  

    TempEncKey = InputBox("Introduzca la clave (imprescindible para el descifrado):", "Descifrar")
    
    If TempEncKey = "" Then Exit Sub
    
    EncStr = ""   ' Initialise
    EncPos = 1
    EncKeyPos = 1
    
    prgBar.Visible = True
    sbStatusBar.Panels(1).Text = "Descifrando...."
    
    ' Real Encryption Key is now Created
    ' With this Algorithm
    ' Thanks to a Book on Algorithms
    For x = 1 To Len(TempEncKey)
        EncKey = EncKey & Asc(Mid$(TempEncKey, x, 1))
    Next
    
    ' Length of the Key
    EncLen = Len(EncKey)
    
    ' Set the Prograss Bar
    prgBar.Min = 0
    prgBar.Max = Len(rtfText.Text)
    
    ' Encryption will be Character by Character
    For x = 1 To Len(rtfText.Text) Step 8
        TB = Asc(Mid$(EncKey, EncKeyPos, 1))
        EncKeyPos = EncKeyPos + 1
        
        If EncKeyPos > EncLen Then EncKeyPos = 1
        
        ' Binary code of a character is
        ' 8 characters long
        tempChar = Mid$(rtfText.Text, x, 8)
    
        ' BintoDec is a function to convert
        ' string to Binary
        ' Defination in Module1
        TA = BintoDec(tempChar)
        ' Main Encryption occurs here
        ' With the Xor operator
        TC = TB Xor TA
        
        ' Now Concatenate the Encrypted text
        EncStr = EncStr & Chr$(TC)
        
        ' Increment the progressbar
        prgBar.Value = x
    Next
  
    ' Paste the Encrypted text in the Document
    Me.rtfText.Text = EncStr
    
    prgBar.Visible = False
    DisplayLineNumber
End Sub

Private Sub mnuToolsEncrypt_Click()
    Dim Warning As VbMsgBoxResult
    ' Yes =6  No = 7 Cancel = 2
  
    Dim EncStr As String
    Dim EncKey As String, TempEncKey As String
    Dim EncLen As Integer
    Dim EncPos As Integer
    Dim EncKeyPos As Integer
    Dim tempChar As String
    Dim TA As Integer, TB As Integer, TC As Integer

    Warning = MsgBox("Puede perder la información si olvidas la contraseña. ¿Continuar?", vbYesNoCancel + vbExclamation, "Editor")
  
    If Warning = vbNo Or Warning = vbCancel Then
        Exit Sub
    End If
  
    TempEncKey = InputBox("Introduzca clave (necesaria para descifrar la información):", "")
  
    If TempEncKey = "" Then Exit Sub
  
    EncStr = ""
    EncPos = 1
    EncKeyPos = 1
    prgBar.Visible = True
  
    For x = 1 To Len(TempEncKey)
        EncKey = EncKey & Asc(Mid$(TempEncKey, x, 1))
    Next

    EncLen = Len(EncKey)
  
    sbStatusBar.Panels(1).Text = "Cifrando texto..."
    prgBar.Min = 0
    prgBar.Max = Len(rtfText.Text)
  
    For x = 1 To Len(rtfText.Text)
        TB = Asc(Mid$(EncKey, EncKeyPos, 1))
        EncKeyPos = EncKeyPos + 1
        
        If EncKeyPos > EncLen Then
            EncKeyPos = 1
        End If
        
        TA = Asc(Mid$(rtfText.Text, x, 1))
        TC = TB Xor TA
        ' Get Binary is Function to convert
        ' Binary to text again
        tempChar = GetBinary(TC)
        EncStr = EncStr & tempChar
        prgBar.Value = x
    Next
    
    rtfText.Text = EncStr
    prgBar.Visible = False
    DisplayLineNumber

End Sub

Private Sub mnuToolsOptions_Click()
   frmOptions.Show vbModal
End Sub

Private Sub mnuViewFullScreen_Click()
    On Error Resume Next
    strSaveClipBoard = Clipboard.GetText
    Clipboard.Clear
    If Len(rtfText.Text) = 0 Then
        MsgBox "Cannot view Full Screen " + vbNewLine + "There is no text in the Document."
        Exit Sub
    End If
    mnuEditSelectAll_Click
    SendMessage rtfText.hwnd, WM_COPY, 0, 0&
    frmFScreen.Show
    SendMessage frmFScreen.FSRTB.hwnd, WM_PASTE, 0, 0&
    frmFScreen.FSRTB.SelStart = 0
End Sub

Private Sub mnuViewPrintPreview_Click()
    frmDocPreview.Show vbModal
    If gprint = True Then
         frmDocPreview.DocPrintProc
    End If
End Sub

Private Sub mnuViewRuler_Click()
    If mnuViewRuler.Checked = True Then
        mnuViewRuler.Checked = False
        Picture1.Visible = False
    Else
        mnuViewRuler.Checked = True
        Picture1.Visible = True
    End If
End Sub

Private Sub mnuWindowArrangeIcons_Click()
    'Me.Arrange vbArrangeIcons
End Sub
Private Sub mnuWindowTileVertical_Click()
  '  Me.Arrange vbTileVertical
End Sub
Private Sub mnuWindowTileHorizontal_Click()
  '  Me.Arrange vbTileHorizontal
End Sub
Private Sub mnuWindowCascade_Click()
   ' Me.Arrange vbCascade
End Sub
Private Sub mnuWindowNewWindow_Click()
'    mnuFileNew_Click
End Sub
Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub
Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
End Sub
Private Sub mnuEditPaste_Click()
    On Error Resume Next
    rtfText.SelRTF = Clipboard.GetText
End Sub
Private Sub mnuEditCopy_Click()
    On Error Resume Next
    Clipboard.SetText rtfText.SelText
End Sub
Private Sub mnuEditCut_Click()
    On Error Resume Next
    Clipboard.SetText rtfText.SelText
    rtfText.SelText = vbNullString
End Sub

Private Sub Picture1_DblClick()
    If frmCount >= 1 Then
        frmPageSetup.Show
    End If
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Dim bBullets As Boolean
    Select Case Button.Key
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Bold"
            rtfText.SelBold = Not rtfText.SelBold
            Button.Value = IIf(rtfText.SelBold, tbrPressed, tbrUnpressed)
        Case "Italic"
            rtfText.SelItalic = Not rtfText.SelItalic
            Button.Value = IIf(rtfText.SelItalic, tbrPressed, tbrUnpressed)
        Case "Underline"
            rtfText.SelUnderline = Not rtfText.SelUnderline
            Button.Value = IIf(rtfText.SelUnderline, tbrPressed, tbrUnpressed)
        Case "Align Left"
            rtfText.SelAlignment = rtfLeft
        Case "Center"
            rtfText.SelAlignment = rtfCenter
        Case "Align Right"
            rtfText.SelAlignment = rtfRight
        Case "Find"
            frmFind.Show , Me
        Case "Spelling"
            frmSpellIt.Show
        Case "Preview"
            mnuViewPrintPreview_Click
        Case "FPt"
            Button.Value = tbrPressed
            ' 1 means clicked and format
            ' is to be copied now
            fPaint = 1
            ' Function for Format Painter
            ' Defination in Utils module
            FPainter
        Case "help"
            AboutBox.Show
        Case "Full"
            mnuViewFullScreen_Click
        Case "Time"
            mnuInsertTimeDate_Click
        Case "Bullets"
            mnuFormatBullets_Click
    End Select
End Sub

Private Sub mnuFilePrint_Click()
     ' Check to see if there are any open documents, if not show error
   ' If ActiveForm Is Nothing Then MsgBox "No hay documentos abiertos !", , App.Title: Exit Sub
    PrintRTF rtfText, 720, 720, 720, 720 'Call PrintRTF sub
End Sub


Public Sub setMenu()
    Dim Found As Boolean
    
  '  If frmCount >= 1 Then
        Found = True
   ' Else
   '     Found = False
   ' End If
    
    On Error Resume Next
    
    mnuFileSave.Enabled = Found
    mnuFileSaveAs.Enabled = Found
    mnuFileClose.Enabled = Found
    mnuFilePrint.Enabled = Found
    mnuEditSelectAll.Enabled = Found
    mnuInsertTimeDate.Enabled = Found
    mnuEditFind.Enabled = Found
    'mnuEditFindNext.Enabled = Found
    mnuFormatFont.Enabled = Found
    mnuFormatChangeCase.Enabled = Found
    mnuFormatScript.Enabled = Found
    mnuToolsEncrypt.Enabled = Found
    mnuToolsDecrypt.Enabled = Found
    mnuWindowCascade.Enabled = Found
    mnuWindowTileHorizontal.Enabled = Found
    mnuWindowTileVertical.Enabled = Found
    mnuWindowArrangeIcons.Enabled = Found
    mnuFormatPageSetup.Enabled = Found
    tbToolBar.Buttons(19).Enabled = Found
    tbToolBar.Buttons(3).Enabled = Found
    tbToolBar.Buttons(5).Enabled = Found
    tbToolBar.Buttons(6).Enabled = Found
    tbToolBar.Buttons(8).Enabled = Found
    tbToolBar.Buttons(9).Enabled = Found
    tbToolBar.Buttons(10).Enabled = Found
    tbToolBar.Buttons(12).Enabled = Found
    tbToolBar.Buttons(13).Enabled = Found
    tbToolBar.Buttons(14).Enabled = Found
    tbToolBar.Buttons(15).Enabled = Found
    tbToolBar.Buttons(16).Enabled = Found
    tbToolBar.Buttons(17).Enabled = Found
    tbToolBar.Buttons(18).Enabled = Found
    mnuToolsSpellCheck.Enabled = Found
    tbToolBar.Buttons(20).Enabled = Found
    tbToolBar.Buttons(21).Enabled = Found
    tbToolBar.Buttons(23).Enabled = Found
    tbToolBar.Buttons(24).Enabled = Found
    tbToolBar.Buttons(25).Enabled = Found
    mnuFormatParagraph.Enabled = Found
    mnuFormatBullets.Enabled = Found
    mnuInsertPicture.Enabled = Found
    mnuViewFullScreen.Enabled = Found
    mnuToolsOptions.Enabled = Found
    mnuFilePrintConf.Enabled = Found
    mnuViewPrintPreview.Enabled = Found
    cmbFontName.Visible = Found
    cmbFontSize.Visible = Found
    
    If Me.rtfText.Text = "" Then
        mnuToolsEncrypt.Enabled = False
        mnuToolsDecrypt.Enabled = False
        mnuViewFullScreen.Enabled = False
        mnuViewPrintPreview.Enabled = False
        tbToolBar.Buttons(6).Enabled = False
    End If
    
    If Found = False Then
        mnuEditCopy.Enabled = False
        mnuEditCut.Enabled = False
        mnuEditPaste.Enabled = False
        tbToolBar.Buttons(8).Enabled = False
        tbToolBar.Buttons(9).Enabled = False
        tbToolBar.Buttons(10).Enabled = False
        mnuPopCopy.Enabled = False
        mnuPopCut.Enabled = False
        mnuPopPaste.Enabled = False
        sbStatusBar.Panels(1).Text = "Bienvenido"

    Else
        If rtfText.SelLength > 0 Then
            mnuEditCopy.Enabled = True
            mnuEditCut.Enabled = True
            mnuPopCopy.Enabled = True
            mnuPopCut.Enabled = True
            tbToolBar.Buttons(8).Enabled = True
            tbToolBar.Buttons(9).Enabled = True
            
        Else
            mnuEditCopy.Enabled = False
            mnuEditCut.Enabled = False
            tbToolBar.Buttons(8).Enabled = False
            tbToolBar.Buttons(9).Enabled = False
            mnuPopCopy.Enabled = False
            mnuPopCut.Enabled = False
        End If
        
        If Clipboard.GetFormat(vbCFText) Then
            mnuEditPaste.Enabled = True
            tbToolBar.Buttons(10).Enabled = True
            mnuPopPaste.Enabled = True
        Else
            mnuEditPaste.Enabled = False
            tbToolBar.Buttons(10).Enabled = False
            mnuPopPaste.Enabled = False
        End If
    End If

End Sub
