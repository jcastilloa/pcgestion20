VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   ClientHeight    =   6240
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10080
   Icon            =   "frmMainMod.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin miBwordOCX.epCmDlg epCmDlg1 
      Left            =   3105
      Top             =   4185
      _ExtentX        =   661
      _ExtentY        =   635
   End
   Begin VB.PictureBox dlgBPrint 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   10020
      TabIndex        =   9
      Top             =   1200
      Width           =   10080
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
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
         TabIndex        =   8
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
         Height          =   315
         Left            =   2280
         TabIndex        =   4
         Text            =   "cmbFontSize"
         Top             =   40
         Width           =   735
      End
      Begin VB.ComboBox cmbFontName 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Text            =   "cmbFontName"
         Top             =   45
         Width           =   1935
      End
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
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
      TabIndex        =   0
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
            Enabled         =   0   'False
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
            TextSave        =   "08/11/2003"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1508
            MinWidth        =   1499
            TextSave        =   "10:14"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1920
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
            Picture         =   "frmMainMod.frx":0442
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod.frx":0554
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod.frx":0666
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod.frx":0778
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod.frx":088A
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod.frx":099C
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod.frx":0AAE
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod.frx":0BC0
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod.frx":0CD2
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod.frx":0DE4
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod.frx":0EF6
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod.frx":1008
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod.frx":111A
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod.frx":122C
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod.frx":1340
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod.frx":1884
            Key             =   "Painter"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod.frx":1DC8
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod.frx":230C
            Key             =   "Spelling"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod.frx":2850
            Key             =   "Bullets"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod.frx":2E38
            Key             =   "FSreen"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainMod.frx":3504
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar2 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
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
         Left            =   435
         Picture         =   "frmMainMod.frx":3B40
         ScaleHeight     =   270
         ScaleWidth      =   11490
         TabIndex        =   6
         Top             =   120
         Width           =   11490
      End
   End
   Begin ComCtl3.CoolBar CoolBar3 
      Align           =   3  'Align Left
      Height          =   4185
      Left            =   0
      TabIndex        =   7
      Top             =   1680
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   7382
      BandCount       =   1
      Orientation     =   1
      _CBWidth        =   390
      _CBHeight       =   4185
      _Version        =   "6.7.8988"
      Child1          =   "Toolbar1"
      MinHeight1      =   330
      Width1          =   615
      NewRow1         =   0   'False
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

 Private Sub MDIForm_Load()
    fPaint = 3    ' Means Format Painter is not Active
    
    'These Lines Will get the last window position from registry
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    BMruInit
    dlgCommonDialog.FileName = GetSetting(App.Title, "Settings", "Default Dir", "C:\My Documents") & "\" & "Document"
    ' These lines will not allow the application to
    ' Load with a too small size window
    If GetSetting(App.Title, "Settings", "MainHeight", 6500) < 405 Then
        Me.Height = 6780
        Me.Left = 960
        Me.Top = 915
        Me.Width = 10335
    End If
    
    'To trigger the help menu option and set the help file path
    
    
    DocTemp = 0     'Currently no Document loaded
    frmCount = 0    'Variable to keep track of forms currently opened
    
    'New default Document created
    mnuFileNew_Click
    
    'Active Form Number
    ret = frm
    
    ' Fill the combo box with available fonts in the system
    For x = 1 To Screen.FontCount
        cmbFontName.AddItem Screen.Fonts(x)
    Next
    
    ' Fill the combo box with sizes 1 to 72 to choose from
    For x = 5 To 72: cmbFontSize.AddItem str$(x): Next
    
    ' The default font should be shown in the combo box
    For x = 0 To cmbFontName.ListCount - 1
        If ChildForms(ret).rtfText.SelFontName = cmbFontName.List(x) Then
            cmbFontName.ListIndex = x
            Exit For
        End If
    Next
    
    ' This one is to show the default size
    For x = 0 To cmbFontSize.ListCount - 1
        If Int(Val(ChildForms(ret).rtfText.SelFontSize)) = Val(cmbFontSize.List(x)) Then
            cmbFontSize.ListIndex = x
            Exit For
        End If
    Next
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
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

    lngPos = ActiveForm.rtfText.SelStart + ActiveForm.rtfText.SelLength
    ' Get position of the searched word
    lngResult = ActiveForm.rtfText.Find(frmFind.cboFind.Text, lngPos, , intOptions)
    If lngResult = -1 Then 'Text not found
        MsgBox "No se ha encontrado el texto", , App.Title
        frmFind.cmdFind.Caption = "Buscar" 'Set caption
        frmFind.cmdReplace.Enabled = False 'Disable Replace button
        frmFind.cmdReplaceAll.Enabled = False 'Disable ReplaceAll button
        mnuEditFindNext.Enabled = False 'Disable Find Next menu
    Else
        ActiveForm.rtfText.SetFocus 'Set focus
    End If
    Exit Sub
FindNextError:
    MsgBox Err.Description
End Sub
Private Sub mnuEditSelectAll_Click()
    ChildForms(frm).rtfText.SelStart = 0
    ChildForms(frm).rtfText.SelLength = Len(ChildForms(frm).rtfText.Text)
End Sub
Private Sub mnuFileClose_Click()
  Dim ret As Integer
  ' ret will store the form which is unloaded
  ' from the array of frmDocuments
  ret = frm
  Unload ChildForms(ret)
  Set ChildForms(ret) = Nothing
  ' Set the Unalail flag to false
  ' ie. the next document loaded will get this number
  UnAvail(ret) = False
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileMRU_Click(Index As Integer)
    OpenFile mnuFileMRU(Index).Caption
    storeMRU (mnuFileMRU(Index).Caption)
    storeMRUinReg
End Sub

Private Sub mnuFileNew_Click()
    Dim ret As Integer
    ' DocTemp will have the first number from
    ' the UnAvail array whose value is false
    DocTemp = FirstAvail
    
    ' If DocTemp is -1 then 30 documents have already
    ' been loaded which is the maximum number .
    If DocTemp <> -1 Then
        ' Creating a new document at runtime
        Set ChildForms(DocTemp) = New frmDocument
        ' Setting the Document title Document 1 ,2,3 etc.
        'ChildForms(DocTemp).Caption = "Document " & DocTemp
        ChildForms(DocTemp).Tag = DocTemp
        file(frm) = ""
        ' Count of forms opened incremented
        frmCount = frmCount + 1
        ' Function for context sensitive menu
        setMenu
    Else
        MsgBox "You are only allowed 30 documents opened at one time."
    End If
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
    OpenFile dlgCommonDialog.FileName
    Opened = True
    storeMRU (dlgCommonDialog.FileName)
    storeMRUinReg
    ' These array variables come handy in Save As Routine
    file(frm) = dlgCommonDialog.FileTitle
OpenCanceled:

End Sub
Private Sub OpenFile(str As String)
    ' Open *.rtf file types in RTF format
    ' Other files in Text Format
  
    If UCase(Right(str, 3)) = "RTF" Then
        mnuFileNew_Click
        On Error GoTo errFileOpen
        ChildForms(frm).rtfText.LoadFile str, rtfRTF
    Else
        mnuFileNew_Click
        On Error GoTo errFileOpen
        ChildForms(frm).rtfText.LoadFile str, rtfText
    End If
    
    ' Change the Document title to the name of the file
    ChildForms(frm).Caption = str
    Exit Sub
errFileOpen:
    MsgBox Err.Description
End Sub

Private Sub mnuFilePageSetup_Click()
dlgBPrint.ShowPageSetup 'Show Page Setup dialog
End Sub

Private Sub mnuFilePrintSetup_Click()
dlgBPrint.ShowPrinter 'Show Printer dialog
End Sub

Public Sub mnuFileSave_Click()
    ' Setting flags for readOnly and Overwrite prompt
    dlgCommonDialog.flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    
    On Error Resume Next
    
    ' If file name is Blank or its a New Document
    If dlgCommonDialog.FileName = "" Or Left(ChildForms(frm).Caption, 8) = "Document" Then
        
        mnuFileSaveAs_Click
    
    ' if document caption and file name are same
    ' which means save dialog box need not appear
    ' Save the file as it is with the same name ie. Update the file
    ElseIf ChildForms(frm).Caption = dlgCommonDialog.FileName Then
        ' If RTF save it in RTF format else in text format
        If UCase(Right(dlgCommonDialog.FileName, 3)) = "RTF" Then
            ChildForms(frm).rtfText.SaveFile dlgCommonDialog.FileName, rtfRTF
        Else
            ChildForms(frm).rtfText.SaveFile dlgCommonDialog.FileName, rtfText
        End If
        
        ' File already saved so no need to save again
        ' So that if we close the file now it will
        ' not prompt the user to save again
        If NeedSaved(frm) = True Then
            NeedSaved(frm) = False
            ChildForms(frm).Caption = Right$(ChildForms(frm).Caption, Len(ChildForms(frm).Caption))
        End If
        
        ' If somehow user doesnt save and exits
        ' Formcount is decremented
        ' and menus refreshed
        If DocumentClosed = True Then
            frmCount = frmCount - 1
            setMenu
        End If
End If
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
        ChildForms(frm).rtfText.SaveFile dlgCommonDialog.FileName, rtfRTF
        storeMRU (dlgCommonDialog.FileName)
        storeMRUinReg
    Else
        ChildForms(frm).rtfText.SaveFile dlgCommonDialog.FileName, rtfText
        storeMRU (dlgCommonDialog.FileName)
        storeMRUinReg
    End If
    
    If NeedSaved(frm) = True Then
        NeedSaved(frm) = False
        ChildForms(frm).Caption = Right$(ChildForms(frm).Caption, Len(ChildForms(frm).Caption))
    End If
  
    ChildForms(frm).Caption = dlgCommonDialog.FileTitle
    file(frm) = dlgCommonDialog.FileTitle
    
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
    Menu_FormatBullet ChildForms(frm)
End Sub
Private Sub mnuFormatCSLower_Click()
    ' LowerCase Conversion
    Clipboard.SetText ChildForms(frm).rtfText.SelText
    ChildForms(frm).rtfText.SelText = LCase(Clipboard.GetText)
    Clipboard.Clear
End Sub

Private Sub mnuFormatCSSC_Click()
    ' Sentence Case Conversion
    Clipboard.SetText ChildForms(frm).rtfText.SelText
    ChildForms(frm).rtfText.SelText = StrConv(Clipboard.GetText, vbProperCase)
    Clipboard.Clear
End Sub

Private Sub mnuFormatCSUpper_Click()
    ' UpperCase Conversion
    Clipboard.SetText ChildForms(frm).rtfText.SelText
    ChildForms(frm).rtfText.SelText = UCase(Clipboard.GetText)
    Clipboard.Clear
End Sub

Private Sub mnuFormatFont_Click()
    ' cdlCFBoth will Display both Screen and Printer Fonts
    ' cdlCFEffects will display all the attributes
    ' including Underline , StrikeThru and Color
    dlgCommonDialog.flags = cdlCFBoth Or cdlCFEffects
    
    
    With ActiveForm.rtfText
        ' Update the Font Dialog Box with
        ' the Current Font Format of the selected text
         On Error GoTo fnterr
        dlgCommonDialog.FontName = .SelFontName
        dlgCommonDialog.FontSize = .SelFontSize
        dlgCommonDialog.FontBold = .SelBold
        dlgCommonDialog.FontItalic = .SelItalic
        dlgCommonDialog.FontUnderline = .SelUnderline
        dlgCommonDialog.Color = .SelColor
        
       
        dlgCommonDialog.ShowFont
        
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
    ChildForms(frm).rtfText.SelCharOffset = 0
End Sub

Private Sub mnuFormatPageSetup_Click()
    frmPageSetup.Show
End Sub

Private Sub mnuFormatParagraph_Click()
    frmParagraph1.Show
End Sub

Private Sub mnuFormatSubScript_Click()
    ChildForms(frm).rtfText.SelCharOffset = -55
End Sub

Private Sub mnuFormatSuperScript_Click()
    ChildForms(frm).rtfText.SelCharOffset = 55
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
    ChildForms(frm).rtfText.SelFontName = cmbFontName.List(cmbFontName.ListIndex)
End Sub

Private Sub cmbFontSize_Click()
    ' Change the Font Size.
    ChildForms(frm).rtfText.SelFontSize = cmbFontSize.List(cmbFontSize.ListIndex)
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
    prgBar.Max = Len(ChildForms(frm).rtfText.Text)
    
    ' Encryption will be Character by Character
    For x = 1 To Len(ChildForms(frm).rtfText.Text) Step 8
        TB = Asc(Mid$(EncKey, EncKeyPos, 1))
        EncKeyPos = EncKeyPos + 1
        
        If EncKeyPos > EncLen Then EncKeyPos = 1
        
        ' Binary code of a character is
        ' 8 characters long
        tempChar = Mid$(ChildForms(frm).rtfText.Text, x, 8)
    
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
    ChildForms(frm).rtfText.Text = EncStr
    
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
    prgBar.Max = Len(ChildForms(frm).rtfText.Text)
  
    For x = 1 To Len(ChildForms(frm).rtfText.Text)
        TB = Asc(Mid$(EncKey, EncKeyPos, 1))
        EncKeyPos = EncKeyPos + 1
        
        If EncKeyPos > EncLen Then
            EncKeyPos = 1
        End If
        
        TA = Asc(Mid$(ChildForms(frm).rtfText.Text, x, 1))
        TC = TB Xor TA
        ' Get Binary is Function to convert
        ' Binary to text again
        tempChar = GetBinary(TC)
        EncStr = EncStr & tempChar
        prgBar.Value = x
    Next
    
    ChildForms(frm).rtfText.Text = EncStr
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
    If Len(ChildForms(frm).rtfText.Text) = 0 Then
        MsgBox "Cannot view Full Screen " + vbNewLine + "There is no text in the Document."
        Exit Sub
    End If
    mnuEditSelectAll_Click
    SendMessage ChildForms(frm).rtfText.hwnd, WM_COPY, 0, 0&
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
    Me.Arrange vbArrangeIcons
End Sub
Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub
Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub
Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub
Private Sub mnuWindowNewWindow_Click()
    mnuFileNew_Click
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
    ActiveForm.rtfText.SelRTF = Clipboard.GetText
End Sub
Private Sub mnuEditCopy_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtfText.SelText
End Sub
Private Sub mnuEditCut_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtfText.SelText
    ActiveForm.rtfText.SelText = vbNullString
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
        Case "New"
            mnuFileNew_Click
        Case "Open"
            mnuFileOpen_Click
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
            ActiveForm.rtfText.SelBold = Not ActiveForm.rtfText.SelBold
            Button.Value = IIf(ActiveForm.rtfText.SelBold, tbrPressed, tbrUnpressed)
        Case "Italic"
            ActiveForm.rtfText.SelItalic = Not ActiveForm.rtfText.SelItalic
            Button.Value = IIf(ActiveForm.rtfText.SelItalic, tbrPressed, tbrUnpressed)
        Case "Underline"
            ActiveForm.rtfText.SelUnderline = Not ActiveForm.rtfText.SelUnderline
            Button.Value = IIf(ActiveForm.rtfText.SelUnderline, tbrPressed, tbrUnpressed)
        Case "Align Left"
            ActiveForm.rtfText.SelAlignment = rtfLeft
        Case "Center"
            ActiveForm.rtfText.SelAlignment = rtfCenter
        Case "Align Right"
            ActiveForm.rtfText.SelAlignment = rtfRight
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
    If ActiveForm Is Nothing Then MsgBox "No hay documentos abiertos !", , App.Title: Exit Sub
    PrintRTF ActiveForm.rtfText, 720, 720, 720, 720 'Call PrintRTF sub
End Sub


Public Sub setMenu()
    Dim Found As Boolean
    
    If frmCount >= 1 Then
        Found = True
    Else
        Found = False
    End If
    
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
    
    If ChildForms(frm).rtfText.Text = "" Then
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
        If ActiveForm.rtfText.SelLength > 0 Then
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
