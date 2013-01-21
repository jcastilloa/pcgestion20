VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6514F5A0-641C-11D2-9FD0-0020AF131A57}#2.1#0"; "fpFlp20.ocx"
Begin VB.Form frmCabPagPro 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pagos a ..."
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10485
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin LpADOLib.fpComboADO cbMESES 
      Height          =   390
      Left            =   3255
      TabIndex        =   3
      Top             =   975
      Width           =   855
      _Version        =   131073
      _ExtentX        =   1508
      _ExtentY        =   688
      _StockProps     =   68
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   0
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
      WrapList        =   0   'False
      WrapWidth       =   0
      SelMax          =   -1
      AutoSearch      =   1
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
      ThreeDOutsideHighlightColor=   16777215
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   1
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ScrollHScale    =   2
      ScrollHInc      =   0
      ColsFrozen      =   0
      ScrollBarV      =   1
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   0
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   2
      SearchIgnoreCase=   -1  'True
      ColDesigner     =   "frmCabPagPro.frx":0000
      ScrollBarH      =   1
      DataFieldList   =   ""
      ColumnEdit      =   -1
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   -1
      ComboGap        =   7
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   -2147483637
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   -2147483637
      ExtendRow       =   0
   End
   Begin PCGestion.chameleonButton cmdPrevious 
      Height          =   630
      Left            =   1080
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6015
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   1111
      BTYPE           =   9
      TX              =   "F6"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCabPagPro.frx":027C
      PICN            =   "frmCabPagPro.frx":0298
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblstatus 
      Height          =   375
      Left            =   15
      Top             =   5595
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   661
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   15640462
      Colour2         =   7177785
      CaptionAlignment=   1
   End
   Begin PCGestion.miText ioIMPORTE 
      Height          =   525
      Left            =   990
      TabIndex        =   2
      Top             =   915
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   926
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      dspFormat       =   ""
      Enabled         =   -1  'True
      EsPassword      =   -1  'True
   End
   Begin PCGestion.chameleonButton cmdFirst 
      Height          =   630
      Left            =   45
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6015
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   1111
      BTYPE           =   9
      TX              =   "F5"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCabPagPro.frx":0F6A
      PICN            =   "frmCabPagPro.frx":0F86
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbLista 
      Height          =   630
      Left            =   4800
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1111
      BTYPE           =   9
      TX              =   "Lista F4"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCabPagPro.frx":1CBC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cmdNext 
      Height          =   630
      Left            =   8340
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6015
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   1111
      BTYPE           =   9
      TX              =   "F7"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCabPagPro.frx":1CD8
      PICN            =   "frmCabPagPro.frx":1CF4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cmdLast 
      Height          =   630
      Left            =   9405
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6015
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   1111
      BTYPE           =   9
      TX              =   "F8"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCabPagPro.frx":29C6
      PICN            =   "frmCabPagPro.frx":29E2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbAgregar 
      Height          =   795
      Left            =   30
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6675
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "&Agregar F1"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCabPagPro.frx":3718
      PICN            =   "frmCabPagPro.frx":3734
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbActualizar 
      Height          =   795
      Left            =   1125
      TabIndex        =   8
      Top             =   6675
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "&Actualizar F2"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCabPagPro.frx":440E
      PICN            =   "frmCabPagPro.frx":442A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbEdicion 
      Height          =   795
      Left            =   2355
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6675
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "&Edicion F3"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCabPagPro.frx":4D04
      PICN            =   "frmCabPagPro.frx":4D20
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbCancelar 
      Height          =   795
      Left            =   7320
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6675
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "&Cancelar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCabPagPro.frx":557E
      PICN            =   "frmCabPagPro.frx":559A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbEliminar 
      Height          =   795
      Left            =   8280
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6675
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "E&liminar F9"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCabPagPro.frx":5E74
      PICN            =   "frmCabPagPro.frx":5E90
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbCerrar 
      Height          =   795
      Left            =   9405
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6675
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "Cerrar ESC"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCabPagPro.frx":6A62
      PICN            =   "frmCabPagPro.frx":6A7E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.miCombo cbCODPROV 
      Height          =   495
      Left            =   5940
      TabIndex        =   1
      Top             =   465
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PCGestion.miCombo ioTIPOPAGO 
      Height          =   495
      Left            =   975
      TabIndex        =   0
      Top             =   450
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PCGestion.miText ioCUOTA 
      Height          =   525
      Left            =   5955
      TabIndex        =   4
      Top             =   915
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   926
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      dspFormat       =   ""
      Enabled         =   -1  'True
      EsPassword      =   -1  'True
   End
   Begin PCGestion.miText ioPEDIDO 
      Height          =   525
      Left            =   7740
      TabIndex        =   5
      Top             =   915
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   926
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      dspFormat       =   ""
      Enabled         =   -1  'True
      EsPassword      =   -1  'True
   End
   Begin PCGestion.miText ioFACTURA 
      Height          =   525
      Left            =   9450
      TabIndex        =   6
      Top             =   915
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   926
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      dspFormat       =   ""
      Enabled         =   -1  'True
      EsPassword      =   -1  'True
   End
   Begin PCGestion.miCombo cbCODCAJA 
      Height          =   495
      Left            =   3915
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6825
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PCGestion.miText ioDESCRIPCION 
      Height          =   525
      Left            =   990
      TabIndex        =   7
      Top             =   1425
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   926
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      dspFormat       =   ""
      Enabled         =   -1  'True
      EsPassword      =   -1  'True
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   3540
      Left            =   3896
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   1965
      Width           =   5700
      _cx             =   10054
      _cy             =   6244
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCabPagPro.frx":7758
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin PCGestion.chameleonButton cmNuevoPago 
      Height          =   525
      Left            =   1384
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   3405
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   926
      BTYPE           =   9
      TX              =   "-N- Nuevo Pago"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCabPagPro.frx":77FD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cmBorrarPago 
      Height          =   525
      Left            =   1384
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   3975
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   926
      BTYPE           =   9
      TX              =   "-B- Borrar Pago"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCabPagPro.frx":7819
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "USUARIO"
      Height          =   330
      Left            =   6300
      TabIndex        =   43
      Top             =   1530
      Width           =   960
   End
   Begin VB.Label lblUsuario 
      Alignment       =   2  'Center
      BackColor       =   &H00EEA78E&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7335
      TabIndex        =   42
      Top             =   1500
      Width           =   3105
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DESC."
      Height          =   345
      Left            =   225
      TabIndex        =   38
      Top             =   1515
      Width           =   705
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CAJA"
      Height          =   330
      Left            =   3345
      TabIndex        =   37
      Top             =   6915
      Width           =   555
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "FACT."
      Height          =   345
      Left            =   8775
      TabIndex        =   36
      Top             =   990
      Width           =   660
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PEDIDO"
      Height          =   345
      Left            =   6915
      TabIndex        =   35
      Top             =   990
      Width           =   750
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MESES"
      Height          =   330
      Left            =   4185
      TabIndex        =   34
      Top             =   990
      Width           =   705
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CADA"
      Height          =   285
      Left            =   2460
      TabIndex        =   33
      Top             =   990
      Width           =   645
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PENDIENTE"
      Height          =   315
      Left            =   810
      TabIndex        =   32
      Top             =   2595
      Width           =   1140
   End
   Begin VB.Label lblPendiente 
      Alignment       =   2  'Center
      BackColor       =   &H00EEA78E&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   2055
      TabIndex        =   31
      Top             =   2520
      Width           =   1560
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CUOTA"
      Height          =   345
      Left            =   5190
      TabIndex        =   30
      Top             =   990
      Width           =   705
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO"
      Height          =   330
      Left            =   360
      TabIndex        =   29
      Top             =   540
      Width           =   555
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PROVEEDOR"
      Height          =   330
      Left            =   4665
      TabIndex        =   28
      Top             =   540
      Width           =   1245
   End
   Begin MSForms.CheckBox ioMBAJA 
      Height          =   435
      Left            =   2220
      TabIndex        =   27
      Top             =   6030
      Width           =   840
      VariousPropertyBits=   746596375
      BackColor       =   12632256
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "1482;767"
      Value           =   "0"
      Caption         =   "Baja"
      FontName        =   "Trebuchet MS"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label ioFBAJA 
      Alignment       =   2  'Center
      BackColor       =   &H00EEA78E&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   6900
      TabIndex        =   16
      Top             =   6045
      Width           =   1245
   End
   Begin VB.Label ioFMODI 
      Alignment       =   2  'Center
      BackColor       =   &H00EEA78E&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   8025
      TabIndex        =   15
      Top             =   60
      Width           =   2355
   End
   Begin VB.Label ioCODIGO 
      Alignment       =   2  'Center
      BackColor       =   &H00AC998C&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1020
      TabIndex        =   14
      Top             =   60
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      Height          =   330
      Left            =   105
      TabIndex        =   13
      Top             =   75
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ultima modificación"
      Height          =   315
      Left            =   5790
      TabIndex        =   12
      Top             =   75
      Width           =   2160
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "F. Baja"
      Height          =   330
      Left            =   6030
      TabIndex        =   11
      Top             =   6075
      Width           =   780
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "IMPORTE"
      Height          =   345
      Left            =   -60
      TabIndex        =   10
      Top             =   1005
      Width           =   990
   End
End
Attribute VB_Name = "frmCabPagPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Plantilla
' DateTime  : 31/10/2003 10:08
' Author    : Administrador
' Purpose   : Plantilla de código para los formularios de maestros.
'---------------------------------------------------------------------------------------
'·································································································································
' Convenio:
'·································································································································
'
' Para los campos de texto:  usar miText
' Para los combos:              usar miCombo.
'
'
' - Instrucciones:
'
' Enlazar los controles a los campos en Form_Load()   (ver ). Y especificar la tabla
' y orden (y otras cosas que se pudieran necesitar) mediante los parametros del
' oSQL (objecto SmartSQL):
'
'  oSQL.AddTable "SECCIONES"
'  oSQL.AddOrderClause "CODIGO"

'---------------------------------------------------------------------------------------
' - Cambiar en:
'
' Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
' set plantilla = nothing
' por el nombre del formulario. Ej:   set frmMntPrue = nothing

'---------------------------------------------------------------------------------------
' - Si se utiliza algun campo simulando que sea incremental (que se incremente en cada
' registro) cambia en Private Sub cbAgregar_Click()
'
' tmpcodigo = devuelve_campo("select max(codigo) + 1 from secciones")
'
' y poner el SQL correcto para que nos devuelva el proximo codigo para nuestro campo
'
'
'
'---------------------------------------------------------------------------------------
' - Colocar 2 tipos de validaciones para los datos.
'
'  Una validación a nivel de campo. Por ejemplo, comprobar al salir del campo
'  que la información es correcta, usando el evento validate. (si es > X, <> "", etc)
'
'- Otra validación es en:
'
'Private Sub rc_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'
' El evento cuando cambia un registro del recordset, pones el codigo por ejemplo en

'        Case adRsnUpdate
'
' Mira el formulario frmMntCli para ver un ejemplo de esto.
'
'---------------------------------------------------------------------------------------
' - Formularios de Lista. Para llamar al formulario de lista estandar FrmFlexSimple, ver el
' codigo de cbLista_Click. Cambiar los colformats y otras cosas que puedan ser
' necesarias, para adecuar a cada formulario
'
'
'---------------------------------------------------------------------------------------
'Otras notas:
'
' - comprobar el orden correcto de los tabindex para permitir recorrer miText y miCombo
' del formulario con el teclado (soportan el avance con ENTER). Desde el primero hasta
' el ultimo.
'
'- cambiar en cbAgregar_clik y cbEditar_click
'
' ioDescripcion.setfocus
'
' por el nombre del control que tengamos que activar en primer lugar.
'
' cambiar en Private Sub ioCODIGO_Change(), y poner el numero de 000
' correcto en cada caso


Option Explicit
Dim WithEvents rc As ADODB.Recordset
Attribute rc.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim prime As Boolean

Dim oSQL As New clsSmartSQL

Dim TmpUsr As Long


'si entra por Pago_A_Proveedor = true, fijar el combo tipo pago con el valor 1
'pagos a proveedores, y no mostrar el dialogo de seleccion de usuario
'(asignar a usuario actual)
Public Pago_A_Proveedor As Boolean



Private Sub cbMESES_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

Case vbKeyUp
    SendKeys "{+TAB}"
    KeyAscii = 0

Case 13
    SendKeys "{TAB}"
    KeyAscii = 0
    
Case vbKeyDown
    SendKeys "{TAB}"
    KeyAscii = 0
    
End Select

End Sub

Private Sub cbMESES_Validate(Cancel As Boolean)

   On Error GoTo cbMESES_Validate_Error

   ' If cbMESES.Text <> "" And IsNumeric(cbMESES.Text) Then
    
       ' If ioIMPORTE.Text <> "" And IsNumeric(ioIMPORTE.Text) Then

        '    ioCUOTA.Text = ioIMPORTE.Text / cbMESES.Text

       ' Else
        
       '     lblstatus.Caption = "Importe no puede estar en blanco para calcular una cuota"
       '     Exit Sub

       ' End If
    
  '  End If

   On Error GoTo 0
   Exit Sub

cbMESES_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbMESES_Validate de Formulario frmCabPagPro"

End Sub

Private Sub cmBorrarPago_Click()
Dim linea As Long

   On Error GoTo cmBorrarPago_Click_Error

With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
End With

If fg.TextMatrix(fg.Row, 1) = "" Then Exit Sub
If Not IsNumeric(fg.TextMatrix(fg.Row, 1)) Then Exit Sub

If MsgBox("¿Desea borrar el pago seleccionado?. un importe de: " & fg.TextMatrix(fg.Row, 2) & ". Y fecha: " & fg.TextMatrix(fg.Row, 3), vbQuestion + vbYesNo) = vbNo Then Exit Sub

linea = fg.TextMatrix(fg.Row, 1)
   
locCnn.Execute "UPDATE DETPAGOS SET MBAJA = 1 WHERE CODIGO = " & rc.fields("CODIGO") & " AND CODCAJA = " & rc.fields("CODCAJA") & " AND LINEA = " & linea

Call carga_grid
DoEvents

rc.fields("PAGADO") = CDbl(fg.TextMatrix(1, 2))
rc.UpdateBatch adAffectAll

DoEvents
rc.Move 0

   On Error GoTo 0
   Exit Sub

cmBorrarPago_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmBorrarPago_Click de Formulario frmCabPagPro"

End Sub

Private Sub cmNuevoPago_Click()

   On Error GoTo cmNuevoPago_Click_Error

    If (rc.BOF And rc.EOF) Then Exit Sub
    
    With frmDPagosPr
    .ID_Agregar = True
    .ID_Caja = rc.fields("CODCAJA")
    .ID_Pago = rc.fields("CODIGO")
    .Total_Deuda = rc.fields("IMPORTE")
    .ID_Cuota = rc.fields("CUOTA")
    .Show 1
    
    'actualizar el total pagado
    rc.fields("PAGADO") = .Total_Pagado
    
    If .Deuda_Pagada Then
    
        'poner como estado = 2 PAGADO
        rc.fields("ESTADO") = 2
    
    End If
    
        
    rc.UpdateBatch adAffectAll
        
    Set frmDPagosPr = Nothing
    
    Call carga_grid
    DoEvents
    rc.Move 0
    
    End With
    

   On Error GoTo 0
   Exit Sub

cmNuevoPago_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmNuevoPago_Click de Formulario frmCabPagPro"

End Sub

Private Sub fg_DblClick()

   On Error GoTo fg_dblClick_Error

If fg.TextMatrix(fg.Row, 1) = "" Then Exit Sub
If Not IsNumeric(fg.TextMatrix(fg.Row, 1)) Then Exit Sub

With frmDPagosPr
    .ID_A_Id = True
    .ID_Caja = rc.fields("CODCAJA")
    .ID_Pago = rc.fields("CODIGO")
    .ID_ID_Seleccionado = fg.TextMatrix(fg.Row, 1)
    .Show 1
    
    'actualizar el total pagado
    rc.fields("PAGADO") = .Total_Pagado
    rc.UpdateBatch adAffectAll
        
    Set frmDPagosPr = Nothing
    
    Call carga_grid
    
    DoEvents
    rc.Move 0
    
End With

   On Error GoTo 0
   Exit Sub

fg_dblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento fg_dblClick de Formulario frmCabPagPro"

End Sub


Private Sub ioCODIGO_Change()
ioCODIGO.Caption = Format(ioCODIGO.Caption, "00000")
End Sub

Private Sub Form_Activate()

If Not prime Then


  If rc.RecordCount = 0 Then
         

 '       If MsgBox("No se encuentran Pagos. ¿Crear?", vbYesNo + vbQuestion, "Pagos") = vbNo Then
 '       Unload Me
 '       Else
 '       Call cbAgregar_Click
 '       End If
        
  Else
        Call cmdFirst_Click
        Call cbCancelar_Click
        
  End If

prime = True
End If
    
End Sub

Private Sub Form_Load()
  
   Move (Screen.Width - Width) \ 2, Separacion_MDIForm
   
   With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
   End With
   
  Set rc = New Recordset
  oSQL.AddTable "PAGOS"
  oSQL.AddOrderClause "CODIGO"
  
  If Pago_A_Proveedor Then
    oSQL.AddSimpleWhereClause "TIPOPAGO", 1
  Else
    oSQL.AddSimpleWhereClause "TIPOPAGO", 1, , CLAUSE_DOESNOTEQUAL
  End If
    
  rc.Open oSQL.SQL, locCnn, adOpenStatic, adLockOptimistic
    

''''''''
'////////////////////////////////////////////////////////
' Enlazar controles:
'////////////////////////////////////////////////////////
  With ioCODIGO
  Set .DataSource = rc
        .DataField = "CODIGO"
  End With
  
  With ioIMPORTE
        .Alineacion = 1
        .PermitirBlanco = True
        .LongMaxima = 10
        .dspFormat = "Currency"
  End With
  
    With ioCUOTA
        .Alineacion = 1
        .PermitirBlanco = True
        .dspFormat = "Currency"
        .LongMaxima = 10
       ' .Locked = True
  End With
  
  With ioPEDIDO
        .Alineacion = 1
  Set .DataSource = rc
        .DataField = "NUMPED"
        .PermitirBlanco = True
        .LongMaxima = 12
  End With
  
 With ioFACTURA
         .Alineacion = 1
  Set .DataSource = rc
        .DataField = "FACTURA"
        .PermitirBlanco = True
        .LongMaxima = 12
  End With
  
  With ioDESCRIPCION
    Set .DataSource = rc
        .DataField = "Descripcion"
        .PermitirBlanco = True
        .LongMaxima = 30
  End With
   
    
  'solo mostrar caja cuando entran los dependientes
  'tipo pago <> 1
  If (Pago_A_Proveedor = False) Then
  
  With cbCODCAJA
     .LenCodigo = 5
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM CAJAS WHERE MBAJA = 0 ORDER BY CODIGO"
    .CodigoWidth = 800
    .carga
    DoEvents
    .Text = CajaActual
  End With
  
  Else
  
    Label13.Visible = False
    cbCODCAJA.Visible = False
  
  End If
  
  
    With ioTIPOPAGO
    .ConexionString = locCnn
    .LenCodigo = 5
    
    If Pago_A_Proveedor Then
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM MAPAG WHERE (MBAJA = 0) AND (CODIGO=1) ORDER BY CODIGO"
    Else
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM MAPAG WHERE (MBAJA = 0) AND (CODIGO>1) ORDER BY CODIGO"
    End If
    
    .DataField = "TIPOPAGO"
    .CodigoWidth = 800
    .carga
    DoEvents
    Set .DataSource = rc
    
    If Pago_A_Proveedor Then
        .Text = 1   'fijar en pagos a proveedores
        .Enabled = False
    Else
        .Text = ""
    End If
    
  End With
    
    
  If Not Pago_A_Proveedor Then
  
    Label9.Visible = False
    cbMESES.Visible = False
    Label11.Visible = False
    Label8.Visible = False
    ioCUOTA.Visible = False
    Label4.Visible = False
    ioPEDIDO.Visible = False
    Label12.Visible = False
    ioFACTURA.Visible = False
    Label10.Visible = False
    lblPendiente.Visible = False
    Label6.Visible = False
    cbCODPROV.Visible = False
    
  Else
  
    With cbMESES
        .AddItem "0"
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        .AddItem "6"
        .AddItem "7"
        .AddItem "8"
        .AddItem "9"
        .AddItem "10"
        .AddItem "11"
        .AddItem "12"
    Set .DataSource = rc
        .DataField = "MESES"
   End With
   
   With cbCODPROV
       
    .ConexionString = locCnn
    .LenCodigo = 5
    .DataField = "CODPROV"
    .SQLString = "SELECT CODIGO, NOMBRE FROM MAPROV WHERE MBAJA = 0 ORDER BY CODIGO"
    .CodigoWidth = 800
    .carga
    DoEvents
    Set .DataSource = rc
    
  End With
  
  
  End If
    
    

  
  With ioFBAJA
  Set .DataSource = rc
        .DataField = "FBAJA"
  End With
  
  With ioFMODI
  Set .DataSource = rc
        .DataField = "FMODI"
  End With
  
  With ioMBAJA
  Set .DataSource = rc
        .DataField = "MBAJA"
  End With
              
  mbDataChanged = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      
      If MsgBox("¿Desea cerrar la ventana actual?", vbQuestion + vbYesNo, "Cerrar") = vbYes Then
        cbcerrar_Click
      End If
      
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
      
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
      
      Case vbKeyF1
            Call cbAgregar_Click
        
      Case vbKeyF2
            Call cbactualizar_Click
        
      Case vbKeyF3
            Call cbedicion_Click
        
      Case vbKeyF4
            Call cbLista_click
      
      Case vbKeyF5
            Call cmdFirst_Click
    
       Case vbKeyF6
            Call cmdPrevious_Click
      
       Case vbKeyF7
            Call cmdNext_Click
    
       Case vbKeyF8
        Call cmdLast_Click
            
      Case vbKeyN
      
        Call cmNuevoPago_Click
           
      Case vbKeyB
      
        Call cmBorrarPago_Click
       
  End Select
  KeyCode = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error Resume Next

If rc.EditMode <> adEditNone Then rc.CancelUpdate

rc.Close
Set rc = Nothing

  ' With locCnn
  '  If .State <> 0 Then .Close
  ' End With

Set oSQL = Nothing

'If Me.MDIChild = True Then frmMenuTactil.Show
Set frmCabPagPro = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub



Private Sub cbLista_click()

If rc.EditMode = adEditNone Then

With frmFlexPagPro

    .Pago_A_Proveedor = Pago_A_Proveedor
    .Caption = "Ver Pagos ..."
            
    .Show 1

End With

Else

    MsgBox "Debe guardar o cancelar cambios antes de seleccionar un nuevo registro", vbInformation, "Atención"

End If

End Sub

Private Sub rc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  If rc.AbsolutePosition > 0 Then
    lblstatus.Caption = "Registro: " & CStr(rc.AbsolutePosition)
    
    ioIMPORTE.Text = rc.fields("IMPORTE")
    ioCUOTA.Text = rc.fields("CUOTA")
    
    If rc.fields("CODCAJA") > 0 Then
        cbCODCAJA.Text = rc.fields("CODCAJA")
        Call carga_grid
    End If
    
    lblPendiente.Caption = Format(rc.fields("IMPORTE") - rc.fields("PAGADO"), "Currency")
    
    
    If Not IsNull(rc.fields("CODPER")) And rc.fields("CODPER") > 0 Then
             
         lblUsuario.Caption = Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & rc.fields("CODPER"), locCnn))
         DoEvents
    Else
    
         lblUsuario.Caption = ""
    
    End If
      
    Select Case (rc.fields("IMPORTE") - rc.fields("PAGADO"))
    
    'si es 0 cambiar el color del borde
    Case 0
        lblPendiente.BorderStyle = 1
      
    
    Case Else
    
        lblPendiente.BorderStyle = 0
    
    End Select
      
  Else
  
    ioIMPORTE.Text = ""
    ioCUOTA.Text = ""
    lblPendiente.Caption = ""
    fg.Clear
  
  End If
  
End Sub

Private Sub rc_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Aquí se coloca el código de validación
  'Se llama a este evento cuando ocurre la siguiente acción
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cbAgregar_Click()
  Dim tmpcodigo As Variant
  
  On Error GoTo AddErr
  With rc
       
         TmpUsr = 0
                
         'si no entra por pagos a proveedores
         If Pago_A_Proveedor = False Then
         
         Do
         
            With frmSelDep
                .Show 1
                TmpUsr = .ID_Dependiente
                Unload frmSelDep
            End With
        
         Loop Until TmpUsr <> 0
  
         Set frmSelDep = Nothing
          
         'si es a proveedores ...
         Else
            TmpUsr = UsuarioActual
         End If
         
         lblUsuario.Caption = Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & TmpUsr, locCnn))
         DoEvents
             
    fg.Clear
    fg.Rows = 1
    
    .AddNew
    
    If Pago_A_Proveedor = False Then
    
    If cbCODCAJA.Text = "" Then
        lblstatus.Caption = "Caja no puede estar en blanco para añadir nuevos pagos"
        cbCODCAJA.SetFocus
        Exit Sub
    End If
    
    End If
    
    If Pago_A_Proveedor Then
        ioTIPOPAGO.Text = 1   'fijar en pagos a proveedores
    Else
        ioTIPOPAGO.Text = ""
    End If
    
    tmpcodigo = devuelve_campo("select max(codigo) + 1 from PAGOS where CODCAJA = " & cbCODCAJA.Text)
    
    'Si devuelve @ esque ha habido un error
    If tmpcodigo = "@" Then tmpcodigo = 1
    
    .fields("CODIGO") = tmpcodigo
    .fields("CODCAJA") = cbCODCAJA.Text
    .fields("CODPER") = TmpUsr
    
    cbMESES.Text = "0"
    
    'End If

    Set tmpcodigo = Nothing
    
    lblstatus.Caption = "Agregar registro"
    mbAddNewFlag = True
    SetButtons False
    cbActualizar.Visible = True
  
  If Pago_A_Proveedor Then
    cbCODPROV.SetFocus
  Else
    ioTIPOPAGO.SetFocus
  End If
  
  End With

    
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cbEliminar_Click()
   
  On Error GoTo cbEliminar_Click_Error
  
  
  If MsgBox("¿Desea anular la DEUDA actual? Importe: " & Format(rc.fields("IMPORTE"), "Currency") & ". (se anulará la deuda y todos los pagos realizados para la deuda actual)", vbQuestion + vbYesNo, titulo) = vbYes Then
         
    With rc

        .fields("mbaja") = True
        .fields("FBAJA") = Now
    
        'poner los apuntes de pagos en baja
        locCnn.Execute "UPDATE DETPAGOS SET MBAJA = 1 WHERE CODIGO = " & .fields("CODIGO") & " AND CODCAJA = " & .fields("CODCAJA")
        
        DoEvents
    
    
        If .EOF Then .MoveLast
    
     End With
 
     Call cbactualizar_Click
  
     MsgBox "La deuda actual se ha anulado correctamente", vbInformation, titulo

   End If
     
  Exit Sub

   On Error GoTo 0
   Exit Sub

cbEliminar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbEliminar_Click de Formulario frmCabPagPro"

End Sub



Private Sub cbedicion_Click()
  On Error GoTo EditErr

  lblstatus.Caption = "Modificar registro"
  mbEditFlag = True
  SetButtons False
  cbActualizar.Visible = True
  
  If Pago_A_Proveedor Then
    cbCODPROV.SetFocus
  Else
    ioTIPOPAGO.SetFocus
  End If
  
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cbCancelar_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  rc.CancelUpdate
  If mvBookMark > 0 Then
    rc.Bookmark = mvBookMark
  Else
    rc.MoveFirst
  End If
  mbDataChanged = False

End Sub

Private Sub cbactualizar_Click()
  On Error GoTo UpdateErr
 
  With ioTIPOPAGO
    If .Text = "" Then
        lblstatus.Caption = "No se permite TIPO de PAGO en blanco"
        .SetFocus
        Exit Sub
    End If
  End With
    
  'si es pagos a proveedores
  If Pago_A_Proveedor Then
  
  With cbCODPROV
    If .Text = "" Then
        'solo si tipo pago es pago por proveedor
        If ioTIPOPAGO.Text = "1" Then
            lblstatus.Caption = "No se permite PROVEEDOR en blanco"
            .SetFocus
            Exit Sub
        End If
    End If
  End With
  
  With ioPEDIDO
    If (.Text = "") And (ioFACTURA.Text = "") Then
        lblstatus.Caption = "No se permite PEDIDO y FACTURA ambos en blanco"
        .SetFocus
        Exit Sub
    End If
  End With
  
  'asignar la caja actual (pues caja esta sin cargar)
  rc.fields("CODCAJA") = CajaActual
  rc.fields("CUOTA") = ioCUOTA.Text
    
 'si es un pago realizado por un dependiente
 Else
    
     rc.fields("PAGADO") = ioIMPORTE.Text
     
    With cbCODCAJA
    If .Text = "" Then
        lblstatus.Caption = "No se permite CAJA en blanco"
        .SetFocus
        Exit Sub
    End If
  End With
  
 End If
  
  With ioIMPORTE
    If .Text = "" Then
        lblstatus.Caption = "No se permite IMPORTE en blanco"
        .SetFocus
        Exit Sub
    End If
  End With
  
  With ioDESCRIPCION
    If .Text = "" Then
        lblstatus.Caption = "No se permite DESCRIPCION en blanco"
        .SetFocus
        Exit Sub
    End If
  End With
  

 
  
  rc.fields("IMPORTE") = ioIMPORTE.Text
  
 
  
  rc.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    rc.MoveLast              'va al nuevo registro
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  
  lblstatus.Caption = ""
  
  cbAgregar.SetFocus

  Exit Sub
UpdateErr:
  If Err.Number = -2147217887 Then Exit Sub
  MsgBox Err.Description, vbInformation, "Atención"
End Sub

Private Sub cbcerrar_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  rc.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
If Err.Number = -2147217842 Then Exit Sub
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  rc.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
If Err.Number = -2147217842 Then Exit Sub
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  If Not rc.EOF Then rc.MoveNext
  If rc.EOF And rc.RecordCount > 0 Then
    Beep
     'ha sobrepasado el final; vuelva atrás
    rc.MoveLast
  End If
  'muestra el registro actual
  mbDataChanged = False

  Exit Sub
GoNextError:
If Err.Number = -2147217842 Then Exit Sub
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  If Not rc.BOF Then rc.MovePrevious
  If rc.BOF And rc.RecordCount > 0 Then
    Beep
    'ha sobrepasado el final; vuelva atrás
    rc.MoveFirst
  End If
  'muestra el registro actual
  mbDataChanged = False

  Exit Sub

GoPrevError:
If Err.Number = -2147217842 Then Exit Sub
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cbAgregar.Visible = bVal
  cmNuevoPago.Visible = bVal
  cbEdicion.Visible = bVal
  cbActualizar.Visible = Not bVal
  cbCancelar.Visible = Not bVal
  cbEliminar.Visible = bVal
  cbCerrar.Visible = bVal
  cbLista.Visible = bVal
   
  'cbActualizar.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub


Private Sub carga_grid()
Dim rcdet As New ADODB.Recordset

   On Error GoTo carga_grid_Error

        With fg
            .Clear
            .Cols = 4
            .Rows = 1
            .ColHidden(1) = True
            
        
            .TextMatrix(0, 2) = "Importe"
            .TextMatrix(0, 3) = "Fecha"
            
            .ColFormat(2) = "Currency"
            
        
        rcdet.Open "SELECT * FROM DETPAGOS WHERE MBAJA = 0 AND CODIGO = " & rc.fields("CODIGO") & " AND CODCAJA = " & rc.fields("CODCAJA"), locCnn, adOpenStatic, adLockReadOnly
        
        
        If rcdet.RecordCount <= 0 Then
            Exit Sub
        End If
        
        
        Do Until rcdet.EOF
        
            
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = rcdet.fields("LINEA")
            .TextMatrix(.Rows - 1, 2) = rcdet.fields("IMPORTE")
            .TextMatrix(.Rows - 1, 3) = rcdet.fields("FMODI")
            
            rcdet.MoveNext
        
        Loop
        
        
            .SubtotalPosition = flexSTAbove
            .subtotal flexSTSum, , 2, , vbBlue, vbWhite, True
            .AutoSize 1, .Cols - 1
                    
        End With

   On Error GoTo 0
   Exit Sub

carga_grid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_grid de Formulario frmDPagosPr"
End Sub

