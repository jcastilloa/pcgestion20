VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmDetVenta 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Artículos ..."
   ClientHeight    =   6690
   ClientLeft      =   2010
   ClientTop       =   2370
   ClientWidth     =   10425
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
   ScaleHeight     =   6690
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin PCGestion.chameleonButton cmdPrevious 
      Height          =   630
      Left            =   1080
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5205
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
      MICON           =   "frmDetVenta.frx":0000
      PICN            =   "frmDetVenta.frx":001C
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
      Left            =   30
      Top             =   1530
      Width           =   10365
      _ExtentX        =   18283
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
   Begin PCGestion.chameleonButton cmdFirst 
      Height          =   630
      Left            =   45
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5205
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
      MICON           =   "frmDetVenta.frx":0CEE
      PICN            =   "frmDetVenta.frx":0D0A
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
      Left            =   8280
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5205
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
      MICON           =   "frmDetVenta.frx":1A40
      PICN            =   "frmDetVenta.frx":1A5C
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
      Left            =   9345
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5205
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
      MICON           =   "frmDetVenta.frx":272E
      PICN            =   "frmDetVenta.frx":274A
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5865
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
      MICON           =   "frmDetVenta.frx":3480
      PICN            =   "frmDetVenta.frx":349C
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
      TabIndex        =   14
      Top             =   5865
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
      MICON           =   "frmDetVenta.frx":4176
      PICN            =   "frmDetVenta.frx":4192
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5865
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
      MICON           =   "frmDetVenta.frx":4A6C
      PICN            =   "frmDetVenta.frx":4A88
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
      Left            =   7245
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5865
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
      MICON           =   "frmDetVenta.frx":52E6
      PICN            =   "frmDetVenta.frx":5302
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
      Left            =   8220
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5865
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
      MICON           =   "frmDetVenta.frx":5BDC
      PICN            =   "frmDetVenta.frx":5BF8
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
      Left            =   9345
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5865
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
      MICON           =   "frmDetVenta.frx":67CA
      PICN            =   "frmDetVenta.frx":67E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.miText ioCODBAR 
      Height          =   525
      Left            =   330
      TabIndex        =   0
      Top             =   495
      Width           =   3735
      _ExtentX        =   6588
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
   Begin VSFlex8Ctl.VSFlexGrid fgArt 
      Height          =   3255
      Left            =   30
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1935
      Width           =   10380
      _cx             =   18309
      _cy             =   5741
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
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
      FormatString    =   $"frmDetVenta.frx":74C0
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
   Begin PCGestion.bsGradientLabel lblModelo 
      Height          =   420
      Left            =   5565
      Top             =   525
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   741
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
      Colour1         =   11513775
      Colour2         =   16558731
      CaptionAlignment=   1
   End
   Begin PCGestion.miText ioUNIDADES 
      Height          =   525
      Left            =   4485
      TabIndex        =   1
      Top             =   495
      Width           =   1080
      _ExtentX        =   1905
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
   Begin PCGestion.bsGradientLabel lblTalla 
      Height          =   420
      Left            =   780
      Top             =   1035
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   741
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
      Colour1         =   11513775
      Colour2         =   16558731
      CaptionAlignment=   1
   End
   Begin PCGestion.miText ioPREVEN 
      Height          =   525
      Left            =   6900
      TabIndex        =   2
      Top             =   1035
      Width           =   1185
      _ExtentX        =   2170
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
   Begin PCGestion.miText ioIVA 
      Height          =   525
      Left            =   9735
      TabIndex        =   4
      Top             =   1035
      Width           =   675
      _ExtentX        =   1191
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
   Begin PCGestion.bsGradientLabel lblColorDesc 
      Height          =   420
      Left            =   4515
      Top             =   1050
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   741
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
      Colour1         =   11513775
      Colour2         =   16558731
      CaptionAlignment=   1
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel1 
      Height          =   345
      Left            =   4275
      Top             =   5925
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   609
      Caption         =   "- T - Terminar "
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   15640462
      Colour2         =   7177785
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel2 
      Height          =   345
      Left            =   4275
      Top             =   6285
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   609
      Caption         =   "- S - Salir"
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   15640462
      Colour2         =   7177785
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel3 
      Height          =   345
      Left            =   4275
      Top             =   5565
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   609
      Caption         =   "- I - Ir a Rejilla"
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   15640462
      Colour2         =   7177785
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel4 
      Height          =   345
      Left            =   4275
      Top             =   5205
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   609
      Caption         =   "- B - Borrar último"
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   15640462
      Colour2         =   7177785
   End
   Begin PCGestion.miText ioDCTO 
      Height          =   525
      Left            =   8760
      TabIndex        =   3
      Top             =   1035
      Width           =   675
      _ExtentX        =   1191
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
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DCTO"
      Height          =   300
      Left            =   8145
      TabIndex        =   29
      Top             =   1110
      Width           =   585
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "IVA"
      Height          =   300
      Left            =   9315
      TabIndex        =   28
      Top             =   1110
      Width           =   420
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PVP"
      Height          =   300
      Left            =   6450
      TabIndex        =   27
      Top             =   1110
      Width           =   420
   End
   Begin VB.Label lblColor 
      Height          =   435
      Left            =   3345
      TabIndex        =   26
      Top             =   1035
      Width           =   1125
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "COLOR"
      Height          =   330
      Left            =   2565
      TabIndex        =   25
      Top             =   1095
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TALLA"
      Height          =   330
      Left            =   -15
      TabIndex        =   24
      Top             =   1095
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "UDS."
      Height          =   360
      Left            =   4035
      TabIndex        =   23
      Top             =   585
      Width           =   495
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      Height          =   330
      Left            =   15
      TabIndex        =   22
      Top             =   90
      Width           =   285
   End
   Begin VB.Label ioID 
      Alignment       =   1  'Right Justify
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
      Left            =   360
      TabIndex        =   21
      Top             =   75
      Width           =   1485
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CB"
      Height          =   360
      Left            =   -15
      TabIndex        =   19
      Top             =   570
      Width           =   315
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
      Left            =   7935
      TabIndex        =   8
      Top             =   90
      Width           =   2445
   End
   Begin VB.Label ioCODVEN 
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
      Left            =   2580
      TabIndex        =   7
      Top             =   75
      Width           =   1440
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "VENTA"
      Height          =   330
      Left            =   1815
      TabIndex        =   6
      Top             =   90
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ultima modificación"
      Height          =   315
      Left            =   5805
      TabIndex        =   5
      Top             =   105
      Width           =   2040
   End
End
Attribute VB_Name = "frmDetVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------------
' Module    : frmDetVenta
' DateTime  : 31/10/2003 10:08
' Author    : Administrador
' Purpose   : Artículos para la venta. Detalle.
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
Public WithEvents rc As ADODB.Recordset
Attribute rc.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim prime As Boolean

'para q no nos haga los devuelve_campo de rc_movecomplete
'durante la carga del grid
Dim cargando_Grid As Boolean

'Dim oSQL As New clsSmartSQL

Dim Codigo_B As MiCodBar

Dim conta_lineas As Long

'codigo de la venta actual, para añadir nuevos tegistros
Public CODIGO_VENTA As Long
Public IR_A_ID As Integer

Public CODIGO_USR As Long

Dim crear_arreglo As Boolean
Dim edita_arreglo As Boolean

'código del arreglo que se esta editando
Dim cod_arr_edit As Long

'------------------------------------------------------------------------------------------------
'un tipo predefinido para guardar los datos para establecer la oferta 2x1
'(OfertaActual = 1)
Private Type Of2x1
    CODIGO_VENTA As Long   'venta a la que pertenece el 2x1
    id() As Integer                'IDs que se ven implicados en la operación
    Preven() As Double          'PVP para esos IDs
    cuantos As Byte              'cuantos llevamos (valores permitidos 1 o 2)
End Type

Dim mi2x1 As Of2x1
'------------------------------------------------------------------------------------------------


Private Sub fgArt_Click()

   On Error GoTo fgArt_Click_Error

If fgArt.TextMatrix(fgArt.Row, 1) <> "" Then

    'si es un artículo
    If (IsNumeric(fgArt.TextMatrix(fgArt.Row, 1)) And fgArt.TextMatrix(fgArt.Row, 11) = "") Then
        rc.Find "LINEA = " & fgArt.TextMatrix(fgArt.Row, 1), , adSearchForward, 1
    'si es un arreglo
    Else
        lblModelo.Caption = "Arreglo Varios"
        ioPREVEN.Text = fgArt.TextMatrix(fgArt.Row, 6)
        cod_arr_edit = fgArt.TextMatrix(fgArt.Row, 1)
        edita_arreglo = True
    End If
    
    DoEvents
     
    'salir del grid para editar
    fgArt.TabStop = False
    
    Call cbedicion_Click

End If

   On Error GoTo 0
   Exit Sub

fgArt_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento fgArt_Click de Formulario frmDetVenta"

End Sub

Private Sub fgArt_KeyPress(KeyAscii As Integer)

'si pulsa intro, ir a articulo
   On Error GoTo fgArt_KeyPress_Error

If KeyAscii = 13 Then

'ir a artículo
Call fgArt_Click

End If

   On Error GoTo 0
   Exit Sub

fgArt_KeyPress_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento fgArt_KeyPress de Formulario frmDetVenta"

End Sub

Private Sub Form_Activate()

   On Error GoTo Form_Activate_Error

If Not prime Then

  'si entra como IR_A_ID = 0, añadir un nuevo registro. De lo contrario
  'ir al registro con un find
  If IR_A_ID = 0 Then
        
        Call cbAgregar_Click
        
  Else
                        
        Call cmdFirst_Click
        rc.Find "LINEA = " & IR_A_ID
        DoEvents
        Call cbedicion_Click
               
  End If

  prime = True
End If

   On Error GoTo 0
   Exit Sub

Form_Activate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Form_Activate de Formulario frmDetPtrans"
    
End Sub



Private Sub Form_Load()
  
   'Move (Screen.Width - Width) \ 2, Separacion_MDIForm
   
   With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
   End With
   
 ' Set rc = New Recordset
 ' oSQL.AddTable "DETVENTA"
 ' oSQL.AddSimpleWhereClause "CODVEN", CODIGO_VENTA
 ' oSQL.AddSimpleWhereClause "CODCAJA", CajaActual
 ' oSQL.AddOrderClause "LINEA"

 ' rc.Open oSQL.SQL, locCnn, adOpenStatic, adLockOptimistic
 '

''''''''
'////////////////////////////////////////////////////////
' Enlazar controles:
'////////////////////////////////////////////////////////

With ioID
    Set .DataSource = rc
        .DataField = "LINEA"
End With

With ioCODVEN
    Set .DataSource = rc
        .DataField = "CODVEN"
End With
  
With ioCODBAR
         .SoloNumeros = True
         .LongMaxima = LenCodBar
         .dspFormat = ""
         .intFormat = ""
         .PermitirBlanco = False
End With

With ioPREVEN
         .SoloNumeros = True
         .LongMaxima = 12
         .Alineacion = 1
         .dspFormat = "Currency"
         .PermitirBlanco = False
End With

With ioIVA
         .SoloNumeros = True
         .LongMaxima = 2
         .PermitirBlanco = False
End With

With ioDCTO
         .SoloNumeros = True
         .LongMaxima = 3
         .PermitirBlanco = False
End With

With ioUNIDADES
         .SoloNumeros = True
         .LongMaxima = 8
         .Alineacion = 1
         .PermitirBlanco = True
       '  .DataField = "UNIDADES"
       '  Set .DataSource = rc
End With
  
 
  'poner títulos al grid
  'With fgArt
  '      .TextMatrix(0, 1) = "Artículo"
  '      .TextMatrix(0, 2) = "Talla"
  '      .TextMatrix(0, 3) = "Color"
  '      .TextMatrix(0, 4) = "Unidades"
  '      .TextMatrix(0, 5) = "PVP"
  '      .TextMatrix(0, 6) = "Importe"
  '      .TextMatrix(0, 7) = "IVA"
  '      .TextMatrix(0, 8) = "Total"
  'End With

 If OfertaActual > 0 Then Me.Caption = Me.Caption & "    " & OfertaDSC

  'cargar el grid con los artículos de esta venta
  Call cargar_grid
  
  mbDataChanged = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If (mbEditFlag Or mbAddNewFlag) And Not (KeyCode = vbKeyT Or KeyCode = vbKeyS) Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      
      If MsgBox("¿Desea cerrar la ventana actual?", vbQuestion + vbYesNo, "Cerrar") = vbYes Then
        cbcerrar_Click
        KeyCode = 0
      End If
      
    Case vbKeyEnd
      cmdLast_Click
      KeyCode = 0
    Case vbKeyHome
      cmdFirst_Click
      KeyCode = 0
      
    Case vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
        KeyCode = 0
      Else
        cmdPrevious_Click
        KeyCode = 0
      End If
      
    Case vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
        KeyCode = 0
      Else
        cmdNext_Click
        KeyCode = 0
      End If
      
      Case vbKeyF1
            Call cbAgregar_Click
            KeyCode = 0
        
      Case vbKeyF2
            Call cbactualizar_Click
            KeyCode = 0
        
      Case vbKeyF3
            Call cbedicion_Click
            KeyCode = 0

      Case vbKeyF5
            Call cmdFirst_Click
            KeyCode = 0
           
       Case vbKeyF6
            Call cmdPrevious_Click
            KeyCode = 0
      
       Case vbKeyF7
            Call cmdNext_Click
            KeyCode = 0
    
       Case vbKeyF8
        Call cmdLast_Click
        KeyCode = 0
        
       'terminar edicion
       'termina de añadir nuevos articulos ...
       Case vbKeyT
       
       'cancelar ...
        Call cbCancelar_Click
        KeyCode = 0
        
        'ir a grid despues de terminar
        Call Form_KeyDown(vbKeyI, 0)
                   
       'salir, sale de esta ventana
       Case vbKeyS
        
        Call cbcerrar_Click
        KeyCode = 0
        
       'ir a grid, solo si no esta editando o  añadiendo
       Case vbKeyI
       
        If (mbEditFlag Or mbAddNewFlag) Then Exit Sub
        
        'ya se encuentra en el grid, salir de el
        If fgArt.TabStop Then
        
            fgArt.TabStop = False
            ioCODBAR.SetFocus
        
        Else    'ir a grid
                
            fgArt.TabStop = True
            fgArt.Select 1, 1, 1, fgArt.Cols - 1
            fgArt.SetFocus
      
        End If
        KeyCode = 0
        
       'Borrar último registro ...
       Case vbKeyB
       
       'si no hay registros salir (y borrar los campos q pudieran
       'quedar con algo
       If rc.RecordCount <= 0 Then
       ' ioUNIDADES.Text = ""
       ' ioPREVEN.Text = ""
       ' ioIVA.Text = ""
       ' Exit Sub
       End If
       
       'si no lo quiere quitar ...
       If Not rc.EOF Then rc.MoveLast
       If MsgBox("¿Desea quitar el artículo: " & Chr(13) & lblModelo.Caption & " - " & lblTalla.Caption & " - " & lblColorDesc.Caption & " - Uds(" & ioUNIDADES.Text & ")", vbQuestion + vbYesNo) = vbNo Then Exit Sub
       
       
       rc.Delete
       
       DoEvents
       
       If Not rc.BOF Then rc.MoveFirst
              
      ' rc.Requery
       
       
       Call cargar_grid
       
       'si no hay registros y borrar los campos q pudieran
       'quedar con algo
       If rc.RecordCount <= 0 Then
        ioUNIDADES.Text = ""
        ioPREVEN.Text = ""
        ioIVA.Text = ""
        lblColor.BackColor = Me.BackColor
        Exit Sub
       End If
       
       KeyCode = 0
      
  End Select
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error Resume Next

If rc.EditMode <> adEditNone Then rc.CancelUpdate

rc.Close
Set rc = Nothing

'Set oSQL = Nothing

Set frmDetVenta = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub


Private Sub ioDCTO_Validate(Cancel As Boolean)

If ioDCTO.Text <> "" Then
    If ioDCTO.Text > 100 Then
        MsgBox "No se permite un descuento mayor del 100 %", vbExclamation, titulo
        Cancel = True
        ioDCTO.SetFocus
        DoEvents
        ioDCTO.CancelarValidacion
    End If
End If

End Sub

Private Sub ioIVA_lostfocus()

If Not (mbEditFlag Or mbAddNewFlag) Then Exit Sub

If ioIVA.Text = "" Then ioIVA.Text = "0"

Call cbactualizar_Click

End Sub

Private Sub rc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  
  'Esto mostrará la posición de registro actual para este Recordset
  If rc.AbsolutePosition > 0 And Not cargando_Grid Then
    
   lblstatus.Caption = "Registro: " & CStr(rc.AbsolutePosition)
   
 '  If adReason = adRsnAddNew Then Exit Sub
  ' if rc.Fields("UNIDADES").Value is nothing
   
   ioUNIDADES.Text = rc.fields("UNIDADES").Value
   ioPREVEN.Text = rc.fields("PREVEN").Value
   ioIVA.Text = rc.fields("IVA").Value
   ioDCTO.Text = rc.fields("DCTO").Value
      
   'la talla
   If rc.fields("CODTALLA") > 0 Then
    lblTalla.Caption = Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & rc.fields("CODTALLA").Value, locCnn))
   End If
   'el color
   
   If rc.fields("CODCOL") > 0 Then
    lblColor.BackColor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & rc.fields("CODCOL").Value, locCnn)
    lblColorDesc.Caption = Trim(devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & rc.fields("CODCOL").Value, locCnn))
   End If

   If rc.fields("CODART") > 0 Then
   lblModelo.Caption = Trim(devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & rc.fields("CODART") & " AND TEMPOR = " & rc.fields("TEMPOR"), locCnn))
   End If
     
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

'---------------------------------------------------------------------------------------
Private Sub cbAgregar_Click()
Dim tmpcodigo As Variant
  
   On Error GoTo cbAgregar_Click_Error

  
  With rc
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    
       .AddNew
       
       If ioIVA.Text = "" Then ioIVA.Text = "0"
       If ioDCTO.Text = "" Then ioDCTO.Text = "0"
    
    
    tmpcodigo = devuelve_campo("select max(LINEA) + 1 from DETVENTA WHERE CODCAJA = " & CajaActual & " AND CODVEN = " & CODIGO_VENTA, locCnn)
    
    'Si devuelve @ esque ha habido un error
    If tmpcodigo <> "@" Then
    
    .fields("LINEA") = CLng(tmpcodigo)
    
    Else
    
    .fields("LINEA") = 1
    
    End If
    
    .fields("CODVEN") = CODIGO_VENTA
    .fields("CODCAJA") = CajaActual
    
    Set tmpcodigo = Nothing
    
    'guardar q estamos en oferta
    If OfertaActual = 1 Then
    
        With mi2x1
        
            Select Case .cuantos
                Case 0
                    lblstatus.Caption = "Introduzca la PRIMERA prenda 2x1"
                Case 1
                    lblstatus.Caption = "Introduzca la SEGUNDA prenda 2x1"
            End Select
        
        End With
        
    
    Else
        lblstatus.Caption = "Agregar registro"
        
    End If

    'Set tmpcodigo = Nothing
    
    
    mbAddNewFlag = True
    SetButtons False
    cbActualizar.Visible = True
    ioCODBAR.SetFocus
    
  End With

   On Error GoTo 0
   Exit Sub

cbAgregar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbAgregar_Click de Formulario frmDetPtrans"

End Sub


Private Sub cbedicion_Click()
  On Error GoTo EditErr

  If edita_arreglo Then
  lblstatus.Caption = "Modificar Arreglo"
  Else
  lblstatus.Caption = "Modificar Artículo"
  End If
  
  If edita_arreglo = False Then mbEditFlag = True
  
  SetButtons False
  cbActualizar.Visible = True
  
  ioCODBAR.Enabled = False
  
  If edita_arreglo Then
    ioPREVEN.SetFocus
  Else
    ioUNIDADES.SetFocus
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
  edita_arreglo = False
  crear_arreglo = False
  cod_arr_edit = 0
  
  rc.CancelUpdate
  If mvBookMark > 0 Then
    rc.Bookmark = mvBookMark
  Else
    rc.MoveFirst
  End If
  mbDataChanged = False
   ioCODBAR.Enabled = True

End Sub

Private Sub cbactualizar_Click()
Dim tmpsal As Long
  On Error GoTo UpdateErr

  If crear_arreglo Then
    
    If ioPREVEN.Text = "" Then
        lblstatus.Caption = "No se permite PVP en blanco"
        ioPREVEN.SetFocus
        Exit Sub
    End If
    
    tmpsal = añade_arreglo_rapido(CODIGO_VENTA, CajaActual, 0, CCur(ioPREVEN.Text), " - Editar -", CODIGO_USR, locCnn)
    Call añade_arreglo_grid(ioPREVEN.Text, tmpsal)
    
    If tmpsal > 0 Then
        lblstatus.Caption = "Se ha añadido un arreglo satisfactoriamente"
    Else
        lblstatus.Caption = "Error al crear el arreglo"
    End If
    
    crear_arreglo = False
    
    
    mbEditFlag = False
    mbAddNewFlag = False
    mbDataChanged = False
    
    ioCODBAR.Text = ""
    ioPREVEN.Text = ""
    ioIVA.Text = ""
    ioUNIDADES.Text = "1"
    lblModelo.Caption = ""

    'agregar un nuevo registro para continuar el ciclo
    Call cbAgregar_Click
    
    Exit Sub
    
  End If
  
  If edita_arreglo Then
        
        If ioPREVEN.Text = "" Then
            lblstatus.Caption = "No se permite PVP en blanco"
            ioPREVEN.SetFocus
            Exit Sub
        End If
    
        locCnn.Execute "UPDATE ARREGLOS SET PVP = " & Replace(CDbl(ioPREVEN.Text), ",", ".") & " WHERE ID = " & cod_arr_edit & " AND CODCAJ = " & CajaActual
    
        lblstatus.Caption = "Se ha actualizado el arreglo satisfactoriamente"
  
        cod_arr_edit = 0
    
        edita_arreglo = False
    
        mbEditFlag = False
        mbAddNewFlag = False
        mbDataChanged = False
    
        ioCODBAR.Text = ""
        ioPREVEN.Text = ""
        ioIVA.Text = ""
        ioUNIDADES.Text = "1"
        lblModelo.Caption = ""
    
        Call cbCancelar_Click
    
        DoEvents
    
        Call cargar_grid
        'agregar un nuevo registro para continuar el ciclo
        'Call cbAgregar_Click
        Exit Sub
    
  End If
  
  With ioUNIDADES
    
    'si es distinto de blanco
    If .Text <> "" Then
        
        'que sea mayor de 0
        If CDbl(.Valor) <= 0 Then
    
        lblstatus.Caption = "No se permiten 0 unidades"
        .SetFocus
        .CancelarValidacion
        Exit Sub
    
        End If
        
        'si no hay unidades para ese artículo
        If comprueba_unidades = False Then
        '
        End If
    
    Else

        lblstatus.Caption = "No se permite unidades en blanco"
        .SetFocus
        .CancelarValidacion
        Exit Sub
    
    End If
    
  End With
  
  
  If ioDCTO.Text <> "" Then
    If ioDCTO.Text > 100 Then
        MsgBox "No se permite un descuento mayor del 100 %", vbExclamation, titulo
        ioDCTO.SetFocus
        DoEvents
        ioDCTO.CancelarValidacion
        Exit Sub
    End If
  End If
  
  If IsNull(rc.fields("CODCOL")) Then Exit Sub
  If IsNull(rc.fields("CODTALLA")) Then Exit Sub
    
  If ioIVA.Text = "" Then ioIVA.Text = "0"
  If ioDCTO.Text = "" Then ioDCTO.Text = "0"
  
  rc.fields("DCTO") = ioDCTO.Text
  rc.fields("CODVEN") = CODIGO_VENTA
  rc.fields("UNIDADES") = CDbl(ioUNIDADES.Valor)
  rc.fields("PREVEN").Value = ioPREVEN.Text
  rc.fields("IVA") = ioIVA.Text
  rc.fields("CODCAJA") = CajaActual
    
  'guardar q estamos en oferta
  If OfertaActual > 0 Then rc.fields("OFERTA") = True
  
  
  'ACTUALIZAR
  rc.UpdateBatch adAffectAll
  
  DoEvents
    
  If OfertaActual = 1 Then
    
    With mi2x1
        .cuantos = .cuantos + 1
        .CODIGO_VENTA = CODIGO_VENTA
        
        ReDim Preserve .id(.cuantos)
        ReDim Preserve .Preven(.cuantos)
        
        .id(.cuantos) = rc.fields("LINEA")
        .Preven(.cuantos) = rc.fields("PREVEN")
                
        If .cuantos = 2 Then
                
                'respetar el precio del artículo mayor, y el precio del meno ponerlo a 0
                If .Preven(1) > .Preven(2) Then
                
                    'poner el segundo a 0  (se puede cambiar por DCTO)
                    
                    rc.Find "LINEA = " & .id(2), , adSearchForward, 1
                    
                   ' locCnn.Execute "UPDATE DETVENTA SET PREVEN = 0 WHERE LINEA = " & .id(2) & " AND CODVEN = " & .CODIGO_VENTA & " AND CODCAJA = " & CajaActual
                    
                Else
                
                    rc.Find "LINEA = " & .id(1), , adSearchForward, 1
                    'poner el primero a 0   (se puede cambiar por DCTO)
                   ' locCnn.Execute "UPDATE DETVENTA SET PREVEN = 0 WHERE LINEA = " & .id(1) & " AND CODVEN = " & .CODIGO_VENTA & " AND CODCAJA = " & CajaActual
              
                
                End If
                
                
                rc.fields("PREVEN") = 0
                rc.UpdateBatch adAffectAll
                
                
                
                DoEvents
                rc.Requery
                
                'Call cargar_grid
                
                'poner todo a 0
                .cuantos = 0
                ReDim .id(.cuantos)
                ReDim .Preven(.cuantos)
                
                
        End If
        
                
    End With
    
  End If

  

  SetButtons True
  
  
  DoEvents
  'añade el articulo actual al grid de modificaciones

  'Call añade_linea_grid(rc.Fields("CODART"), rc.Fields("TEMPOR"), rc.Fields("CODTALLA"), rc.Fields("CODCOL"), rc.Fields("UNIDADES"), rc.Fields("PREVEN"), rc.Fields("IVA"))
  
  ioCODBAR.Text = ""
  ioPREVEN.Text = ""
  ioIVA.Text = ""
  ioUNIDADES.Text = "1"
  
  lblColor.BackColor = Me.BackColor
  DoEvents
  
  'ioCODART.Text = ""
  'cbTEMPOR.Text = ""
  'ioUNIDADES.Text = "0"
  'cbCODTALLA.Text = ""
  'cbCODCOL.Text = ""
  
  If mbAddNewFlag Then
    rc.MoveLast              'va al nuevo registro
  End If

 
 
  
  lblstatus.Caption = ""
  
  DoEvents
  
  Call cargar_grid
  
  DoEvents

  ioCODBAR.Enabled = True
  
  rc.Move 0
  
  If mbEditFlag Then
  
        mbEditFlag = False
        mbAddNewFlag = False
        mbDataChanged = False
  
  Else
  
        mbEditFlag = False
        mbAddNewFlag = False
        mbDataChanged = False

  'agregar un nuevo registro para continuar el ciclo
    Call cbAgregar_Click


  End If
  
  
  
  
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
  cbEdicion.Visible = bVal
  cbActualizar.Visible = Not bVal
  cbCancelar.Visible = Not bVal
  cbEliminar.Visible = bVal
  cbCerrar.Visible = bVal
'  cbLista.Visible = bVal
   
  'cbActualizar.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
 
  
End Sub


'---------------------------------------------------------------------------------------
' Subrutina   : ioCODBAR_validate
' Fecha/Hora  : 07/12/2003 11:01
' Autor       : JCASTILLO
' Propósito   : Descompone el codigo de barras y introduce la información en los campos
'               correspondientes
'---------------------------------------------------------------------------------------
Private Sub ioCODBAR_Validate(Cancel As Boolean)
Dim campos As Variant
Dim cadena As String

   On Error GoTo ioCODBAR_Validate_Error


crear_arreglo = False

With ioCODBAR

If Trim(.Text) = "" Then
    Cancel = True
    .CancelarValidacion
    DoEvents
    Exit Sub
End If

If Not IsNumeric(.Text) Then
    Cancel = True
    .CancelarValidacion
    DoEvents
    Exit Sub
End If

    'si es un arreglo
    If Trim(.Text) = CodigoArreglos Then
       
        crear_arreglo = True
        
        'abortar la creacion de un nuevo registro
        rc.CancelUpdate
        mbAddNewFlag = False
        
        DoEvents
        
        ioUNIDADES.Text = 1
        ioPREVEN.Text = 0
        lblModelo.Caption = "Arreglo Varios"
        
        DoEvents
         
        Exit Sub
    End If

    'si es un codigo de barras con la longitud válidad
    'o un codigo de un digito para los restos
    'RES1
    'buscar por referencia "RES" + el codigo de un digito
    'introducido
    If (Len(Trim(.Text)) = LenCodBar) Or (Len(Trim(.Text)) = 1) Then
        
        'comprobar si existe el artículo/temporada
        If (Len(Trim(.Text)) = 1) Then
            cadena = "SELECT MODELO FROM MAARTIC WHERE REF = 'RES" & Trim(.Text) & "' AND TEMPOR = " & TemporadaActual
        Else
            cadena = "SELECT MODELO FROM MAARTIC WHERE CODIGO = " & CLng(Left(.Text, 5)) & " AND TEMPOR = " & CLng(Mid(.Text, 6, 3))
        End If
        
        If devuelve_campo(cadena, locCnn) = "@" Then
        
                lblstatus.Caption = "No existe el artículo para esa temporada!, Codigo de Barras no Válido"
                ioCODBAR.Text = ""
                .CancelarValidacion
                Cancel = True
                       
                Beep
                Call Espera(1)
                Beep
                Call Espera(1)
                Beep
                
                Exit Sub
                
        End If
                
        
        If Len(Trim(.Text)) = 1 Then
        
            Codigo_B.CODIGO_ART = Trim(.Text)
            Codigo_B.TEMPORADA_ART = TemporadaActual
            Codigo_B.TALLA_ART = "0"
            Codigo_B.COLOR_ART = "0"
        Else
            Codigo_B = Descompone_CBAR(ioCODBAR.Text)
            
        End If
        
        
        '0 MODELO
        '1 PREVEN
        '2 TIPOIVA
        '3 STOCKMIN
        '4 STOCKMAX
        
                
        If (Len(Trim(.Text)) = 1) Then
            cadena = "SELECT MODELO, PREVEN, TIPOIVA, STOCKMIN, STOCKMAX, CODIGO FROM MAARTIC WHERE REF = 'RES" & Trim(.Text) & "' AND TEMPOR = " & TemporadaActual
        Else
            cadena = "SELECT MODELO, PREVEN, TIPOIVA, STOCKMIN, STOCKMAX FROM MAARTIC WHERE CODIGO = " & Codigo_B.CODIGO_ART & " AND TEMPOR = " & Codigo_B.TEMPORADA_ART
        End If
        
        campos = devuelve_matriz(cadena, locCnn)
        
        If (Len(Trim(.Text)) = 1) Then Codigo_B.CODIGO_ART = campos(5)
             'codigo de artículo
             
        'obtener el % de iva
        ioIVA.Text = devuelve_campo("SELECT IVA FROM IVA WHERE CODIGO = " & campos(2), locCnn)
        'la talla
        
        If Trim(Codigo_B.TALLA_ART) <> "" Then _
        lblTalla.Caption = Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & Codigo_B.TALLA_ART, locCnn))
        'el color
        
        If Trim(Codigo_B.COLOR_ART) <> "" Then
            lblColor.BackColor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & Codigo_B.COLOR_ART, locCnn)
            lblColorDesc.Caption = Trim(devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & Codigo_B.COLOR_ART, locCnn))
        End If

        lblModelo.Caption = Trim(campos(0))
        ioPREVEN.Text = campos(1)
        
        If Trim(Codigo_B.TALLA_ART) <> "" Then
            rc.fields("CODTALLA") = Codigo_B.TALLA_ART
        Else
            rc.fields("CODTALLA") = 0
        End If
        
        If Trim(Codigo_B.COLOR_ART) <> "" Then
            rc.fields("CODCOL") = Codigo_B.COLOR_ART
        Else
            rc.fields("CODCOL") = 0
        End If
        
        rc.fields("TEMPOR") = Codigo_B.TEMPORADA_ART
        rc.fields("CODART") = Codigo_B.CODIGO_ART
        
        ioUNIDADES.Text = "1"
        'ioPREVEN.SetFocus
        

    Else

        lblstatus.Caption = "Código de Barras no válido"
        ioCODBAR.Text = ""
        .CancelarValidacion
        Cancel = True
        Beep
        DoEvents
        Beep
        DoEvents
        Beep
        DoEvents
        Exit Sub

    End If

End With


'si la oferta actual OfertaActual = 1 (2x1), entonces que pase directamente al
'siguiente artículo, si la oferta es 2 (%) o 3 (precio fijo) que permita cambiar el
'importe y dcto de la prenda
If (Len(Trim(ioCODBAR.Text)) = LenCodBar) And (OfertaActual <= 1) Then

    Call cbactualizar_Click
    DoEvents
    DoEvents
    Cancel = True
    
ElseIf Len(Trim(ioCODBAR.Text)) = 1 And (OfertaActual > 1) Then

    Select Case OfertaActual
    '%
    Case 2
        If crear_arreglo = False Then ioDCTO.Text = OfertaDcto
    'precio fijo
    Case 3
        If crear_arreglo = False Then ioPREVEN.Text = OfertaImp
    End Select
    
    ioPREVEN.SetFocus
    
End If

   On Error GoTo 0
   Exit Sub

ioCODBAR_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioCODBAR_validate de Formulario frmDetPtrans"

End Sub

Private Sub ioCODVEN_Change()
ioCODVEN.Caption = Format(ioCODVEN.Caption, "00000")
End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : Carga_Almacenes_Origen
' Fecha/Hora     : 25/11/2003 18:13
' Autor             : JCastillo
' Propósito       : carga la lista de almacenes que tienen el articulo:
'                        CODART,TEMPOR,TALLA,COLOR en un miCombo
'                        Carga una lista de los almacenes que tienen existencias actualmente
'                        de ese artículo para poder seleccionarlos como ORIGEN.
'                        Conexión al Servidor.
'---------------------------------------------------------------------------------------

'Private Sub carga_almacenes_origen(miC As miCombo)
'Dim tmprc As New ADODB.Recordset
'
 '  On Error GoTo carga_almacenes_origen_Error
'
 '  miC.borra_combo
  ' 'si alguno de los campos clave esta en blanco, salir
'   If ioCODART.Text = "" Or cbTEMPOR.Text = "" Or cbCODTALLA.Text = "" Or _
 '     cbCODCOL.Text = "" Then Exit Sub
'
 ' With SrvCnn
'    If .State = 0 Then
 '       .CursorLocation = adUseClient
'  '      .Open strSrvCnn
'    End If
'  End With
'
'  'abrir
'  tmprc.Open "Select CODALM FROM STOCK WHERE CODART = " & CLng(ioCODART.Text) & " AND TEMPOR = " & CLng(cbTEMPOR.Text) & " AND TALLA = " & CLng(cbCODTALLA.Text) & " AND COLOR = " & CLng(cbCODCOL.Text) & " AND STOCK > 0 ORDER BY CODALM", SrvCnn, adOpenDynamic, adLockReadOnly
'
'  'si esta vacio indicarlo
'  If tmprc.EOF Then lblstatus.Caption = "No se encuentra esa combinación de Artículo, Temporada, Talla y Color en ningún almacén"
'
 ' Do Until tmprc.EOF
'
'    miC.añade_item tmprc.Fields(0) & "   " & devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & tmprc.Fields(0))
'    tmprc.MoveNext
'
'  Loop
'
'   tmprc.Close
'  Set tmprc = Nothing
''

'  With SrvCnn
'    If .State <> 0 Then
'               .Close
'    End If
'  End With
'
'   On Error GoTo 0
'   Exit Sub
'
'carga_almacenes_origen_Error:'

  'With SrvCnn
'    If .State <> 0 Then
'               .Close
'    End If
'  End With'

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Carga_Almacenes_Origen de Formulario frmPtrans"
'End Sub





'---------------------------------------------------------------------------------------
' Procedimiento :ioUNIDADES_Validate
' Fecha/Hora    : 28/11/2003 10:21
' Autor         : JCastillo
' Propósito     : Validación para el campo unidades. Comprueba que existan
'                 unidades para ese articulo en ese almacen
'---------------------------------------------------------------------------------------
Private Sub ioUNIDADES_Validate(Cancel As Boolean)
    
   On Error GoTo ioUNIDADES_Validate_Error

    With ioUNIDADES
    
    'si es distinto de blanco
    If .Text <> "" Then
        
        'que sea mayor de 0
        If CDbl(.Valor) <= 0 Then
    
        lblstatus.Caption = "No se permiten 0 unidades"
        .SetFocus
        .CancelarValidacion
        Cancel = True
        Exit Sub
    
        Else
        
            If CDbl(.Valor) > 100 Then
                MsgBox "¡Atención!, revise el valor introducido en unidades, parece ser incorrecto", vbExclamation, titulo
                ioPREVEN.SetFocus
            End If
    
        End If
            
    Else

        lblstatus.Caption = "No se permite unidades en blanco"
        .SetFocus
        .CancelarValidacion
        Cancel = True
        Exit Sub
    
    End If
        
    
            
    End With
    
    'Call carga_almacenes_origen(cbCODALMORIG)

   On Error GoTo 0
   Exit Sub

ioUNIDADES_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioUNIDADES_Validate de Formulario frmPtrans"
    
End Sub



'---------------------------------------------------------------------------------------
' Subrutina     : comprueba_unidades
' Fecha/Hora    : 07/12/2003 11:42
' Autor         : JCASTILLO
' Propósito     : Validación para el campo unidades. Comprueba que existan
'                 unidades para ese articulo en ese almacen.
'                 Devuelve TRUE si existen unidades para ese artículo, y false
'                 si no existen unidades en stock
'---------------------------------------------------------------------------------------
Private Function comprueba_unidades() As Boolean
Dim tmpunis As Variant

  On Error GoTo comprueba_unidades_Error
    
    'comprobar si hay stock de ese artículo para ese almacén
        
        If mbEditFlag Then
        
        tmpunis = devuelve_campo("SELECT STOCK FROM STOCK WHERE CODART = " & rc.fields("CODART") & _
        " AND TALLA = " & rc.fields("CODTALLA") & " AND COLOR = " & rc.fields("CODCOL") & _
        " AND CODALM = " & AlmacenActual & " AND TEMPOR = " & rc.fields("TEMPOR"), locCnn)
     
        Else
        
        tmpunis = devuelve_campo("SELECT STOCK FROM STOCK WHERE CODART = " & Codigo_B.CODIGO_ART & _
        " AND TALLA = " & Codigo_B.TALLA_ART & " AND COLOR = " & Codigo_B.COLOR_ART & _
        " AND CODALM = " & AlmacenActual & " AND TEMPOR = " & Codigo_B.TEMPORADA_ART, locCnn)
       
        End If
       
        Select Case tmpunis
        Case "@"  'no existe el registro, por lo tanto stock = 0
                  '(no permitir)
           lblstatus.Caption = "No hay unidades en STOCK de ese artículo talla/color"
           '.SetFocus
           '.CancelarValidacion
           'Cancel = True
           comprueba_unidades = False
           Exit Function
            
        Case Is < CDbl(ioUNIDADES.Text) 'stock negativo o cero (no permitir)
            
            lblstatus.Caption = "No hay unidades en STOCK de ese artículo talla/color"
            '.SetFocus
            '.CancelarValidacion
            'Cancel = True
            comprueba_unidades = False
            Exit Function
        
        End Select
         
         
        comprueba_unidades = True
        lblstatus.Caption = ""
        Set tmpunis = Nothing
  
   On Error GoTo 0
   Exit Function

comprueba_unidades_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento comprueba_unidades de Formulario frmDetPtrans"
End Function

Private Sub añade_linea_grid(mi_Codart As Long, mi_Tempor As Byte, mi_CodTalla As Integer, mi_CodCol As Integer, mi_Uds As Single, mi_PVP As Single, mi_Dcto As Byte, mi_Iva As Single, mi_Linea As Long)
Dim tmpcodcolor As Long
Dim tmpsubt As Double
Dim linea As Long
Dim miArt As Variant

conta_lineas = conta_lineas + 1

With fgArt
         
        .Redraw = flexRDNone
       
        .Rows = .Rows + 1
        
        .TextMatrix(.Rows - 1, 1) = mi_Linea
        .TextMatrix(.Rows - 1, 2) = Format(mi_Codart, "00000") & " " & Trim(devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & mi_Codart & " AND TEMPOR = " & mi_Tempor, locCnn))
        .TextMatrix(.Rows - 1, 3) = Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & mi_CodTalla, locCnn))
        
       'obtener el texto del color y su codigo de color (para colorear
       'la celda del grid)
       If mi_CodCol > 0 Then
      
            tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & mi_CodCol)
            .TextMatrix(.Rows - 1, 4) = Trim(devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & mi_CodCol, locCnn))
            .Col = 4
            .Row = .Rows - 1
            .CellBackColor = tmpcodcolor
            .Col = 2
        
       End If
                
        'Unidades
        .TextMatrix(.Rows - 1, 5) = mi_Uds
        'PVP
        .TextMatrix(.Rows - 1, 6) = mi_PVP
        'DCTO
        .TextMatrix(.Rows - 1, 7) = mi_Dcto & " %"
                
        tmpsubt = (mi_Uds * mi_PVP) - (((mi_Uds * mi_PVP) * mi_Dcto) / 100)
        
        'importe
        .TextMatrix(.Rows - 1, 8) = tmpsubt
        'iva
        .TextMatrix(.Rows - 1, 9) = mi_Iva & " %"
        'total (importe + IVA)
        .TextMatrix(.Rows - 1, 10) = tmpsubt + ((tmpsubt * mi_Iva) / 100)
        
        
      '  .TextMatrix(.Rows - 1, 5) = devuelve_campo("SELECT AÑO + ' ' + TEMPORADA FROM TEMPOR WHERE IDTEM = " & rc.Fields("TEMPOR").Value)
      '  .TextMatrix(.Rows - 1, 6) = rc.Fields("FMODI").Value
 
    '.Enabled = True
    .Redraw = True
    End With
    

End Sub


'---------------------------------------------------------------------------------------
' Procedimiento : añade_arreglo_grid
' Fecha/Hora    : 28/04/2004 16:55
' Autor         : JCastillo
' Propósito     :
'---------------------------------------------------------------------------------------
Private Sub añade_arreglo_grid(mi_PVP As Single, mi_Linea As Long)
Dim tmpcodcolor As Long
Dim tmpsubt As Double
Dim linea As Long
Dim miArt As Variant

   On Error GoTo añade_arreglo_grid_Error

conta_lineas = conta_lineas + 1

With fgArt
         
        .Redraw = flexRDNone
       
        
        .Rows = .Rows + 1
        
        .TextMatrix(.Rows - 1, 1) = mi_Linea
        .TextMatrix(.Rows - 1, 2) = "Arreglo Varios"
        '.TextMatrix(.Rows - 1, 3) = Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & mi_CodTalla, locCnn))
        
       'obtener el texto del color y su codigo de color (para colorear
       'la celda del grid)
      ' If mi_CodCol > 0 Then
      
     '       tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & mi_CodCol)
     '       .TextMatrix(.Rows - 1, 4) = Trim(devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & mi_CodCol, locCnn))
     '       .Col = 4
     '       .Row = .Rows - 1
      '      .CellBackColor = tmpcodcolor
     '       .Col = 2
        
      ' End If
                
        'Unidades
       ' .TextMatrix(.Rows - 1, 5) = mi_Uds
        'PVP
        .TextMatrix(.Rows - 1, 6) = mi_PVP
        'DCTO
      '  .TextMatrix(.Rows - 1, 7) = mi_Dcto & " %"
                
    '    tmpsubt = (mi_Uds * mi_PVP) - (((mi_Uds * mi_PVP) * mi_Dcto) / 100)
      '
        'importe
        .TextMatrix(.Rows - 1, 8) = mi_PVP
        'iva
       ' .TextMatrix(.Rows - 1, 9) = mi_Iva & " %"
        'total (importe + IVA)
        .TextMatrix(.Rows - 1, 10) = mi_PVP
        
        .TextMatrix(.Rows - 1, 11) = 1
        
      '  .TextMatrix(.Rows - 1, 5) = devuelve_campo("SELECT AÑO + ' ' + TEMPORADA FROM TEMPOR WHERE IDTEM = " & rc.Fields("TEMPOR").Value)
      '  .TextMatrix(.Rows - 1, 6) = rc.Fields("FMODI").Value
             
    .SubtotalPosition = flexSTAbove
    .subtotal flexSTSum, , 5, , vbBlue, vbWhite
    .subtotal flexSTSum, , 8, , vbBlue, vbWhite
'    .subtotal flexSTSum, , 8, , vbBlue, vbWhite
    .subtotal flexSTSum, , 10, , vbBlue, vbWhite
    .TextMatrix(1, 4) = "Totales:"
    .TextMatrix(1, 2) = "Total (" & conta_lineas & ") Art."
    .AutoSize 1, .Cols - 1
      
    .Redraw = True
  
  
  DoEvents
  cargando_Grid = False
  
  'rc.Sort = ""
  DoEvents
  
 .HighLight = flexHighlightWithFocus
  .FocusRect = flexFocusHeavy
  .AllowBigSelection = False
  .AllowSelection = True
    End With
    

   On Error GoTo 0
   Exit Sub

añade_arreglo_grid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento añade_arreglo_grid de Formulario frmDetVenta"

End Sub


'---------------------------------------------------------------------------------------
' Subrutina   : cargar_grid
' Fecha/Hora  : 17/01/2004 17:31
' Autor       : JCASTILLO
' Propósito   : Carga los articulos en el grid
'---------------------------------------------------------------------------------------
Private Sub cargar_grid()
Dim act As Boolean
Dim rcarreglos As ADODB.Recordset

  On Error GoTo cargar_grid_Error

  cargando_Grid = True
  conta_lineas = 0

  'poner títulos al grid
  With fgArt
        .Clear
        .Rows = 1
        .Cols = 12
        .ColHidden(1) = True
        .ColHidden(11) = True   'tiene un 1 si es un arreglo
        .ColFormat(3) = ""
        .ColFormat(6) = "Currency"
        .ColFormat(8) = "Currency"
        .ColFormat(10) = "Currency"
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(9) = flexAlignRightCenter
        .TextMatrix(0, 2) = "Artículo"
        .TextMatrix(0, 3) = "Talla"
        .TextMatrix(0, 4) = "Color"
        .TextMatrix(0, 5) = "Unidades"
        .TextMatrix(0, 6) = "PVP"
        .TextMatrix(0, 7) = "DCTO"
        .TextMatrix(0, 8) = "Importe"
        .TextMatrix(0, 9) = "IVA"
        .TextMatrix(0, 10) = "Total"
  End With
  
  If rc.RecordCount <= 0 Then Exit Sub
    
  Set ioID.DataSource = Nothing
    
  'rc.Sort = "LINEA DESCENDING"
  DoEvents
  If Not rc.BOF Then rc.MoveFirst
  
  Do Until rc.EOF
  
    'añadir una linea al grid
    Call añade_linea_grid(rc.fields("CODART"), rc.fields("TEMPOR"), rc.fields("CODTALLA"), rc.fields("CODCOL"), rc.fields("UNIDADES"), rc.fields("PREVEN"), rc.fields("DCTO"), rc.fields("IVA"), rc.fields("LINEA"))
    DoEvents
    
   rc.MoveNext
  
  Loop
  
  'cargar los arreglos
  Set rcarreglos = New ADODB.Recordset
  rcarreglos.Open "SELECT ID, PVP FROM ARREGLOS WHERE CODVEN = " & CODIGO_VENTA & " AND CODCAJ = " & CajaActual, locCnn, adOpenStatic, adLockReadOnly
  
  Do Until rcarreglos.EOF
  
    'añadir una linea al grid
    Call añade_arreglo_grid(rcarreglos.fields("PVP"), rcarreglos.fields("ID"))
    DoEvents
    
    rcarreglos.MoveNext
  
  Loop
  
  
  rcarreglos.Close
  Set rcarreglos = Nothing
  
  rc.MoveFirst
  
  Set ioID.DataSource = rc
  
  With fgArt
       
    .SubtotalPosition = flexSTAbove
    .subtotal flexSTSum, , 5, , vbBlue, vbWhite
    .subtotal flexSTSum, , 8, , vbBlue, vbWhite
'    .subtotal flexSTSum, , 8, , vbBlue, vbWhite
    .subtotal flexSTSum, , 10, , vbBlue, vbWhite
    .TextMatrix(1, 4) = "Totales:"
    .TextMatrix(1, 2) = "Total (" & conta_lineas & ") Art."
    .AutoSize 1, .Cols - 1
      
    .Redraw = True
  End With
  
  DoEvents
  cargando_Grid = False
  
  'rc.Sort = ""
  DoEvents
  
  fgArt.HighLight = flexHighlightWithFocus
  fgArt.FocusRect = flexFocusHeavy
  fgArt.AllowBigSelection = False
  fgArt.AllowSelection = True
  
   On Error GoTo 0
   Exit Sub

cargar_grid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cargar_grid de Formulario frmDetVenta"

End Sub
