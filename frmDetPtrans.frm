VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmDetPtrans 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Peticiones de Transferencia"
   ClientHeight    =   7005
   ClientLeft      =   2400
   ClientTop       =   1920
   ClientWidth     =   10425
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   Begin PCGestion.chameleonButton cmdPrevious 
      Height          =   630
      Left            =   1080
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5505
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
      MICON           =   "frmDetPtrans.frx":0000
      PICN            =   "frmDetPtrans.frx":001C
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
      Top             =   2415
      Width           =   10380
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5505
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
      MICON           =   "frmDetPtrans.frx":0CEE
      PICN            =   "frmDetPtrans.frx":0D0A
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5505
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
      MICON           =   "frmDetPtrans.frx":1A40
      PICN            =   "frmDetPtrans.frx":1A5C
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5505
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
      MICON           =   "frmDetPtrans.frx":272E
      PICN            =   "frmDetPtrans.frx":274A
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6165
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
      MICON           =   "frmDetPtrans.frx":3480
      PICN            =   "frmDetPtrans.frx":349C
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
      TabIndex        =   7
      Top             =   6165
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
      MICON           =   "frmDetPtrans.frx":4176
      PICN            =   "frmDetPtrans.frx":4192
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6165
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
      MICON           =   "frmDetPtrans.frx":4A6C
      PICN            =   "frmDetPtrans.frx":4A88
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
      Left            =   7260
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6165
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
      MICON           =   "frmDetPtrans.frx":52E6
      PICN            =   "frmDetPtrans.frx":5302
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
      Left            =   8235
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6165
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
      MICON           =   "frmDetPtrans.frx":5BDC
      PICN            =   "frmDetPtrans.frx":5BF8
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
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6165
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
      MICON           =   "frmDetPtrans.frx":67CA
      PICN            =   "frmDetPtrans.frx":67E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.miCombo cbCODTALLA 
      Height          =   495
      Left            =   705
      TabIndex        =   3
      Top             =   1890
      Width           =   3270
      _ExtentX        =   7064
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
   Begin PCGestion.miCombo cbCODCOL 
      Height          =   480
      Left            =   4815
      TabIndex        =   4
      Top             =   1890
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   847
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
   Begin PCGestion.miText ioCODBAR 
      Height          =   525
      Left            =   420
      TabIndex        =   0
      Top             =   1395
      Width           =   3585
      _ExtentX        =   6535
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
   Begin PCGestion.miText ioCODART 
      Height          =   525
      Left            =   4830
      TabIndex        =   1
      Top             =   1395
      Width           =   1110
      _ExtentX        =   4233
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
   Begin PCGestion.miCombo cbTEMPOR 
      Height          =   480
      Left            =   7140
      TabIndex        =   2
      Top             =   1380
      Width           =   3240
      _ExtentX        =   7011
      _ExtentY        =   847
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
   Begin PCGestion.miText ioUNIDADES 
      Height          =   525
      Left            =   8490
      TabIndex        =   5
      Top             =   1890
      Width           =   750
      _ExtentX        =   1323
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
   Begin PCGestion.bsGradientLabel lblOrigen 
      Height          =   375
      Left            =   945
      Top             =   540
      Width           =   4200
      _ExtentX        =   7408
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
      Colour1         =   14737632
      Colour2         =   16761024
      Colour3         =   16761024
      Colour4         =   12632256
      CaptionAlignment=   1
   End
   Begin PCGestion.bsGradientLabel lblDestino 
      Height          =   375
      Left            =   6090
      Top             =   540
      Width           =   4275
      _ExtentX        =   7541
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
      Colour1         =   14737632
      Colour2         =   16761024
      Colour3         =   16761024
      Colour4         =   12632256
      CaptionAlignment=   1
   End
   Begin VSFlex8Ctl.VSFlexGrid vsModificaciones 
      Height          =   3345
      Left            =   2160
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2805
      Width           =   6105
      _cx             =   10769
      _cy             =   5900
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
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
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
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
      FormatString    =   $"frmDetPtrans.frx":74C0
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
   Begin PCGestion.chameleonButton cbQuitarArt 
      Height          =   375
      Left            =   4005
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   6570
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Quitar Artículo"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   16776960
      FCOLO           =   16776960
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmDetPtrans.frx":7565
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbIrACodigo 
      Height          =   360
      Left            =   4005
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   6180
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "Buscar Codigo ..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   11513775
      BCOLO           =   11513775
      FCOL            =   16776960
      FCOLO           =   16776960
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmDetPtrans.frx":7581
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.miText ioRE 
      Height          =   525
      Left            =   9660
      TabIndex        =   6
      Top             =   1890
      Width           =   750
      _ExtentX        =   1323
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
   Begin PCGestion.bsGradientLabel lblModelo 
      Height          =   375
      Left            =   0
      Top             =   975
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   661
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   15640462
      Colour2         =   12632256
      CaptionAlignment=   1
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RE"
      Height          =   360
      Left            =   9225
      TabIndex        =   34
      Top             =   1965
      Width           =   450
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      Height          =   330
      Left            =   15
      TabIndex        =   31
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
      TabIndex        =   30
      Top             =   60
      Width           =   1485
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DESTINO"
      Height          =   330
      Left            =   5190
      TabIndex        =   28
      Top             =   585
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ORIGEN"
      Height          =   330
      Left            =   15
      TabIndex        =   27
      Top             =   585
      Width           =   855
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "COLOR"
      Height          =   285
      Left            =   4035
      TabIndex        =   26
      Top             =   1980
      Width           =   735
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TALLA"
      Height          =   300
      Left            =   0
      TabIndex        =   25
      Top             =   1965
      Width           =   690
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Uds"
      Height          =   360
      Left            =   8010
      TabIndex        =   24
      Top             =   1965
      Width           =   510
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CB"
      Height          =   360
      Left            =   30
      TabIndex        =   23
      Top             =   1470
      Width           =   315
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      Height          =   360
      Left            =   4020
      TabIndex        =   22
      Top             =   1470
      Width           =   795
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TEMPORADA"
      Height          =   285
      Left            =   5910
      TabIndex        =   21
      Top             =   1470
      Width           =   1215
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
      TabIndex        =   11
      Top             =   60
      Width           =   2445
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
      Left            =   3705
      TabIndex        =   10
      Top             =   60
      Width           =   1440
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSFERENCIA"
      Height          =   330
      Left            =   1920
      TabIndex        =   9
      Top             =   90
      Width           =   1680
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ultima modificación"
      Height          =   315
      Left            =   5700
      TabIndex        =   8
      Top             =   90
      Width           =   2160
   End
End
Attribute VB_Name = "frmDetPtrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmDetPtrans
' DateTime  : 31/10/2003 10:08
' Author    : Administrador
' Purpose   : Detalle de peticiones de transferencia
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

Dim conta_lineas As Long

'codigo de la transferencia actual, para añadir nuevos tegistros
Public CODIGO_TRANSF As Long
 
'si entra como IR_A_ID = 0, añadir un nuevo registro. De lo contrario
'ir al registro con un find
Public IR_A_ID As Long

'junto con CODIGO_TRANSF es la clave unica del fichero
Public miCODALMORIG As Byte

Public ESTADO_TRANSF As Byte

Public NUMERO_PEDIDO As Long

'---------------------------------------------------------------------------------------
' Subrutina   : cbEliminar_Click
' Fecha/Hora  : 08/12/2003 12:07
' Autor       : JCASTILLO
' Propósito   : Eliminar la linea actual de transferencia
'---------------------------------------------------------------------------------------
Private Sub cbEliminar_Click()

   On Error GoTo cbEliminar_Click_Error

If ioID.Caption = "" Then Exit Sub

If MsgBox("¿Desea quitar el articulo:" & Chr(13) & ioCODART.Text & "-" & _
devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & CLng(ioCODART.Text) & " AND TEMPOR = " & CLng(cbTEMPOR.Text), locCnn) & " de la transferencia actual?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

locCnn.Execute "DELETE FROM DETTRANS WHERE ID = " & CLng(ioID.Caption)
DoEvents
rc.Requery

If rc.RecordCount = 0 Then
   Unload Me
Else
Call cmdFirst_Click
End If

   On Error GoTo 0
   Exit Sub

cbEliminar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbEliminar_Click de Formulario frmDetPtrans"

End Sub


'---------------------------------------------------------------------------------------
' Procedimiento : cbQuitarArt_Click
' Fecha/Hora    : 14/01/2004 13:40
' Autor         : JCastillo
' Propósito     :si la transferencia esta marcada como pendiente, devolver las unidades al stock
' cuando quites el artículo
'---------------------------------------------------------------------------------------
'
Private Sub cbQuitarArt_Click()
Dim miConn As New ADODB.Connection

   On Error GoTo cbQuitarArt_Click_Error


If mbEditFlag Or mbAddNewFlag Then Exit Sub

If rc.RecordCount <= 0 Or rc.EOF Then Exit Sub

If MsgBox("¿Esta seguro que desea quitar el artículo actual?", vbQuestion + vbYesNo, "Quitar Artículo") = vbNo Then Exit Sub

Select Case ESTADO_TRANSF

'si la transferencia esta como en creación, borrar el articulo sin mas
Case 0

    locCnn.Execute "DELETE FROM DETTRANS WHERE ID = " & rc.fields("ID") & " AND CODALM = " & rc.fields("CODALM")

'si la transferencia esta como PENDIENTE, borrar el artículo y devolver las unidades al
'stock del almacen de origen
Case 1
     
   With miConn
        .CursorLocation = adUseServer
        .Open strLocCnn
        .BeginTrans
   End With
     
   'restar unidades
   Call stock(rc.fields("CODART"), rc.fields("TEMPOR"), rc.fields("CODTALLA"), rc.fields("CODCOL"), rc.fields("CODALM"), rc.fields("UNIDADES"), False, miConn)

   'borrar registro
   miConn.Execute "DELETE FROM DETTRANS WHERE ID = " & rc.fields("ID") & " AND CODALM = " & rc.fields("CODALM")
   DoEvents
   
   With miConn
        .CommitTrans
        .Close
        Set miConn = Nothing
   End With
   
   DoEvents
   
   rc.Requery
      
   DoEvents
   
   Call cmdFirst_Click

End Select

   On Error GoTo 0
   Exit Sub
cbQuitarArt_Click_Error:

        If miConn.State = 1 Then miConn.Close
        Set miConn = Nothing

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbQuitarArt_Click de Formulario frmDetPtrans"

End Sub

Private Sub cbIrACodigo_Click()
Dim tmpcodigo As String * LenCodBar
Dim miCod As MiCodBar

   On Error GoTo cbIrACodigo_Click_Error

If mbEditFlag Or mbAddNewFlag Then Exit Sub

tmpcodigo = InputBox("Introduzca Código de Barras", "Ir a código ...")

If (Len(tmpcodigo) <> LenCodBar) Or (Not IsNumeric(tmpcodigo)) Then Exit Sub
''TERMINAR

    'descomponer el codigo de barras ...
    miCod = Descompone_CBAR(tmpcodigo)
    
    rc.MoveFirst
    
    Do
    'si coincide ...
    If (rc.fields("CODART") = miCod.CODIGO_ART) And (rc.fields("CODTALLA") = miCod.TALLA_ART) And (rc.fields("TEMPOR") = miCod.TEMPORADA_ART) And (rc.fields("CODCOL") = miCod.COLOR_ART) Then
    
          lblstatus.Caption = "Se ha encontrado el registro"
          Exit Do
          
    Else
                
        If Not rc.EOF Then
            rc.MoveNext
        Else
            lblstatus.Caption = "No se encontró el artículo"
        End If
            
    End If
    
    Loop Until rc.EOF

   On Error GoTo 0
   Exit Sub

cbIrACodigo_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbIrACodigo_Click de Formulario frmDetPtrans"

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
        rc.Find "ID = " & IR_A_ID
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
  
   Move (Screen.Width - Width) \ 2, Separacion_MDIForm
   
   With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
   End With
   
  Set rc = New Recordset
  oSQL.AddTable "DETTRANS"
  oSQL.AddSimpleWhereClause "CODIGO", CODIGO_TRANSF
  oSQL.AddSimpleWhereClause "CODALM", miCODALMORIG
  oSQL.AddOrderClause "CODIGO"

  rc.Open oSQL.SQL, locCnn, adOpenStatic, adLockOptimistic
''''''''
'////////////////////////////////////////////////////////
' Enlazar controles:
'////////////////////////////////////////////////////////

With ioID
    Set .DataSource = rc
        .DataField = "ID"
End With

With ioCODIGO
    Set .DataSource = rc
        .DataField = "CODIGO"
End With
  
With ioCODBAR
         .SoloNumeros = True
         .LongMaxima = LenCodBar
         .dspFormat = ""
         .intFormat = ""
         .Font.Name = "EAN13"
         .Font.Size = 13
End With

With cbCODTALLA
      .ConexionString = locCnn
      .LenCodigo = 5
      .SQLString = "SELECT CODIGO, DESCRIPCION FROM TALLAS WHERE (MBAJA = 0) AND (CODIGO > 0) ORDER BY CODIGO"
      .DataField = "CODTALLA"
      .carga
      Set .DataSource = rc
      .CodigoWidth = 800
End With
  
With cbCODCOL
      .ConexionString = locCnn
      .LenCodigo = 5
      .SQLString = "SELECT CODIGO, DESCRIPCION FROM COLORES WHERE (MBAJA = 0) AND (CODIGO > 0) ORDER BY CODIGO"
      .DataField = "CODCOL"
      .carga
      Set .DataSource = rc
      .CodigoWidth = 800
End With
  
  'no enlazar a datos
With cbTEMPOR
    .ConexionString = locCnn
    .SQLString = "SELECT IDTEM, AÑO + '  ' + TEMPORADA FROM TEMPOR WHERE MBAJA = 0 ORDER BY ACTUAL DESC, IDTEM"
    .LenCodigo = 3
    .CodigoWidth = 500
    .DataField = "TEMPOR"
    .carga
    Set .DataSource = rc
End With

With ioCODART
         .Alineacion = 1
         .DataField = "CODART"
    Set .DataSource = rc
         .LongMaxima = 5
         .SoloNumeros = True
         .PermitirBlanco = True
End With

With ioUNIDADES
         .Alineacion = 1
         .dspFormat = ""
         .LongMaxima = 9
         .SoloNumeros = True
         .PermitirBlanco = False
End With
 
  
With ioFMODI
    Set .DataSource = rc
        .DataField = "FMODI"
End With
  
  'poner títulos al grid
  With vsModificaciones
        .TextMatrix(0, 1) = "Artículo"
        .TextMatrix(0, 2) = "Talla"
        .TextMatrix(0, 3) = "Color"
        .TextMatrix(0, 4) = "Unidades"
        .TextMatrix(0, 5) = "Temporada"
        .TextMatrix(0, 6) = "Fecha"
  End With
  
  With ioRE
    .SoloNumeros = True
    .LongMaxima = 3
    .Alineacion = 1
    Set .DataSource = rc
    .DataField = "RE"
  End With
  
  If TipoPermiso = 0 Then
    ioRE.Visible = False
  'solo el supervisor puede ver o insertar RE
  Else
    ioRE.Visible = True
  End If
  
       
  vsModificaciones.ColAlignment(2) = flexAlignLeftCenter
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

      Case vbKeyF5
            Call cmdFirst_Click
    
       Case vbKeyF6
            Call cmdPrevious_Click
      
       Case vbKeyF7
            Call cmdNext_Click
    
       Case vbKeyF8
        Call cmdLast_Click
      
      
  End Select
  KeyCode = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dim tmpcodigo As Double
Dim tmpalmorig As Byte

On Error Resume Next
tmpcodigo = rc.fields("CODIGO")
tmpalmorig = rc.fields("CODALM")

If rc.EditMode <> adEditNone Then rc.CancelUpdate

If rc.State = 1 Then rc.Close
Set rc = Nothing

DoEvents

With frmPtrans
     .WindowState = vbNormal
Call .carga_grid_detalle(tmpcodigo, tmpalmorig)
End With

Set oSQL = Nothing
Set frmDetPtrans = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub




'Private Sub ioCODART_GotFocus()
'
'si mete codigo de barras y pasa al campo codigo, que agrege
'automaticamente el registro y meta otro
'If ioCODBAR.Text = "" Then Exit Sub

'If ioCODART.Text <> "" And cbTEMPOR.Text <> "" And _
'cbCODTALLA.Text <> "" And cbCODCOL.Text <> "" And ioUNIDADES.Text <> "" Then
 '
 '
 '
 '           'comprobar si existe el artículo/temporada
'        If devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & CLng(Left(.Text, 5)) & " AND TEMPOR = " & CLng(Mid(.Text, 6, 3))) = "@" Then
 '
  '              lblstatus.Caption = "No existe el artículo para esa temporada!, Codigo de Barras no Válido"
   '             ioCODBAR.Text = ""
    '            .CancelarValidacion
     '           Cancel = True
'                Beep
 '               Call Espera(1)
  '              Beep
   '             Call Espera(1)
    '            Beep
    '            Exit Sub
'
 ''       End If
    
   '
'    Call cbactualizar_Click
    'Call cbAgregar_Click
'End If
'End Sub



Private Sub ioRE_GotFocus()

If Not mbAddNewFlag Or mbEditFlag Then Exit Sub

If mbEditFlag Then _
    If ioRE.Text = "0" Then ioRE.Text = ""

End Sub

Private Sub rc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  If rc.AbsolutePosition > 0 Then
  
  If rc.fields("CODART").Value > 0 Then
    lblModelo.Caption = Format(rc.fields("CODART"), "00000") & "  " & Trim(devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & rc.fields("CODART") & " AND TEMPOR = " & rc.fields("TEMPOR"), locCnn)) & "  " & Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & rc.fields("CODTALLA"), locCnn)) & "  " & Trim(devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & rc.fields("CODCOL"), locCnn))
  Else
    lblModelo.Caption = ""
  End If
  
  lblstatus.Caption = "Registro: " & CStr(rc.AbsolutePosition)
   ioUNIDADES.Text = rc.fields("UNIDADES").Value
   
   
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
    
    tmpcodigo = devuelve_campo("select max(ID) + 1 from DETTRANS where CODALM = " & AlmacenActual)
    
    If tmpcodigo = "@" Then tmpcodigo = 1
    'Si devuelve @ esque ha habido un error
    'If tmpcodigo <> "@" Then
    
    .fields("ID") = CLng(tmpcodigo)
    .fields("CODIGO") = CODIGO_TRANSF
    .fields("CODALM") = miCODALMORIG
    
    Set tmpcodigo = Nothing
    
    'End If

    'Set tmpcodigo = Nothing
    
    lblstatus.Caption = "Agregar registro"
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
  
  If Me.ESTADO_TRANSF = 1 Then
  
    lblstatus.Caption = "No se puede modificar, el estado es PENDIENTE (puede AÑADIR o BORRAR)"
    Exit Sub
  
  End If

  lblstatus.Caption = "Modificar registro"
  mbEditFlag = True
  SetButtons False
  cbActualizar.Visible = True
    
  If TipoPermiso = 1 Then
    ioRE.SetFocus
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
        
        'comprobar solo si esta añadiendo
        If mbAddNewFlag Then
        'si no hay unidades para ese artículo
        If comprueba_unidades = False Then
           If MsgBox("No existen unidades para ese artículo en " & lblORIGEN.Caption & "." & Chr(13) & "¿Continuar?", _
           vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        
        End If
    
    Else

        lblstatus.Caption = "No se permite unidades en blanco"
        .SetFocus
        .CancelarValidacion
        Exit Sub
    
    End If
    
  End With
    
  If ioCODART.Text = "" Then
    lblstatus.Caption = "El código de artículo no puede estar en blanco"
    ioCODART.SetFocus
    DoEvents
    ioCODART.CancelarValidacion
  Exit Sub
  End If
    
  'solo si esta visible ...
  If ioRE.Visible Then
  
    If ioRE.Text = "" Then
        ioRE.Text = "0"
    End If
  
    If Not IsNumeric(ioRE.Text) Then
        ioRE.Text = "0"
    End If
  End If
  
  If cbTEMPOR.Text = "" Then
    lblstatus.Caption = "Temporada no puede estar en blanco"
    cbTEMPOR.SetFocus
    Exit Sub
  End If
     
  If cbCODTALLA.Text = "" Then
    lblstatus.Caption = "Talla no puede estar en blanco"
    cbCODTALLA.SetFocus
    Exit Sub
  End If
  
  If cbCODCOL.Text = "" Then
    lblstatus.Caption = "Color no puede estar en blanco"
    cbCODCOL.SetFocus
    Exit Sub
  End If
  
   rc.fields("UNIDADES") = CDbl(ioUNIDADES.Valor)
 
  rc.UpdateBatch adAffectAll

  SetButtons True
    
  DoEvents
  'añade el articulo actual al grid de modificaciones
  Call añade_linea_grid
  
  ioCODBAR.Text = ""
  
  'ioCODART.Text = ""
  'cbTEMPOR.Text = ""
  'ioUNIDADES.Text = "0"
  'cbCODTALLA.Text = ""
  'cbCODCOL.Text = ""

  
  lblstatus.Caption = ""
  
  DoEvents

  'si esta editando, pasar al siguiente registro y editarlo
  If mbEditFlag Then
  
  mbEditFlag = False
  mbDataChanged = False
    
  Call cmdNext_Click
  Call cbedicion_Click
  'DoEvents
  'ioRE.SetFocus
    
  Else
  
  If mbAddNewFlag Then
    rc.MoveLast              'va al nuevo registro
  End If
  
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

Call cbCancelar_Click

On Error Resume Next
Unload Me
'Set frmDetPtrans = Nothing
'If rc.State = 1 Then rc.Close
'DoEvents

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
  
  cbCODTALLA.Enabled = Not bVal
  cbCODCOL.Enabled = Not bVal
  cbTEMPOR.Enabled = Not bVal
  
End Sub


'---------------------------------------------------------------------------------------
' Subrutina   : ioCODART_Validate
' Fecha/Hora  : 25/11/2003 20:56
' Autor       : JCASTILLO
' Propósito   :  Validación de código de artículo y temporada
'---------------------------------------------------------------------------------------
Private Sub ioCODART_Validate(Cancel As Boolean)
  
 On Error GoTo ioCODART_Validate_Error

 If ioCODART.Text <> "" And cbTEMPOR.Text <> "" Then
 
    If devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & CLng(ioCODART.Text) & " AND TEMPOR = " & CLng(cbTEMPOR.Text)) = "@" Then
        
                lblstatus.Caption = "No existe el artículo para esa temporada!"
                ioCODART.Text = ""
                ioCODART.CancelarValidacion
                Cancel = True
                Exit Sub
    Else
    
        lblstatus.Caption = ""
        'Call carga_almacenes_origen(cbCODALMORIG)
                
    End If
                
 End If
 

   On Error GoTo 0
   Exit Sub

ioCODART_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioCODART_Validate de Formulario frmPtrans"

End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : ioCODBAR_validate
' Fecha/Hora  : 07/12/2003 11:01
' Autor       : JCASTILLO
' Propósito   : Descompone el codigo de barras y introduce la información en los campos
'               correspondientes
'---------------------------------------------------------------------------------------
Private Sub ioCODBAR_Validate(Cancel As Boolean)
Dim miCod As MiCodBar

   On Error GoTo ioCODBAR_Validate_Error

With ioCODBAR

If Trim(.Text) = "" Then Exit Sub

    'si es un codigo de barras con la longitud válidad
    If Len(Trim(.Text)) = LenCodBar Then
    
        miCod = Descompone_CBAR(ioCODBAR.Text)
                
If NUMERO_PEDIDO = 0 Then

            ' comprobar si existe el artículo/temporada
            If devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & CLng(Left(.Text, 5)) & " AND TEMPOR = " & CLng(Mid(.Text, 6, 3))) = "@" Then
        
                MsgBox "¡No existe este artículo en la base de datos!", vbExclamation, titulo
                .Text = ""
                .SetFocus
                .CancelarValidacion
                Cancel = True
                Beep
                Call Espera(1)
                Beep
                Call Espera(1)
                Beep
                Exit Sub
           
            End If
            

'si tiene pedido comprobar en los artículos del pedido
Else
        
'
' Valores:
'                           1  Articulo no existe en el pedido (tmpuped unidades 0)
'                           2  Articulo existe, pero ya no admite mas unidades (tmpuds = tmpuped)
'                           3  Articulo correcto para este pedido
'
        Select Case comprueba_unidades_pedido(NUMERO_PEDIDO, CInt(miCod.CODIGO_ART), CInt(miCod.TEMPORADA_ART), CInt(miCod.TALLA_ART), CInt(miCod.COLOR_ART))
        
        ' 1  Articulo no existe en el pedido (tmpuped unidades 0)
        Case 1
                
                MsgBox "No existe el artículo en el pedido: " & NUMERO_PEDIDO, vbExclamation, titulo
                .Text = ""
                .SetFocus
                .CancelarValidacion
                Cancel = True
                Beep
                Call Espera(1)
                Beep
                Call Espera(1)
                Beep
                Exit Sub
                
       '2  Articulo existe, pero ya no admite mas unidades (tmpuds = tmpuped)
        Case 2
        
                MsgBox "El artículo no admite mas unidades (sobran unidades)", vbExclamation, titulo
                .Text = ""
                .SetFocus
                .CancelarValidacion
                Cancel = True
                Beep
                Call Espera(1)
                Beep
                Call Espera(1)
                Beep
                Exit Sub
        
        '3  Articulo correcto para este pedido
        Case 3
        
                lblstatus.Caption = "Artículo OK"
                
        End Select
        
End If
        
        'un pitido para todo OK
        Beep
                
       
        'codigo de artículo
        'ioCODART.Text = CLng(Left(.Text, 5))
        
        'temporada
        'cbTEMPOR.Text = CLng(Mid(.Text, 6, 3))
        
        
        'cbTEMPOR.Text = CLng(Mid(.Text, 6, 3))
                
        'talla
        'cbCODTALLA.Text = CLng(Mid(.Text, 9, 2))
        
        'cbCODTALLA.Text = CLng(Mid(.Text, 9, 2))
        
        'color
        'cbCODCOL.Text = CLng(Mid(.Text, 11, 3))
        
        ioCODART.Text = miCod.CODIGO_ART
        
        cbTEMPOR.Text = miCod.TEMPORADA_ART
        
        cbCODTALLA.Text = miCod.TALLA_ART
        
        cbCODCOL.Text = miCod.COLOR_ART
        
        ioUNIDADES.Text = "1"
        
        Call cbactualizar_Click
        
        'para que no salga el foco
        Cancel = True
        

    Else

        lblstatus.Caption = "Código de Barras no válido"
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
    
      '  DoEvents
       ' ioUNIDADES.SetFocus

End With

   On Error GoTo 0
   Exit Sub

ioCODBAR_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioCODBAR_validate de Formulario frmDetPtrans"

End Sub


Private Sub cbCODCOL_Validate(Cancel As Boolean)
    'Call carga_almacenes_origen(cbCODALMORIG)
End Sub

Private Sub cbCODTALLA_Validate(Cancel As Boolean)
  '  Call carga_almacenes_origen(cbCODALMORIG)
End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : cbTEMPOR_Validate
' Fecha/Hora  : 25/11/2003 20:56
' Autor       : JCASTILLO
' Propósito   : Validación de código de artículo y temporada
'
'---------------------------------------------------------------------------------------
Private Sub cbTEMPOR_Validate(Cancel As Boolean)
   On Error GoTo cbTEMPOR_Validate_Error

 If ioCODART.Text <> "" And cbTEMPOR.Text <> "" Then
 
    If devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & CLng(ioCODART.Text) & " AND TEMPOR = " & CLng(cbTEMPOR.Text), locCnn) = "@" Then
        
                lblstatus.Caption = "No existe el artículo para esa temporada!, Codigo de Barras no Válido"
                cbTEMPOR.Text = ""
                ioCODART.Text = ""
                '.CancelarValidacion
                Cancel = True
                Beep
                Call Espera(1)
                Beep
                Call Espera(1)
                Beep
                Exit Sub
                
    Else
    
        lblstatus.Caption = ""
        'Call carga_almacenes_origen(cbCODALMORIG)
                
    End If
 
 End If

   On Error GoTo 0
   Exit Sub

cbTEMPOR_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbTEMPOR_Validate de Formulario frmPtrans"
End Sub



Private Sub ioCODIGO_Change()
ioCODIGO.Caption = Format(ioCODIGO.Caption, "00000")
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

 If Not mbAddNewFlag Or mbEditFlag Then Exit Sub
 

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
    If ioCODART.Text <> "" And cbCODTALLA.Text <> "" And _
    cbCODCOL.Text <> "" And cbTEMPOR.Text <> "" And frmPtrans.lblAlmOrig.Caption <> "" Then
        
        tmpunis = devuelve_campo("SELECT STOCK FROM STOCK WHERE CODART = " & ioCODART.Text & _
        " AND TALLA = " & cbCODTALLA.Text & " AND COLOR = " & cbCODCOL.Text & _
        " AND CODALM = " & AlmacenActual & " AND TEMPOR = " & cbTEMPOR.Text, locCnn)
       
        Select Case tmpunis
        Case "@"  'no existe el registro, por lo tanto stock = 0
                  '(no permitir)
           lblstatus.Caption = "No hay unidades en STOCK de ese artículo talla/color en " & lblORIGEN.Caption
           '.SetFocus
           '.CancelarValidacion
           'Cancel = True
           comprueba_unidades = False
           Exit Function
            
        Case Is < CDbl(ioUNIDADES.Text) 'stock negativo o cero (no permitir)
            
            lblstatus.Caption = "No hay unidades en STOCK de ese artículo talla/color en " & lblORIGEN.Caption
            '.SetFocus
            '.CancelarValidacion
            'Cancel = True
            comprueba_unidades = False
            Exit Function
        
        End Select
         
         
        comprueba_unidades = True
        lblstatus.Caption = ""
        Set tmpunis = Nothing
    End If

   On Error GoTo 0
   Exit Function

comprueba_unidades_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento comprueba_unidades de Formulario frmDetPtrans"
End Function



Private Sub añade_linea_grid()
Dim tmpcodcolor As Long
Dim linea As Long

conta_lineas = conta_lineas + 1

With vsModificaciones
        
        .Redraw = flexRDNone
       ' .Clear
        
                
        'poner títulos
        .TextMatrix(0, 1) = "Artículo"
        .TextMatrix(0, 2) = "Talla"
        .TextMatrix(0, 3) = "Color"
        .TextMatrix(0, 4) = "Unidades"
        .TextMatrix(0, 5) = "Temporada"
        .TextMatrix(0, 6) = "Fecha"
        
        .Rows = .Rows + 1
                        
        'numero de linea
        '.TextMatrix(.Rows - 1, 1) = tmprc.Fields("IDTRANS").Value
        .TextMatrix(.Rows - 1, 1) = Format(rc.fields("CODART"), "00000") & " " & devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & rc.fields("CODART").Value & " AND TEMPOR = " & rc.fields("TEMPOR"))
        .TextMatrix(.Rows - 1, 2) = Format(rc.fields("CODTALLA").Value, "000") & " " & devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & rc.fields("CODTALLA").Value)
        
       'obtener el texto del color y su codigo de color (para colorear
       'la celda del grid)
       If Not IsNull(rc.fields("CODCOL")) And rc.fields("CODCOL") <> 0 Then
      
            tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & rc.fields("CODCOL"))
            .TextMatrix(.Rows - 1, 3) = devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & rc.fields("CODCOL"))
            .Col = 3
            .Row = .Rows - 1
            .CellBackColor = tmpcodcolor
            .Col = 2
        
       End If
                
        .TextMatrix(.Rows - 1, 4) = rc.fields("UNIDADES").Value
        .TextMatrix(.Rows - 1, 5) = devuelve_campo("SELECT AÑO + ' ' + TEMPORADA FROM TEMPOR WHERE IDTEM = " & rc.fields("TEMPOR").Value)
                   
       
'        .TextMatrix(.Rows - 1, 6) = rc.Fields("FMODI").Value
    
    .SubtotalPosition = flexSTAbove
    .subtotal flexSTSum, , 4, , vbBlue, vbWhite
    .TextMatrix(1, 3) = "Total Uds:"
    .TextMatrix(1, 1) = "Total (" & conta_lineas & ") Art."
    .AutoSize 1, .Cols - 1
    .Redraw = True
    '.Enabled = True
    End With
    

End Sub


'---------------------------------------------------------------------------------------
' Procedimiento : comprueba_unidades_pedido
' Fecha/Hora     : 10/03/2004 09:15
' Autor             : JCastillo
' Propósito       : Comprobar si existen unidades suficientes en el pedido antes de dejar paso. (Comprueba sumando todas las uds. introducidas hasta
'                       ahora en las transferencias "EN CREACION" para este pedido.
'                       devuelve:
'                           1  Articulo no existe en el pedido (tmpuped unidades 0)
'                           2  Articulo existe, pero ya no admite mas unidades (tmpuds = tmpuped)
'                           3  Articulo correcto para este pedido
'---------------------------------------------------------------------------------------
Private Function comprueba_unidades_pedido(gNumped As Long, gCodart As Integer, gTempor As Byte, gCodtalla As Integer, gCodCol As Integer) As Byte
Dim tmpuds As Variant
Dim tmpUped As Variant

   On Error GoTo comprueba_unidades_pedido_Error

    'sacar las unidades introducidas en todas las transferencias relacionadas por número de pedido
    'del artículo/talla/color actual
    tmpuds = devuelve_campo("SELECT Sum(UNIDADES) from DETTRANS WHERE (CAST(CODIGO AS VARCHAR(10)) + CAST(CODALM AS VARCHAR(3))) IN (SELECT CAST(CODIGO AS VARCHAR(10)) + CAST(CODALMORIG AS VARCHAR(3)) FROM PTRANS WHERE ESTADO <> 3 AND NUMPED = " & gNumped & " and CODALMORIG = " & AlmacenActual & ") AND CODART = " & gCodart & " AND TEMPOR = " & gTempor & " AND CODTALLA = " & gCodtalla & " AND CODCOL = " & gCodCol, locCnn)
    
    If tmpuds = "@" Then tmpuds = 0
    
    'sacar las unidades existentes en el pedido para el artículo actual
    tmpUped = devuelve_campo("SELECT Sum(UNIDADES) from DETPEDPRO WHERE NUMERO = " & gNumped & " AND ALMORIG = " & AlmacenActual & " AND CODART = " & gCodart & " AND TEMPOR = " & gTempor & " AND CODTALLA = " & gCodtalla & " AND CODCOL = " & gCodCol, locCnn)
 
    If tmpUped = "@" Then tmpUped = 0
    
    If tmpUped = 0 Then
        'no existen unidades
        comprueba_unidades_pedido = 1
        Exit Function
    End If
    
    If tmpuds >= tmpUped Then
        'existen unidades, pero el articulo ya no admite mas
        comprueba_unidades_pedido = 2
        Exit Function
    End If
    
    If tmpuds < tmpUped Then
        'el artículo puede entrar correctamente ...
        comprueba_unidades_pedido = 3
        Exit Function
    End If
       
    
   On Error GoTo 0
   Exit Function

comprueba_unidades_pedido_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento comprueba_unidades_pedido de Formulario frmDetPtrans"
End Function
