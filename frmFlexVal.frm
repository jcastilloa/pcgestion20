VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmFlexVal 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vales ..."
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11460
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   11460
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   5400
      Left            =   15
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1500
      Visible         =   0   'False
      Width           =   11430
      _cx             =   20161
      _cy             =   9525
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   14331047
      ForeColor       =   -2147483640
      BackColorFixed  =   15120763
      ForeColorFixed  =   -2147483630
      BackColorSel    =   14859077
      ForeColorSel    =   -2147483635
      BackColorBkg    =   -2147483636
      BackColorAlternate=   15573900
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFlexVal.frx":0000
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
      DataMode        =   3
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
   Begin PCGestion.bsGradientLabel bsGradientLabel4 
      Height          =   300
      Left            =   6420
      Top             =   1185
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   529
      Caption         =   "-N- Nu. Cliente  -C- Asig. Cliente  -F5- Ir a Rejilla  -F8- Salir"
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   15640462
      Colour2         =   7177785
      CaptionAlignment=   1
   End
   Begin PCGestion.bsGradientLabel lblCliente 
      Height          =   300
      Left            =   1485
      Top             =   1185
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   529
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   11311500
      Colour2         =   15640462
      CaptionAlignment=   1
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   1455
      Left            =   30
      TabIndex        =   9
      Top             =   30
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   2566
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Hoja 1"
      TabPicture(0)   =   "frmFlexVal.frx":00DE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label15"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cbBorrar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cbLista"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ioIMPORTE"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ioFECHAFIN"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ioFECHA"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ioCODIGO"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "ioTICKET"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cbCAJAS"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Hoja 2"
      TabPicture(1)   =   "frmFlexVal.frx":00FA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cbESTADO"
      Tab(1).Control(1)=   "cbTIPO"
      Tab(1).Control(2)=   "cbImprimir"
      Tab(1).Control(3)=   "chameleonButton2"
      Tab(1).Control(4)=   "Label8"
      Tab(1).Control(5)=   "Label6"
      Tab(1).ControlCount=   6
      Begin PCGestion.miCombo cbCODTALLA 
         Height          =   495
         Left            =   -74220
         TabIndex        =   10
         Top             =   30
         Width           =   2625
         _ExtentX        =   4630
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
         Height          =   465
         Left            =   -70905
         TabIndex        =   11
         Top             =   30
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   820
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
      Begin PCGestion.miCombo cbTIPOAB 
         Height          =   495
         Left            =   -74220
         TabIndex        =   12
         Top             =   480
         Width           =   2625
         _ExtentX        =   4630
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
      Begin PCGestion.miCombo cbCATTALL 
         Height          =   495
         Left            =   -68910
         TabIndex        =   13
         Top             =   525
         Width           =   4155
         _ExtentX        =   5821
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin PCGestion.miCombo cbFAMILIA 
         Height          =   480
         Left            =   -73995
         TabIndex        =   14
         Top             =   45
         Width           =   4155
         _ExtentX        =   7329
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
      Begin PCGestion.chameleonButton chameleonButton1 
         Height          =   450
         Left            =   -69990
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   794
         BTYPE           =   9
         TX              =   "&Imprimir Vale"
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
         MICON           =   "frmFlexVal.frx":0116
         PICN            =   "frmFlexVal.frx":0132
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PCGestion.miCombo cbCAJAS 
         Height          =   540
         Left            =   6585
         TabIndex        =   2
         Top             =   75
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   953
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
      Begin PCGestion.miText ioTICKET 
         Height          =   480
         Left            =   1125
         TabIndex        =   0
         Top             =   60
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   820
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
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
      Begin PCGestion.miText ioCODIGO 
         Height          =   465
         Left            =   4200
         TabIndex        =   1
         Top             =   60
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   820
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
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
      Begin PCGestion.miText ioFECHA 
         Height          =   495
         Left            =   1320
         TabIndex        =   3
         Top             =   555
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
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
      Begin PCGestion.miText ioFECHAFIN 
         Height          =   495
         Left            =   4185
         TabIndex        =   4
         Top             =   555
         Width           =   1380
         _ExtentX        =   2434
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
         dspFormat       =   ""
         Enabled         =   -1  'True
         EsPassword      =   -1  'True
      End
      Begin PCGestion.miText ioIMPORTE 
         Height          =   465
         Left            =   6600
         TabIndex        =   5
         Top             =   555
         Width           =   1275
         _ExtentX        =   2223
         _ExtentY        =   820
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
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
      Begin PCGestion.miCombo cbESTADO 
         Height          =   510
         Left            =   -73860
         TabIndex        =   6
         Top             =   150
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   900
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
      Begin PCGestion.miCombo cbTIPO 
         Height          =   480
         Left            =   -69840
         TabIndex        =   7
         Top             =   165
         Width           =   3210
         _ExtentX        =   5662
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
      Begin PCGestion.ucGrdBttn cbLista 
         Height          =   405
         Left            =   9015
         TabIndex        =   33
         Top             =   615
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   714
         Caption         =   "&Consultar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Image           =   "frmFlexVal.frx":0E0C
      End
      Begin PCGestion.ucGrdBttn cbBorrar 
         Height          =   405
         Left            =   10305
         TabIndex        =   34
         Top             =   615
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   714
         Caption         =   "&Borrar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Image           =   "frmFlexVal.frx":0E28
      End
      Begin PCGestion.chameleonButton cbImprimir 
         Height          =   525
         Left            =   -66510
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   45
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   926
         BTYPE           =   9
         TX              =   "Imprimir Vale"
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
         MICON           =   "frmFlexVal.frx":0E44
         PICN            =   "frmFlexVal.frx":0E60
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PCGestion.chameleonButton chameleonButton2 
         Height          =   450
         Left            =   -66480
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   600
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   794
         BTYPE           =   9
         TX              =   "Imprimir Rejilla"
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
         MICON           =   "frmFlexVal.frx":1B3A
         PICN            =   "frmFlexVal.frx":1B56
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO"
         Height          =   330
         Left            =   -70395
         TabIndex        =   32
         Top             =   255
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO"
         Height          =   330
         Left            =   -74730
         TabIndex        =   31
         Top             =   225
         Width           =   810
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IMPORTE"
         Height          =   330
         Left            =   5670
         TabIndex        =   30
         Top             =   630
         Width           =   885
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA FIN."
         Height          =   330
         Left            =   2970
         TabIndex        =   29
         Top             =   630
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA INI."
         Height          =   330
         Left            =   180
         TabIndex        =   28
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TICKET"
         Height          =   285
         Left            =   135
         TabIndex        =   27
         Top             =   150
         Width           =   765
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CAJA"
         Height          =   330
         Left            =   5985
         TabIndex        =   26
         Top             =   135
         Width           =   540
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CODIGO"
         Height          =   285
         Left            =   3345
         TabIndex        =   25
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FAMILIA"
         Height          =   315
         Left            =   -74985
         TabIndex        =   24
         Top             =   120
         Width           =   960
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CAT. TALLA"
         Height          =   330
         Left            =   -70245
         TabIndex        =   23
         Top             =   615
         Width           =   1305
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "EN"
         Height          =   300
         Left            =   -74670
         TabIndex        =   22
         Top             =   540
         Width           =   450
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO"
         Height          =   345
         Left            =   -67290
         TabIndex        =   21
         Top             =   120
         Width           =   810
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "COLOR"
         Height          =   285
         Left            =   -71610
         TabIndex        =   20
         Top             =   105
         Width           =   735
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TALLA"
         Height          =   300
         Left            =   -74925
         TabIndex        =   19
         Top             =   120
         Width           =   690
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F.INICIAL"
         Height          =   285
         Left            =   -69960
         TabIndex        =   18
         Top             =   585
         Width           =   1035
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "MOTIVO"
         Height          =   285
         Left            =   -74355
         TabIndex        =   17
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA FINAL"
         Height          =   330
         Left            =   -71835
         TabIndex        =   16
         Top             =   180
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmFlexVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Modulo      : frmFlexVal
' Fecha/Hora  : 04/02/2004 22:41
' Autor       : JCASTILLO
' Propósito   : Ver/Seleccionar Vale
'---------------------------------------------------------------------------------------

'**********************************
'CODIGO DE TICKET:
' 9 digitos codigo de venta
' 3 digitos codigo de caja
'**********************************

Option Explicit

Dim first As Boolean

Dim tmpstrcombo As String

Public desde_mnt As Boolean

Public miosql As New clsSmartSQL
'Public masql As New clsSmartSQL
'Public artSQL As New clsSmartSQL

Public miRc As New ADODB.Recordset
Public seleccionado As Boolean

'Dim nif As New clsNIF

Private Sub cbBorrar_click()

ioTICKET.Text = ""
ioCODIGO.Text = ""
cbCAJAS.Text = CajaActual
ioFECHA.Text = Date
ioIMPORTE.Text = ""
cbESTADO.Text = "0"  'PENDIENTE
cbTIPO.Text = ""

fg.Clear
fg.Rows = 1
'Call cbLista_click

End Sub

Private Sub cbCAJAS_Validate(Cancel As Boolean)
Call cbLista_click
End Sub


Private Sub cbestado_GotFocus()

If Tab1.Tab = 0 Then Tab1.Tab = 1

End Sub

Private Sub cbImprimir_Click()

If fg.Rows <= 1 Then Exit Sub

'imprimir el vale seleccionado
If fg.TextMatrix(fg.Row, 1) <> "" And fg.TextMatrix(fg.Row, 2) <> "" Then
    
    'ver que no este ya aceptado
    If fg.TextMatrix(fg.Row, 10) <> "ACEPTADO" Then
        Call Imprime_Vale(fg.TextMatrix(fg.Row, 2), fg.TextMatrix(fg.Row, 1), locCnn)
    Else
        MsgBox "El vale ya ha sido recibido y aceptado. Imposible imprimir.", vbExclamation, titulo
    End If

End If

End Sub

Private Sub cbLista_click()
Dim usa_where As Boolean
Dim nuefech  As String


   On Error GoTo cbLista_click_Error

miosql.ClearWhereClause
'masql.ClearWhereClause
'artSQL.ClearWhereClause

If (ioFECHA.Text <> "") And IsDate(ioFECHA.Text) Then

    'miosql.AddSimpleWhereClause "FALTA", ioFECHA.Text, , CLAUSE_GREATERTHANOREQUAL
    'miosql.AddSimpleWhereClause "FALTA", CStr(DateAdd("d", 1, ioFECHA.Text)), , CLAUSE_LESSTHAN, LOGIC_AND
    'miOsql.AddComplexWhereClause "Year(FALTA IN (" & masql.SQL & ")", LOGIC_AND
    
    
    
    If Trim(ioFECHAFIN.Text) <> "" And IsDate(ioFECHAFIN.Text) Then
     
        nuefech = ioFECHAFIN.Text
    
    Else
    
        '>= q el dia actual
        '< que el dia siguiente
        nuefech = DateAdd("d", 1, ioFECHA.Text)
    
    End If
    
    miosql.AddComplexWhereClause "(FMODI >= '" & Format(Year((ioFECHA.Text)), "0000") & Format(Month((ioFECHA.Text)), "00") & Format(Day((ioFECHA.Text)), "00") & "' AND FMODI < '" & Format(Year((nuefech)), "0000") & Format(Month((nuefech)), "00") & Format(Day((nuefech)), "00") & "'", LOGIC_AND
    miosql.AddComplexWhereClause "FACEP >= '" & Format(Year((ioFECHA.Text)), "0000") & Format(Month((ioFECHA.Text)), "00") & Format(Day((ioFECHA.Text)), "00") & "' AND FACEP < '" & Format(Year((nuefech)), "0000") & Format(Month((nuefech)), "00") & Format(Day((nuefech)), "00") & "')", LOGIC_OR
    usa_where = True
         
End If

If ioIMPORTE.Text <> "" Then
    If CDbl(ioIMPORTE.Text) > 0 Then
        miosql.AddSimpleWhereClause "PVP", CDbl(ioIMPORTE.Text), , , LOGIC_AND
        usa_where = True
    End If
End If

'buscar las correspondientes a un ticket
If (ioTICKET.Text <> "") And Len(ioTICKET.Text) = 12 Then
    miosql.AddSimpleWhereClause "CODVEN", Left(ioTICKET.Text, 9), , , LOGIC_AND
    miosql.AddSimpleWhereClause "CODCAJA", Right(ioTICKET.Text, 3), , , LOGIC_AND
    usa_where = True
End If

'buscar por codigo
If ioCODIGO.Text <> "" And cbCAJAS.Text <> "" Then
    miosql.AddSimpleWhereClause "CODIGO", ioCODIGO.Text, , , LOGIC_AND
    miosql.AddSimpleWhereClause "CODCAJA", CByte(cbCAJAS.Text), , , LOGIC_AND
    usa_where = True
ElseIf ioCODIGO.Text <> "" And cbCAJAS.Text = "" Then
    MsgBox "Para buscar por código, se debe seleccionar también CAJA", vbInformation, titulo
    Exit Sub
End If

If cbCAJAS.Text <> "" Then
    miosql.AddSimpleWhereClause "CODCAJA", CByte(cbCAJAS.Text), , , LOGIC_AND
    usa_where = True
End If

If cbESTADO.Text <> "" Then
    miosql.AddSimpleWhereClause "ESTADO", cbESTADO.Text, , , LOGIC_AND
    usa_where = True
End If

If cbTIPO.Text <> "" Then
    miosql.AddSimpleWhereClause "TIPO", cbTIPO.Text, , , LOGIC_AND
    usa_where = True
End If

'si deja todo en blanco, no mostrar ningun registro
If Not usa_where Then
    fg.Clear
    Exit Sub
End If

'If ioMODELO.Text <> "" Then miosql.AddComplexWhereClause "CODART IN (" & artSQL.SQL & ")", LOGIC_AND
'If ioTICKET.Text <> "" Then miosql.AddComplexWhereClause "CODCOST IN (" & masql.SQL & ")", LOGIC_AND

If miRc.State = 1 Then miRc.Close
miRc.Open miosql.SQL, locCnn, adOpenStatic, adLockOptimistic

fg.Rows = 1
'Set fg.DataSource = miRc

Call carga_grid

fg.HighLight = flexHighlightWithFocus
fg.FocusRect = flexFocusHeavy

'fg.ColHidden(fg.Cols - 1) = True

DoEvents

    With fg
    
    .ColFormat(0) = "000000000"
    .AutoSize 0, .Cols - 1
    
    DoEvents

 End With


   On Error GoTo 0
   Exit Sub

cbLista_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbLista_click de Formulario frmFlexArre"

End Sub


'---------------------------------------------------------------------------------------
' Procedimiento : carga_grid
' Fecha/Hora    : 26/01/2004 09:59
' Autor         : JCastillo
' Propósito     :   Cargar el grid con vales
'---------------------------------------------------------------------------------------
'ESTADO 0=PENDIENTE, 1=SERVIDO, 2=CANCELADO
'---------------------------------------------------------------------------------------
Private Sub carga_grid()
Dim tmpcodcolor As Long
Dim conta_lineas As Long

   On Error GoTo carga_grid_Error

   With fg
   
    .Clear
    .Cols = 12
    .ColFormat(5) = "Currency"
    .ColHidden(0) = True
    .ColHidden(1) = True
    .Rows = 1
    
    .TextMatrix(0, 2) = "Código"
    .TextMatrix(0, 3) = "Fecha Emi"
    .TextMatrix(0, 4) = "Fecha Rec"
    .TextMatrix(0, 5) = "Importe"
    .TextMatrix(0, 6) = "DCTO"
    .TextMatrix(0, 7) = "Estado"
    .TextMatrix(0, 8) = "Tipo"
    .TextMatrix(0, 9) = "Cliente"
    .TextMatrix(0, 10) = "Caja"
    .TextMatrix(0, 11) = "Usuario"
    
    
    If miRc.RecordCount <= 0 Then Exit Sub
           
    Do
             .Rows = .Rows + 1
    
        If Not miRc.EOF Then
     
            conta_lineas = conta_lineas + 1
         
            'contador de linea para ir al registro especificado
            .TextMatrix(.Rows - 1, 0) = conta_lineas
            'CAJA
            .TextMatrix(.Rows - 1, 1) = miRc.fields("CODCAJA")
                        
             'CODIGO
            .TextMatrix(.Rows - 1, 2) = miRc.fields("CODIGO")
            
            'FECHA
            .TextMatrix(.Rows - 1, 3) = miRc.fields("FMODI")
            
            'Importe
            .TextMatrix(.Rows - 1, 5) = miRc.fields("IMPORTE")
            
            'Dcto
            .TextMatrix(.Rows - 1, 6) = miRc.fields("DCTO") & " %"
            
                       
            'estado 0=pendiente, 1=aceptado
            If miRc.fields("ESTADO") = True Then
            .TextMatrix(.Rows - 1, 7) = "ACEPTADO"
            Else
            .TextMatrix(.Rows - 1, 7) = "PENDIENTE"
            End If
            
           '1=VENTA, 2=DEVOLUCION, 3=SOBRANTE 4=ANULADO
            Select Case miRc.fields("TIPO")
            Case 1
                 .TextMatrix(.Rows - 1, 8) = "VENTA"
            Case 2
                .TextMatrix(.Rows - 1, 8) = "DEVOLUCION"
            Case 3
                .TextMatrix(.Rows - 1, 8) = "SOBRANTE"
            Case 4
                .TextMatrix(.Rows - 1, 8) = "ANULADO"
            End Select
            
            'cliente
            If miRc.fields("CODCLI") > 0 And miRc.fields("CAJACLI") > 0 Then
                .TextMatrix(.Rows - 1, 9) = Trim(devuelve_campo("SELECT RAZO FROM CLIENTES WHERE CODIGO = " & miRc.fields("CODCLI") & " AND CODCAJA = " & miRc.fields("CAJACLI"), locCnn))
            End If
            
            'caja
            .TextMatrix(.Rows - 1, 10) = Trim(devuelve_campo("SELECT DESCRIPCION FROM CAJAS WHERE CODIGO = " & miRc.fields("CODCAJA"), locCnn))
                       
            'usuario
            .TextMatrix(.Rows - 1, 11) = Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & miRc.fields("CODPER"), locCnn))
     
        End If
    
    If Not miRc.EOF Then miRc.MoveNext
    
    Loop Until miRc.EOF
          
        .SubtotalPosition = flexSTAbove
        .subtotal flexSTCount, -1, 8, , vbBlue, vbWhite
        .subtotal flexSTSum, -1, 5, , vbBlue, vbWhite
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 7) = "Total Vales: (" & .TextMatrix(1, 8) & ")"
        .TextMatrix(1, 8) = ""
        
    .AutoSize 1, .Cols - 1
    .Redraw = True

  End With
  
   On Error GoTo 0
   Exit Sub

carga_grid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_grid de Formulario frmFlexArre"
 
 End Sub



Private Sub cbTIPO_LostFocus()

ioTICKET.SetFocus

End Sub

Private Sub chameleonButton2_Click()

Dim linea1 As String
Dim linea2 As String
Dim tmpcaja As String
         
   
   On Error GoTo chameleonButton2_Click_Error

    DoEvents

    If cbCAJAS.Text <> "" Then
        tmpcaja = devuelve_campo("SELECT DESCRIPCION FROM CAJAS WHERE CODIGO = " & cbCAJAS.Text, locCnn)
        If tmpcaja = "@" Then tmpcaja = ""
    End If
    
    linea1 = "Vales:  F.Inicial: " & ioFECHA.Text & ". F.Final: " & ioFECHAFIN.Text & ". Caja: " & tmpcaja
    linea2 = "Usuario: " & Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & UsuarioActual, locCnn)) & ". Fecha: " & Now
    
    Call PrintFlexGrid(fg, 1, 1, 2, linea1, linea2, 13, 2, 10)
    'fg.SaveGrid "c:\prueba.txt", flexFileCommaText
    'fg.PrintGrid "Transferencia " & rc.Fields("CODIGO"), , 2, 600, 600
    
    
    'actualizar los colores del grid, volivendolo a cargar
    'rc.Move 0


   On Error GoTo 0
   Exit Sub

chameleonButton2_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento chameleonButton2_Click de Formulario frmFlexVal"

End Sub

'IR A REGISTRO
Private Sub fg_dblClick()
    seleccionado = True
    
    If fg.Rows <= 1 Then Exit Sub
    
    If IsNumeric(fg.TextMatrix(fg.Row, 0)) Then
        'posicionarse en el registro
        miRc.Move (fg.TextMatrix(fg.Row, 0) - 1), 1
        DoEvents
        
        'si se llama desde el mantenimiento
        If desde_mnt Then
            Unload Me
            
        'si no se llama desde el mnt, abrir el mnt y posicionar en
        'el registro seleccionado actual
        Else
            
            If Trim(fg.TextMatrix(fg.Row, 2)) <> "" And Trim(fg.TextMatrix(fg.Row, 1)) <> "" Then
              If IsNumeric(fg.TextMatrix(fg.Row, 2)) And IsNumeric(fg.TextMatrix(fg.Row, 1)) Then
            
            With frmMntVal
                .Codigo_Vale = fg.TextMatrix(fg.Row, 2)
                .Caja_Vale = fg.TextMatrix(fg.Row, 1)
                .Show
            End With
            
            End If
            End If
            
            Unload Me
        
        End If
                
    End If
       
   ' Unload Me
End Sub

Private Sub fg_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

Case 13
    
    Call fg_dblClick
    seleccionado = True
    KeyAscii = 0
   ' Unload Me
    
End Select

End Sub

Private Sub fg_LostFocus()

fg.TabStop = False

End Sub

Private Sub Form_Activate()
    
    Me.Refresh
    DoEvents
    
    If Not first Then
    
       ' Set fg.DataSource = miRc
        DoEvents
        fg.Visible = True
        fg.AutoSearch = flexSearchFromCursor
        fg.ExplorerBar = flexExSortShow
         
       ' first = True
    End If
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode

'Ir al grid, o regresar
Case vbKeyF5
    
    If fg.Rows > 1 Then
        If fg.TabStop Then
            fg.TabStop = False
            ioTICKET.SetFocus
        Else
            fg.TabStop = True
            fg.Select 1, 1, 1, fg.Cols - 1
            fg.SetFocus
        End If
    End If
    KeyCode = 0

'salir del formulario actual
Case vbKeyF8

    KeyCode = 0
    Unload Me

End Select

End Sub

Private Sub Form_Load()

    With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
  End With

  fg.Visible = False
  fg.Rows = 1
  fg.Cols = 0
  
  'Cargar el micombo cajas
  With cbCAJAS
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM CAJAS WHERE MBAJA = 0 ORDER BY CODIGO"
    .LenCodigo = 1
    .CodigoWidth = 500
    .carga
    .Refresh
  End With
  
  With ioFECHA
    .dspFormat = "dd/mm/yyyy"
   .LongMaxima = 10
  End With
  
  With ioFECHAFIN
    .dspFormat = "dd/mm/yyyy"
   .LongMaxima = 10
  End With
  
    With ioIMPORTE
    .dspFormat = "Currency"
   .LongMaxima = 10
   .Alineacion = 1
   .SoloNumeros = True
  End With
  
 With ioCODIGO
   .LongMaxima = 30
 End With
  
 With ioTICKET
    'codigo de ticket:   9 dígitos la venta, y 3 dígitos la caja
     .dspFormat = "000000000000"
     .LongMaxima = 50
     .Alineacion = 1
     .SoloNumeros = True
 End With
  
 With cbESTADO
    .añade_item "0   - PENDIENTE"
    .añade_item "1   - ACEPTADO"
    .LenCodigo = 1
    .CodigoWidth = 300
 End With
  
 miosql.AddTable "VALES"
 
 cbCAJAS.Text = CajaActual
 ioFECHA.Text = Date

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    tmpstrcombo = ""
    'Set nif = Nothing
        
    'si hemos establecido un filtro que no devuelve ningun registro,
    'borrar filtro para que no de error al volver al formulario
    'If miRc.EOF Then Call cbBorrar_click
    
    'No descargar desde aqui, descargar desde el formulario desde donde
    'se llame
    If desde_mnt = False Then Set frmFlexVal = Nothing
End Sub



Private Sub ioFECHA_Validate(Cancel As Boolean)

If ioFECHA.Text <> "" Then Call cbLista_click

End Sub

Private Sub ioTICKET_GotFocus()

If Tab1.Tab > 0 Then Tab1.Tab = 0

End Sub

Private Sub ioTICKET_Validate(Cancel As Boolean)

If ioTICKET.Text <> "" Then Call cbLista_click

End Sub

