VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmFlexCli 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Clientes ..."
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11520
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
   ScaleHeight     =   6750
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   5250
      Left            =   15
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1485
      Visible         =   0   'False
      Width           =   11490
      _cx             =   20267
      _cy             =   9260
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
      FormatString    =   $"frmFlexCli.frx":0000
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
   Begin PCGestion.ucGrdBttn cbLista 
      Height          =   300
      Left            =   9690
      TabIndex        =   0
      Top             =   1155
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   529
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
      Image           =   "frmFlexCli.frx":00DE
   End
   Begin PCGestion.ucGrdBttn cbBorrar 
      Height          =   300
      Left            =   10740
      TabIndex        =   7
      Top             =   1155
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   529
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
      Image           =   "frmFlexCli.frx":00FA
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel4 
      Height          =   300
      Left            =   6375
      Top             =   1155
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   529
      Caption         =   "-F4- Consultar -F5- Ir a Rejilla  -F8- Salir"
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
   Begin TabDlg.SSTab Tab1 
      Height          =   1440
      Left            =   30
      TabIndex        =   8
      Top             =   15
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   2540
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
      TabPicture(0)   =   "frmFlexCli.frx":0116
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fwbajas"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label10"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ioTELEFONO"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ioNOMBRE"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ioCIF"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ioCODIGO"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cbCAJAS"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Hoja 2"
      TabPicture(1)   =   "frmFlexCli.frx":0132
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cbImprimir"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin PCGestion.miCombo cbCODTALLA 
         Height          =   495
         Left            =   -74220
         TabIndex        =   9
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
         TabIndex        =   10
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
      Begin PCGestion.miCombo miCombo2 
         Height          =   495
         Left            =   -66480
         TabIndex        =   11
         Top             =   30
         Width           =   3045
         _ExtentX        =   5821
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
      Begin PCGestion.miText ioFECHAINI 
         Height          =   480
         Index           =   0
         Left            =   -68910
         TabIndex        =   13
         Top             =   495
         Width           =   1425
         _ExtentX        =   2381
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
         dspFormat       =   ""
         Enabled         =   -1  'True
         EsPassword      =   -1  'True
      End
      Begin PCGestion.miText ioFECHAFIN 
         Height          =   480
         Index           =   0
         Left            =   -66465
         TabIndex        =   14
         Top             =   495
         Width           =   1410
         _ExtentX        =   2487
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
         dspFormat       =   ""
         Enabled         =   -1  'True
         EsPassword      =   -1  'True
      End
      Begin PCGestion.miCombo cbCATTALL 
         Height          =   495
         Left            =   -68910
         TabIndex        =   15
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
         TabIndex        =   16
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
      Begin PCGestion.miCombo cbSUBFAM 
         Height          =   480
         Left            =   -68910
         TabIndex        =   17
         Top             =   45
         Width           =   4140
         _ExtentX        =   7303
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
      Begin PCGestion.miCombo miCombo1 
         Height          =   495
         Left            =   -67815
         TabIndex        =   18
         Top             =   120
         Width           =   4005
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
      Begin PCGestion.miText ioIMPORTE 
         Height          =   495
         Left            =   -67800
         TabIndex        =   19
         Top             =   645
         Width           =   1260
         _ExtentX        =   2223
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
      Begin PCGestion.miText ioMOTIVO 
         Height          =   495
         Left            =   -73455
         TabIndex        =   20
         Top             =   645
         Width           =   4485
         _ExtentX        =   7911
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
      Begin PCGestion.miText ioFECHAI 
         Height          =   495
         Left            =   -73455
         TabIndex        =   21
         Top             =   105
         Width           =   1425
         _ExtentX        =   2514
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
      Begin PCGestion.miText ioFECHAF 
         Height          =   495
         Left            =   -70380
         TabIndex        =   22
         Top             =   120
         Width           =   1410
         _ExtentX        =   2487
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
      Begin PCGestion.miText ioVER 
         Height          =   480
         Left            =   -68355
         TabIndex        =   23
         Top             =   240
         Width           =   750
         _ExtentX        =   1323
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
         dspFormat       =   ""
         Enabled         =   -1  'True
         EsPassword      =   -1  'True
      End
      Begin PCGestion.chameleonButton chameleonButton1 
         Height          =   450
         Left            =   -67425
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   255
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
         MICON           =   "frmFlexCli.frx":014E
         PICN            =   "frmFlexCli.frx":016A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin PCGestion.miCombo cbESTADO 
         Height          =   495
         Left            =   -71820
         TabIndex        =   25
         Top             =   240
         Width           =   2790
         _ExtentX        =   4921
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
      Begin PCGestion.miText ioCODVAL 
         Height          =   495
         Left            =   -74160
         TabIndex        =   26
         Top             =   225
         Width           =   1515
         _ExtentX        =   2672
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
      Begin PCGestion.miCombo cbCAJAS 
         Height          =   495
         Left            =   7560
         TabIndex        =   4
         Top             =   60
         Width           =   3120
         _ExtentX        =   7805
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
      Begin PCGestion.miText ioCODIGO 
         Height          =   480
         Left            =   5970
         TabIndex        =   3
         Top             =   60
         Width           =   1050
         _ExtentX        =   1852
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
      Begin PCGestion.miText ioCIF 
         Height          =   480
         Left            =   3450
         TabIndex        =   2
         Top             =   60
         Width           =   1695
         _ExtentX        =   3043
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
      Begin PCGestion.miText ioNOMBRE 
         Height          =   495
         Left            =   1380
         TabIndex        =   5
         Top             =   585
         Width           =   4485
         _ExtentX        =   7911
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
      Begin PCGestion.miText ioTELEFONO 
         Height          =   525
         Left            =   1395
         TabIndex        =   1
         Top             =   60
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
      Begin PCGestion.chameleonButton cbImprimir 
         Height          =   795
         Left            =   -74715
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   105
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   1402
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
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
         MICON           =   "frmFlexCli.frx":0E44
         PICN            =   "frmFlexCli.frx":0E60
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NIF/CIF"
         Height          =   330
         Left            =   2670
         TabIndex        =   49
         Top             =   150
         Width           =   795
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "RAZO / NOMBRE"
         Height          =   615
         Left            =   210
         TabIndex        =   48
         Top             =   465
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CODIGO"
         Height          =   330
         Left            =   5100
         TabIndex        =   47
         Top             =   165
         Width           =   840
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CAJA"
         Height          =   330
         Left            =   7020
         TabIndex        =   46
         Top             =   150
         Width           =   540
      End
      Begin MSForms.CheckBox fwbajas 
         Height          =   435
         Left            =   5970
         TabIndex        =   45
         Top             =   615
         Width           =   1755
         VariousPropertyBits=   746588183
         BackColor       =   12632256
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "3096;767"
         Value           =   "0"
         Caption         =   "Ocultar BAJAS"
         FontName        =   "Trebuchet MS"
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TELEFONO"
         Height          =   300
         Left            =   240
         TabIndex        =   44
         Top             =   120
         Width           =   1125
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SUBFAM."
         Height          =   315
         Left            =   -69855
         TabIndex        =   43
         Top             =   120
         Width           =   960
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FAMILIA"
         Height          =   315
         Left            =   -74985
         TabIndex        =   42
         Top             =   120
         Width           =   960
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CAT. TALLA"
         Height          =   330
         Left            =   -70245
         TabIndex        =   41
         Top             =   615
         Width           =   1305
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "EN"
         Height          =   300
         Left            =   -74670
         TabIndex        =   40
         Top             =   540
         Width           =   450
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO"
         Height          =   345
         Left            =   -67290
         TabIndex        =   39
         Top             =   120
         Width           =   810
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "COLOR"
         Height          =   285
         Left            =   -71610
         TabIndex        =   38
         Top             =   105
         Width           =   735
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TALLA"
         Height          =   300
         Left            =   -74925
         TabIndex        =   37
         Top             =   120
         Width           =   690
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F.INICIAL"
         Height          =   285
         Left            =   -69960
         TabIndex        =   36
         Top             =   585
         Width           =   1035
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F.FINAL"
         Height          =   285
         Left            =   -67500
         TabIndex        =   35
         Top             =   585
         Width           =   1020
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "MOTIVO"
         Height          =   285
         Left            =   -74355
         TabIndex        =   34
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IMPORTE"
         Height          =   330
         Left            =   -68760
         TabIndex        =   33
         Top             =   735
         Width           =   885
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CAJA"
         Height          =   330
         Left            =   -68430
         TabIndex        =   32
         Top             =   210
         Width           =   540
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA INICIAL"
         Height          =   330
         Left            =   -75090
         TabIndex        =   31
         Top             =   180
         Width           =   1635
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA FINAL"
         Height          =   330
         Left            =   -71835
         TabIndex        =   30
         Top             =   180
         Width           =   1395
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO"
         Height          =   330
         Left            =   -72600
         TabIndex        =   29
         Top             =   315
         Width           =   780
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ver"
         Height          =   330
         Left            =   -68865
         TabIndex        =   28
         Top             =   315
         Width           =   480
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "VALE"
         Height          =   330
         Left            =   -74835
         TabIndex        =   27
         Top             =   315
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmFlexCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim first As Boolean

Dim tmprc As New ADODB.Recordset
Dim tmpstrcombo As String


Dim tmpcodrep As String
Dim tmpcodban As String
Dim tmpcodfcobro As String

'si se llama desde las ventas, debe mostrarse otra vez el formulario de
'ventas al salir
Public desde_ventas As Boolean

Public desde_pruebas As Boolean

Public Desde_Devol As Boolean

Public desde_vales As Boolean

'desde pagos de clientes
Public desde_pagos As Boolean


Public miosql As New clsSmartSQL
Public miRc As New ADODB.Recordset
Public seleccionado As Boolean

Dim nif As New clsNIF

Private Sub cbBorrar_click()

ioCODIGO.Text = ""
ioCIF.Text = ""
ioNOMBRE.Text = ""
ioTELEFONO.Text = ""
cbCAJAS.Text = ""
fwbajas.Value = True

fg.Clear
fg.Rows = 1

'Call cbLista_click

End Sub



Private Sub cbCAJAS_Validate(Cancel As Boolean)
Call cbLista_click
End Sub


Private Sub cbImprimir_Click()

Dim linea1 As String
Dim linea2 As String
         

   On Error GoTo cbImprimir_Click_Error

    DoEvents

    'linea1 = "Informe Clientes. Costurera: " & ioNOMBRE.Text & ". F.Inicial: " & ioFECHA.Text & ". F.Final: " & ioFECHAFIN.Text
    linea1 = "Usuario: " & Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & UsuarioActual, locCnn)) & ". Fecha: " & Now
    
    
    'fg.Font.Size = 9
    Call PrintFlexGrid(fg, 1, 1, 2, linea1, ".", 13, 2)
    'fg.Font.Size = 11
    
    'fg.SaveGrid "c:\prueba.txt", flexFileCommaText
    'fg.PrintGrid "Transferencia " & rc.Fields("CODIGO"), , 2, 600, 600
    
    
    'actualizar los colores del grid, volivendolo a cargar
    'rc.Move 0


   On Error GoTo 0
   Exit Sub

cbImprimir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbImprimir_Click de Formulario frmFlexCli"

End Sub

Private Sub fwbajas_GotFocus()

If fg.Rows > 1 Then
    DoEvents
    fg.TabStop = True
    fg.SetFocus
    
    'si solo nos devuelve un registro, seleccionar directamente ...
    If fg.Rows = 2 Then
        DoEvents
        Call fg_KeyPress(13)
    End If
        
End If

End Sub

Private Sub ioCIF_GotFocus()

If fg.Rows > 1 Then
    DoEvents
    fg.TabStop = True
    fg.SetFocus
    
    'si solo nos devuelve un registro, seleccionar directamente ...
    If fg.Rows = 2 Then
        DoEvents
        Call fg_KeyPress(13)
    End If
    
End If

End Sub

Private Sub ioCODIGO_GotFocus()

If fg.Rows > 1 Then
    DoEvents
    fg.TabStop = True
    fg.SetFocus
    
    'si solo nos devuelve un registro, seleccionar directamente ...
    If fg.Rows = 2 Then
        DoEvents
        Call fg_KeyPress(13)
    End If
    
End If

End Sub

Private Sub ioNOMBRE_GotFocus()

If fg.Rows > 1 Then
    DoEvents
    fg.TabStop = True
    fg.SetFocus
    
    'si solo nos devuelve un registro, seleccionar directamente ...
    If fg.Rows = 2 Then
        DoEvents
        Call fg_KeyPress(13)
    End If
    
End If

End Sub

Private Sub iotelefono_Validate(Cancel As Boolean)

If ioTELEFONO.Text <> "" Then Call cbLista_click
DoEvents

End Sub

Private Sub cbLista_click()
Dim usa_where As Boolean

miosql.ClearWhereClause

If ioCODIGO.Text <> "" Then
    miosql.AddSimpleWhereClause "CODIGO", CLng(ioCODIGO.Text)
    usa_where = True
End If

If ioNOMBRE.Text <> "" Then
    miosql.AddSimpleWhereClause "RAZO", ioNOMBRE.Text, , CLAUSE_LIKE
    usa_where = True
End If

If ioCIF.Text <> "" Then
    miosql.AddSimpleWhereClause "NIF", ioCIF.Text
    usa_where = True
End If

If ioTELEFONO.Text <> "" Then
    miosql.AddSimpleWhereClause "TELEFONO1", ioTELEFONO.Text
    usa_where = True
End If

If cbCAJAS.Text <> "" Then
    miosql.AddSimpleWhereClause "CODCAJA", cbCAJAS.Text
    usa_where = True
End If

'si decimos que ocultar bajas,
'MBAJA = FALSE
If fwbajas.Value = True Then
    miosql.AddSimpleWhereClause "MBAJA", 0
    usa_where = True
End If

'si deja todo en blanco, no mostrar ningun registro
If Not usa_where Then
    fg.Clear
    Exit Sub
End If

miRc.Close
miRc.Open miosql.SQL, locCnn, adOpenStatic, adLockOptimistic

fg.Rows = 1
Set fg.DataSource = miRc

fg.HighLight = flexHighlightWithFocus
fg.FocusRect = flexFocusHeavy

fg.ColHidden(fg.Cols - 1) = True


DoEvents

    With fg
    
    .ColFormat(0) = "000000000"
    .ColComboList(1) = tmpstrcombo
    .ColComboList(17) = tmpcodrep
    .ColComboList(22) = tmpcodfcobro
    .ColComboList(25) = tmpcodban
    .ColHidden(.Cols - 1) = True
    .AutoSize 0, .Cols - 1
    
    DoEvents
        
        
 End With




End Sub

Private Sub fg_dblClick()
    seleccionado = True
    Unload Me
End Sub

Private Sub fg_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

Case 13, vbKeyEscape
    seleccionado = True
    KeyAscii = 0
    Unload Me
    
End Select

End Sub

Private Sub fg_LostFocus()

fg.TabStop = False

End Sub

Private Sub Form_Activate()
    
    Me.Refresh
    DoEvents
    
    If Not first Then
    
        With ioCODIGO
            .SoloNumeros = True
            .LongMaxima = 9
            .dspFormat = "000000000"
        End With
               
       ' Set fg.DataSource = miRc
        DoEvents
        fg.Visible = True
        fg.AutoSearch = flexSearchFromCursor
        fg.ExplorerBar = flexExSortShow
          
   ' If tmprc.State = 0 Then
    '    Unload Me
    '    Exit Sub
    'End If
    
    tmprc.Open "SELECT CODIGO, DESCRIPCION FROM CAJAS WHERE MBAJA = 0 ORDER BY CODIGO", locCnn, adOpenDynamic, adLockReadOnly
  
    tmpstrcombo = fg.BuildComboList(tmprc, "DESCRIPCION", "CODIGO", vbBlue)
    'fg.ColComboList(2) = tmpstrcombo
    'fg.ColFormat(1) = "000"
    'fg.AutoSize 1, fg.Cols - 1
    
    
    With tmprc
 
    If .State = 1 Then .Close
    tmprc.Open "SELECT CODIGO, NOMBRE FROM BANCOS WHERE MBAJA = 0 ORDER BY CODIGO", locCnn, adOpenDynamic, adLockReadOnly
    tmpcodban = fg.BuildComboList(tmprc, "NOMBRE", "CODIGO", vbBlue)
    .Close
    tmprc.Open "SELECT CODIGO, DESCRIPCION FROM FCOBRO WHERE MBAJA = 0 ORDER BY CODIGO", locCnn, adOpenDynamic, adLockReadOnly
    tmpcodfcobro = fg.BuildComboList(tmprc, "DESCRIPCION", "CODIGO", vbBlue)
    .Close
    tmprc.Open "SELECT CODIGO, NOMBRE FROM REPRESEN WHERE MBAJA = 0 ORDER BY CODIGO", locCnn, adOpenDynamic, adLockReadOnly
    tmpcodrep = fg.BuildComboList(tmprc, "NOMBRE", "CODIGO", vbBlue)
    .Close

    End With
    
    'tmprc.Close
    Set tmprc = Nothing
    

        first = True
    End If
    
    
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode

'Ir al grid, o regresar
Case vbKeyF5
    
    If fg.Rows > 1 Then
        If fg.TabStop Then
            fg.TabStop = False
            ioNOMBRE.SetFocus
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
    
'consultar
Case vbKeyF4

    KeyCode = 0
    Call cbLista_click

End Select

End Sub

Private Sub Form_Load()

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
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error GoTo Form_QueryUnload_Error

    tmpstrcombo = ""
    Set nif = Nothing
        
    'si hemos establecido un filtro que no devuelve ningun registro,
    'borrar filtro para que no de error al volver al formulario
    If miRc.State = 0 Then Exit Sub
    
    If desde_ventas Or desde_pruebas Or Desde_Devol Or desde_vales Or desde_pagos Then
         
        If seleccionado Then
            'se asigna el cliente de manera exiterna
            
            If desde_ventas Then
                Call frmCabVen.Asignar_cliente_flex(miRc.fields("CODIGO"), miRc.fields("CODCAJA"))
                frmCabVen.Visible = True
                desde_ventas = False
            End If
            
            If desde_pruebas Then
                frmCabPrue.Visible = True
                Call frmCabPrue.Asignar_cliente_flex(miRc.fields("CODIGO"), miRc.fields("CODCAJA"))
                desde_pruebas = False
            End If
            
            If Desde_Devol Then
                frmMntDev.Visible = True
                Call frmMntDev.Asignar_cliente_flex(miRc.fields("CODIGO"), miRc.fields("CODCAJA"))
                Desde_Devol = False
            End If
            
            If desde_vales Then
                frmMntVal.Visible = True
                Call frmMntVal.Asignar_cliente_flex(miRc.fields("CODIGO"), miRc.fields("CODCAJA"))
                desde_vales = False
            End If
            
            If desde_pagos Then
                frmCabPag.Visible = True
                Call frmCabPag.Asignar_cliente_flex(miRc.fields("CODIGO"), miRc.fields("CODCAJA"))
                desde_pagos = False
            End If
            
            DoEvents
        
        Else
        
            If desde_ventas Then frmCabVen.Visible = True
            If desde_pruebas Then frmCabPrue.Visible = True
            If Desde_Devol Then frmMntDev.Visible = True
            If desde_vales Then frmMntVal.Visible = True
            If desde_pagos Then frmCabPag.Visible = True
                
        End If
                  

        
        miRc.Close
        Set miRc = Nothing
        Set frmFlexCli = Nothing
        
        Exit Sub
        
    End If
    
    If miRc.EOF Then Call cbBorrar_click
    
    
    'No descargar desde aqui, descargar desde el formulario desde donde
    'se llame
    'Set frmFlexCli = Nothing

   On Error GoTo 0
   Exit Sub

Form_QueryUnload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Form_QueryUnload de Formulario frmFlexCli"
End Sub



Private Sub ioCIF_Validate(Cancel As Boolean)

'si esta a blancos salir (no consultar nif)
If Trim(ioCIF.Text) = "" Then
   ' ioCIF.CancelarValidacion
    'Cancel = True
    Exit Sub
End If

nif.DarFormato = True
nif.nif = ioCIF.Text

If nif.Err Then
    ioCIF.CancelarValidacion
    Cancel = True
    Exit Sub
Else
    ioCIF.Text = nif.nif
End If


Call cbLista_click

'If ioCIF.Text <> "" Then Call comprueba_DNI(ioCIF.Text, ioCIF)
End Sub



Private Sub ioCODIGO_Validate(Cancel As Boolean)
Call cbLista_click
End Sub

Private Sub ioNOMBRE_Validate(Cancel As Boolean)

Call cbLista_click

End Sub

Private Sub Label16_Click()
End Sub

