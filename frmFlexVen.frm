VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmFlexVen 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver Ventas ..."
   ClientHeight    =   7545
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
   ScaleHeight     =   7545
   ScaleWidth      =   11460
   StartUpPosition =   2  'CenterScreen
   Begin PCGestion.ucGrdBttn cbBorrar 
      Height          =   315
      Left            =   10695
      TabIndex        =   24
      Top             =   1230
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   556
      Caption         =   "&Borrar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "frmFlexVen.frx":0000
   End
   Begin PCGestion.ucGrdBttn cbLista 
      Height          =   315
      Left            =   9450
      TabIndex        =   23
      Top             =   1230
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "&Consultar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "frmFlexVen.frx":001C
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel4 
      Height          =   285
      Left            =   4305
      Top             =   1245
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   503
      Caption         =   "-F4- Consultar -F5- Ir a Rejilla  -F8- Salir  -F9- Imprimir Ticket"
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
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   6015
      Left            =   15
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1530
      Visible         =   0   'False
      Width           =   11445
      _cx             =   20188
      _cy             =   10610
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
      FormatString    =   $"frmFlexVen.frx":0038
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
   Begin TabDlg.SSTab Tab1 
      Height          =   1530
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   2699
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   12632256
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
      TabPicture(0)   =   "frmFlexVen.frx":0116
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cbEN"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cbCODPROV"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ioMODELO"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ioIMPORTE"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ioREF"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "ioFECHA"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cbCAJAS"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Hoja 2"
      TabPicture(1)   =   "frmFlexVen.frx":0132
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cbSECCION"
      Tab(1).Control(1)=   "cbFAMILIA"
      Tab(1).Control(2)=   "cbSUBFAM"
      Tab(1).Control(3)=   "cbCODPER"
      Tab(1).Control(4)=   "Label11"
      Tab(1).Control(5)=   "Label10"
      Tab(1).Control(6)=   "Label9"
      Tab(1).Control(7)=   "Label8"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Hoja 3"
      TabPicture(2)   =   "frmFlexVen.frx":014E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label12"
      Tab(2).Control(1)=   "Label13"
      Tab(2).Control(2)=   "Label14"
      Tab(2).Control(3)=   "Label15"
      Tab(2).Control(4)=   "Label16"
      Tab(2).Control(5)=   "ioTICKET"
      Tab(2).Control(6)=   "cbESTADO"
      Tab(2).Control(7)=   "bsGradientLabel1"
      Tab(2).Control(8)=   "cbFCOBRO"
      Tab(2).Control(9)=   "ioFECHAFIN"
      Tab(2).Control(10)=   "ioFECHAINI"
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "Hoja 4"
      TabPicture(3)   =   "frmFlexVen.frx":016A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label18"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label17"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label20"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label19"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label21"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label22"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "chameleonButton2"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "cbCOLOR"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "cbCATTALL"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "cbTALLA"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "ioCODART"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "cbTEMPOR"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "ioCODBAR"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).ControlCount=   13
      Begin PCGestion.miCombo cbCAJAS 
         Height          =   525
         Left            =   615
         TabIndex        =   4
         Top             =   600
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   926
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
      Begin PCGestion.miText ioFECHA 
         Height          =   495
         Left            =   10005
         TabIndex        =   3
         Top             =   60
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
      Begin PCGestion.miText ioREF 
         Height          =   465
         Left            =   630
         TabIndex        =   0
         Top             =   60
         Width           =   2460
         _ExtentX        =   4339
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
      Begin PCGestion.miText ioIMPORTE 
         Height          =   480
         Left            =   8025
         TabIndex        =   2
         Top             =   60
         Width           =   1260
         _ExtentX        =   2223
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
      Begin PCGestion.miText ioMODELO 
         Height          =   480
         Left            =   3930
         TabIndex        =   1
         Top             =   60
         Width           =   3180
         _ExtentX        =   5609
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
      Begin PCGestion.miCombo cbCODPROV 
         Height          =   525
         Left            =   5400
         TabIndex        =   5
         Top             =   600
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   926
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
      Begin PCGestion.miCombo cbEN 
         Height          =   525
         Left            =   9795
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   926
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
      Begin PCGestion.miCombo cbSECCION 
         Height          =   525
         Left            =   -73455
         TabIndex        =   7
         Top             =   90
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   926
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
      Begin PCGestion.miCombo cbFAMILIA 
         Height          =   495
         Left            =   -68160
         TabIndex        =   8
         Top             =   105
         Width           =   4125
         _ExtentX        =   7276
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
      Begin PCGestion.miCombo cbSUBFAM 
         Height          =   525
         Left            =   -73470
         TabIndex        =   9
         Top             =   585
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   926
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
      Begin PCGestion.miCombo cbCODPER 
         Height          =   525
         Left            =   -68160
         TabIndex        =   10
         Top             =   600
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   926
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
         Height          =   495
         Left            =   -67710
         TabIndex        =   14
         Top             =   600
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
      Begin PCGestion.miText ioFECHAFIN 
         Height          =   495
         Left            =   -65025
         TabIndex        =   15
         Top             =   600
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
      Begin PCGestion.miCombo cbFCOBRO 
         Height          =   495
         Left            =   -73695
         TabIndex        =   11
         Top             =   90
         Width           =   4410
         _ExtentX        =   7779
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
      Begin PCGestion.bsGradientLabel bsGradientLabel1 
         Height          =   315
         Left            =   -66990
         Top             =   135
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   556
         Caption         =   "Intervalo de Fechas"
         BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionColour   =   0
         Colour1         =   14457707
         Colour2         =   11311500
         CaptionAlignment=   1
      End
      Begin PCGestion.miCombo cbESTADO 
         Height          =   495
         Left            =   -71520
         TabIndex        =   13
         Top             =   600
         Width           =   2235
         _ExtentX        =   3942
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
      Begin PCGestion.miText ioTICKET 
         Height          =   495
         Left            =   -73695
         TabIndex        =   12
         Top             =   600
         Width           =   1350
         _ExtentX        =   2381
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
      Begin PCGestion.miText ioCODBAR 
         Height          =   480
         Left            =   -74610
         TabIndex        =   16
         Top             =   90
         Width           =   2550
         _ExtentX        =   4657
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
      Begin PCGestion.miCombo cbTEMPOR 
         Height          =   480
         Left            =   -68895
         TabIndex        =   18
         Top             =   90
         Width           =   1785
         _ExtentX        =   3149
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
      Begin PCGestion.miText ioCODART 
         Height          =   480
         Left            =   -71085
         TabIndex        =   17
         Top             =   90
         Width           =   1620
         _ExtentX        =   2858
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
      Begin PCGestion.miCombo cbTALLA 
         Height          =   540
         Left            =   -74205
         TabIndex        =   20
         Top             =   630
         Width           =   2895
         _ExtentX        =   5106
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
      Begin PCGestion.miCombo cbCATTALL 
         Height          =   495
         Left            =   -66375
         TabIndex        =   19
         Top             =   90
         Width           =   2745
         _ExtentX        =   4842
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
      Begin PCGestion.miCombo cbCOLOR 
         Height          =   540
         Left            =   -70410
         TabIndex        =   21
         Top             =   630
         Width           =   3540
         _ExtentX        =   6244
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
      Begin PCGestion.chameleonButton chameleonButton2 
         Height          =   600
         Left            =   -64605
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   570
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1058
         BTYPE           =   9
         TX              =   ""
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
         MICON           =   "frmFlexVen.frx":0186
         PICN            =   "frmFlexVen.frx":01A2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CAT. TALLA"
         Height          =   555
         Left            =   -67230
         TabIndex        =   47
         Top             =   15
         Width           =   855
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TALLA"
         Height          =   330
         Left            =   -74925
         TabIndex        =   46
         Top             =   690
         Width           =   675
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "COLOR"
         Height          =   330
         Left            =   -71205
         TabIndex        =   45
         Top             =   705
         Width           =   735
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CODIGO ART."
         Height          =   570
         Left            =   -72075
         TabIndex        =   44
         Top             =   30
         Width           =   960
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TEMP."
         Height          =   285
         Left            =   -69525
         TabIndex        =   43
         Top             =   180
         Width           =   675
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CB"
         Height          =   330
         Left            =   -75030
         TabIndex        =   42
         Top             =   165
         Width           =   390
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TICKET"
         Height          =   330
         Left            =   -74535
         TabIndex        =   41
         Top             =   690
         Width           =   825
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO"
         Height          =   330
         Left            =   -72375
         TabIndex        =   40
         Top             =   690
         Width           =   855
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F.COBRO"
         Height          =   330
         Left            =   -74745
         TabIndex        =   39
         Top             =   165
         Width           =   1020
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA FINAL"
         Height          =   330
         Left            =   -66375
         TabIndex        =   38
         Top             =   675
         Width           =   1350
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA INICIAL"
         Height          =   330
         Left            =   -69285
         TabIndex        =   37
         Top             =   675
         Width           =   1530
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "USUARIO"
         Height          =   330
         Left            =   -69210
         TabIndex        =   36
         Top             =   675
         Width           =   1020
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SUBFAMILIA"
         Height          =   330
         Left            =   -74775
         TabIndex        =   35
         Top             =   660
         Width           =   1290
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SECCION"
         Height          =   330
         Left            =   -74400
         TabIndex        =   34
         Top             =   165
         Width           =   930
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FAMILIA"
         Height          =   330
         Left            =   -69015
         TabIndex        =   33
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "EN"
         Height          =   330
         Left            =   9345
         TabIndex        =   32
         Top             =   690
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "REF."
         Height          =   285
         Left            =   45
         TabIndex        =   31
         Top             =   120
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA"
         Height          =   330
         Left            =   9270
         TabIndex        =   30
         Top             =   150
         Width           =   705
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CAJA"
         Height          =   330
         Left            =   45
         TabIndex        =   29
         Top             =   675
         Width           =   540
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IMPORTE"
         Height          =   330
         Left            =   7125
         TabIndex        =   28
         Top             =   135
         Width           =   885
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "MODELO"
         Height          =   285
         Left            =   3075
         TabIndex        =   27
         Top             =   120
         Width           =   870
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PROV."
         Height          =   330
         Left            =   4755
         TabIndex        =   26
         Top             =   690
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmFlexVen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'---------------------------------------------------------------------------------------
' Modulo      : frmFlexVen
' Fecha/Hora  : 09/02/2004 23:56
' Autor       : JCASTILLO
' Propósito   : Ver detalle de ventas
'---------------------------------------------------------------------------------------
Option Explicit

Dim first As Boolean

Dim tmpsubt As Currency
Dim totalgeneral As Currency

'Dim tmpstrcombo As String

Dim CabVenSQL As New clsSmartSQL
Dim DetVenSQL As New clsSmartSQL
Dim artsql As New clsSmartSQL

Dim drc As New ADODB.Recordset
Dim miRc As New ADODB.Recordset
Dim seleccionado As Boolean

Dim importe_arreglos As Currency

Dim mostrar_total_real As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Si entra desde devoluciones, asignar:
'codart
'tempor
'codtalla
'cocol
'unidades
'importe
'a la devolución actual.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Desde_Devol As Boolean
Public D_Codart As Long
Public D_Tempor As Byte
Public D_CodTalla As Integer
Public D_CodCol As Integer
Public D_Unidades As Single
Public D_Importe As Single
Public D_Cancelado As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Sub cbBorrar_click()

ioREF.Text = ""
ioMODELO.Text = ""
ioIMPORTE.Text = ""
ioFECHA.Text = Date
cbCAJAS.Text = CajaActual
cbCODPROV.Text = ""
cbEN.Text = ""
cbSECCION.Text = ""
cbFAMILIA.Text = ""
cbSUBFAM.Text = ""
cbCODPER.Text = ""
cbFCOBRO.Text = ""
ioFECHAINI.Text = ""
ioFECHAFIN.Text = ""

fg.Clear
fg.Rows = 1

ioREF.SetFocus

End Sub









Private Sub cbCATTALL_Validate(Cancel As Boolean)

If cbCATTALL.Text = "" Then Exit Sub

With cbTALLA
    .ConexionString = locCnn
    .LenCodigo = 3
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM TALLAS WHERE (MBAJA = 0) AND (CODIGO > 0) AND CATTALL = " & cbCATTALL.Text & " ORDER BY CODIGO"
    .carga
    .CodigoWidth = 500
End With

End Sub

Private Sub cbCODPER_GotFocus()
If Tab1.Tab <> 1 Then Tab1.Tab = 1
End Sub



Private Sub cbCOLOR_lostfocus()
Tab1.Tab = 0
ioREF.SetFocus
End Sub

Private Sub cbEN_GotFocus()

If Tab1.Tab <> 0 Then Tab1.Tab = 0

End Sub

Private Sub cbFCOBRO_GotFocus()
If Tab1.Tab <> 2 Then Tab1.Tab = 2
End Sub


Private Sub cbLista_click()
Dim usa_where As Boolean
Dim usa_artic As Boolean
Dim nuefech As String
Dim tmpimp As Variant
Dim DevolSQL As New clsSmartSQL
Dim miCod As MiCodBar
Dim tmpffin As Date

   On Error GoTo cbLista_click_Error

CabVenSQL.ClearWhereClause
DetVenSQL.ClearWhereClause
artsql.ClearWhereClause

DevolSQL.AddTable "DEVOL"

'si fechas estan en blanco, poner la de hoy por defecto
If (ioFECHA.Text = "") And (ioFECHAINI.Text = "") Then ioFECHA.Text = Date

'comprobar si quiere solo la fecha de hoy
If (ioFECHA.Text <> "") And IsDate(ioFECHA.Text) Then

    'CabVenSQL.AddSimpleWhereClause "FALTA", ioFECHA.Text, , CLAUSE_GREATERTHANOREQUAL
    'CabVenSQL.AddSimpleWhereClause "FALTA", CStr(DateAdd("d", 1, ioFECHA.Text)), , CLAUSE_LESSTHAN, LOGIC_AND
    'CabVenSQL.AddComplexWhereClause "Year(FALTA IN (" & DetVenSQL.SQL & ")", LOGIC_AND
    
    '>= q el dia actual
    '< que el dia siguiente
    nuefech = DateAdd("d", 1, ioFECHA.Text)
    CabVenSQL.AddComplexWhereClause "(FHORA >= '" & Format(Year((ioFECHA.Text)), "0000") & Format(Month((ioFECHA.Text)), "00") & Format(Day((ioFECHA.Text)), "00") & "' AND FHORA < '" & Format(Year((nuefech)), "0000") & Format(Month((nuefech)), "00") & Format(Day((nuefech)), "00") & "')", LOGIC_AND
    
    DevolSQL.AddComplexWhereClause "(FMODI >= '" & Format(Year((ioFECHA.Text)), "0000") & Format(Month((ioFECHA.Text)), "00") & Format(Day((ioFECHA.Text)), "00") & "' AND FMODI < '" & Format(Year((nuefech)), "0000") & Format(Month((nuefech)), "00") & Format(Day((nuefech)), "00") & "')", LOGIC_AND
    
    usa_where = True
    
       
'comprobar si quiere un rango de fechas
ElseIf (ioFECHAINI.Text <> "" And ioFECHAINI.Text <> "") And (IsDate(ioFECHAINI.Text) And IsDate(ioFECHAINI.Text)) Then

    
    tmpffin = ioFECHAFIN.Text
    tmpffin = DateAdd("d", 1, tmpffin)
    
    CabVenSQL.AddComplexWhereClause "FHORA >= '" & Format(Year((ioFECHAINI.Text)), "0000") & Format(Month((ioFECHAINI.Text)), "00") & Format(Day((ioFECHAINI.Text)), "00") & "' AND FHORA <= '" & Format(Year((ioFECHAFIN.Text)), "0000") & Format(Month((ioFECHAFIN.Text)), "00") & Format(Day((ioFECHAFIN.Text)), "00") & "'", LOGIC_AND
    DevolSQL.AddComplexWhereClause "FMODI >= '" & Format(Year((ioFECHAINI.Text)), "0000") & Format(Month((ioFECHAINI.Text)), "00") & Format(Day((ioFECHAINI.Text)), "00") & "' AND FMODI < '" & Format(Year((tmpffin)), "0000") & Format(Month((tmpffin)), "00") & Format(Day((tmpffin)), "00") & "'", LOGIC_AND
    
    usa_where = True
     
    
End If

   

CabVenSQL.AddSimpleWhereClause "ESTADO", 1, , , LOGIC_AND
DevolSQL.AddSimpleWhereClause "MBAJA", 0, , , LOGIC_AND

 
If ioTICKET.Text <> "" Then
    '"CONVERT(char(10), CODIGO) + CONVERT(char(3), CODCAJA)"
    CabVenSQL.AddComplexWhereClause "CODIGO = " & CLng((Left(ioTICKET.Text, Len(ioTICKET.Text) - 3))) & " AND CODCAJA = " & CLng((Right(ioTICKET.Text, 3)))
    usa_where = True
End If
 
If ioIMPORTE.Text <> "" Then
    If CDbl(ioIMPORTE.Text) > 0 Then
        DetVenSQL.AddSimpleWhereClause "PREVEN", CDbl(ioIMPORTE.Text), , , LOGIC_AND
        usa_where = True
    End If
End If

If cbSECCION.Text <> "" Then
    artsql.AddSimpleWhereClause "SECCION", cbSECCION.Text, , , LOGIC_AND
    usa_artic = True
End If

If cbFAMILIA.Text <> "" Then
    artsql.AddSimpleWhereClause "FAMILIA", cbFAMILIA.Text, , , LOGIC_AND
    usa_artic = True
End If

If cbSUBFAM.Text <> "" Then
    artsql.AddSimpleWhereClause "SUBFAM", cbSUBFAM.Text, , , LOGIC_AND
    usa_artic = True
End If

If ioMODELO.Text <> "" Then
    artsql.AddSimpleWhereClause "MODELO", ioMODELO.Text, , CLAUSE_LIKE, LOGIC_AND
    usa_artic = True
End If

If cbCODPROV.Text <> "" Then
    artsql.AddSimpleWhereClause "CODPROV", cbCODPROV.Text, , , LOGIC_AND
    usa_artic = True
End If

If ioREF.Text <> "" Then
    artsql.AddSimpleWhereClause "REF", ioREF.Text, , CLAUSE_LIKE, LOGIC_AND
    usa_artic = True
End If

If cbCAJAS.Text <> "" Then
    DetVenSQL.AddSimpleWhereClause "CODCAJA", cbCAJAS.Text, , , LOGIC_AND
    DevolSQL.AddSimpleWhereClause "CODCAJA", CLng(cbCAJAS.Text), , , LOGIC_AND
    usa_where = True
End If

'buscar por el artículo exactamente ...

If ioCODBAR.Text <> "" Then

    miCod = Descompone_CBAR(ioCODBAR.Text)
    
    DetVenSQL.AddSimpleWhereClause "CODART", CLng(miCod.CODIGO_ART), , , LOGIC_AND
    DetVenSQL.AddSimpleWhereClause "TEMPOR", CByte(miCod.TEMPORADA_ART), , , LOGIC_AND
    DetVenSQL.AddSimpleWhereClause "CODTALLA", CInt(miCod.TALLA_ART), , , LOGIC_AND
    DetVenSQL.AddSimpleWhereClause "CODCOL", CInt(miCod.COLOR_ART), , , LOGIC_AND
    
    DevolSQL.AddSimpleWhereClause "CODART", CLng(miCod.CODIGO_ART), , , LOGIC_AND
    DevolSQL.AddSimpleWhereClause "TEMPOR", CByte(miCod.TEMPORADA_ART), , , LOGIC_AND
    DevolSQL.AddSimpleWhereClause "CODTALLA", CInt(miCod.TALLA_ART), , , LOGIC_AND
    DevolSQL.AddSimpleWhereClause "CODCOL", CInt(miCod.COLOR_ART), , , LOGIC_AND
    
    usa_where = True
    
End If

If cbTEMPOR.Text <> "" Then
    DetVenSQL.AddSimpleWhereClause "TEMPOR", cbTEMPOR.Text, , , LOGIC_AND
    DevolSQL.AddSimpleWhereClause "TEMPOR", CLng(cbTEMPOR.Text), , , LOGIC_AND
    usa_where = True
End If

If ioCODART.Text <> "" Then
    DetVenSQL.AddSimpleWhereClause "CODART", ioCODART.Text, , , LOGIC_AND
    DevolSQL.AddSimpleWhereClause "CODART", CLng(ioCODART.Text), , , LOGIC_AND
    usa_where = True
End If

If cbTALLA.Text <> "" Then
    DetVenSQL.AddSimpleWhereClause "CODTALLA", cbTALLA.Text, , , LOGIC_AND
    DevolSQL.AddSimpleWhereClause "CODTALLA", CLng(cbTALLA.Text), , , LOGIC_AND
    usa_where = True
End If

If cbCOLOR.Text <> "" Then
    DetVenSQL.AddSimpleWhereClause "CODCOL", cbCOLOR.Text, , , LOGIC_AND
    DevolSQL.AddSimpleWhereClause "CODCOL", CLng(cbCOLOR.Text), , , LOGIC_AND
    usa_where = True
End If

If cbCODPER.Text <> "" Then
    CabVenSQL.AddSimpleWhereClause "CODPER", cbCODPER.Text, , , LOGIC_AND
    DevolSQL.AddSimpleWhereClause "CODUSR", cbCODPER.Text, , , LOGIC_AND
    usa_where = True
End If

If cbEN.Text <> "" Then
    CabVenSQL.AddSimpleWhereClause "CAJADES", cbEN.Text, , , LOGIC_AND
    usa_where = True
End If

If cbESTADO.Text <> "" Then
    CabVenSQL.AddSimpleWhereClause "ESTADO", cbESTADO.Text, , , LOGIC_AND
    usa_where = True
End If

If cbFCOBRO.Text <> "" Then
    CabVenSQL.AddSimpleWhereClause "FCOBRO", cbFCOBRO.Text, , , LOGIC_AND
    usa_where = True
End If


'si deja todo en blanco, no mostrar ningun registro
If (Not usa_where) And (Not usa_artic) Then
    fg.Clear
    ioREF.SetFocus
    Exit Sub
End If

If usa_artic Then DetVenSQL.AddComplexWhereClause "(CONVERT(char(10), CODART) + CONVERT(char(3), TEMPOR)) IN (" & artsql.SQL & ")", LOGIC_AND
If usa_artic Then DevolSQL.AddComplexWhereClause "(CONVERT(char(10), CODART) + CONVERT(char(3), TEMPOR)) IN (" & artsql.SQL & ")", LOGIC_AND

'I  f usa_artic Then DetVenSQL.AddComplexWhereClause "COD IN (" & ArtSql.SQL & ")", LOGIC_AND
If usa_where Then DetVenSQL.AddComplexWhereClause "(CONVERT(char(10), CODVEN) + CONVERT(char(3), CODCAJA)) IN (" & CabVenSQL.SQL & ")", LOGIC_AND
'If usa_where Then DevolSQL.AddComplexWhereClause "(CONVERT(char(10), CODVEN) + CONVERT(char(3), CODCAJA)) IN (" & CabVenSQL.SQL & ")", LOGIC_AND


'If ioNOMBRE.Text <> "" Then CabVenSQL.AddComplexWhereClause "CODCOST IN (" & DetVenSQL.SQL & ")", LOGIC_AND

If miRc.State = 1 Then miRc.Close
miRc.Open DetVenSQL.SQL, locCnn, adOpenStatic, adLockOptimistic

Debug.Print DevolSQL.SQL

If drc.State = 1 Then drc.Close
drc.Open DevolSQL.SQL, locCnn, adOpenStatic, adLockOptimistic

fg.Rows = 1
'Set fg.DataSource = miRc


'tmpimp = devuelve_campo("SELECT SUM(PVP) FROM ARREGLOS WHERE (ESTADO = 2) AND (FMODI >= '" & Format(Year((ioFECHA.Text)), "0000") & Format(Month((ioFECHA.Text)), "00") & Format(Day((ioFECHA.Text)), "00") & "' AND FMODI < '" & Format(Year((nuefech)), "0000") & Format(Month((nuefech)), "00") & Format(Day((nuefech)), "00") & "')", locCnn)

tmpimp = devuelve_campo("SELECT SUM(PVP) FROM ARREGLOS WHERE (ESTADO = 2) AND CONVERT(char(10), CODVEN) + CONVERT(char(3), CODCAJ) IN (" & CabVenSQL.SQL & ")", locCnn)

'cargar el importe de los arreglos para esta selección, en la variable de importe arreglos
'(para luego restar del total, y obtener el total real)
If tmpimp = "@" Or Not IsNumeric(tmpimp) Then
        importe_arreglos = 0
Else
        importe_arreglos = tmpimp
End If

Call carga_grid
'Call carga_grid_devoluciones

fg.HighLight = flexHighlightWithFocus
fg.FocusRect = flexFocusHeavy

'fg.ColHidden(fg.Cols - 1) = True

DoEvents

    With fg
    
    .ColFormat(0) = "000000000"
    .AutoSize 0, .Cols - 1
    
    DoEvents
    
    totalgeneral = 0

 End With
 
   ioREF.SetFocus


   On Error GoTo 0
   Exit Sub

cbLista_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbLista_click de Formulario frmFlexArre"
    ioREF.SetFocus

End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : carga_grid
' Fecha/Hora    : 26/01/2004 09:59
' Autor         : JCastillo
' Propósito     : Cargar el grid con las ventas
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
Private Sub carga_grid()
Dim tmpcodcolor As Long
Dim conta_lineas As Long
Dim T_Cabecera As Variant
Dim t_articulo As Variant
Dim Dsp_Prov As String

Dim tmpdcto As Currency
Dim tmpiva As Currency

Dim totaldcto As Currency
Dim totiva As Currency

Dim totalpvp As Currency

Dim numdev As Long

Dim tmpcambios As Long



Dim totalreal As Currency

Dim totalpvporig As Currency





Dim v_venta As Long
Dim n_venta As Long

   On Error GoTo carga_grid_Error

   
   With fg
   
    .Clear
    .Cols = 21
    .ColFormat(10) = "Currency"
    .ColFormat(11) = "Currency"
    .ColFormat(13) = "Currency"
    .ColFormat(15) = "Currency"
        
    .ColAlignment(4) = flexAlignLeftCenter
    .ColAlignment(6) = flexAlignLeftCenter
    .ColHidden(0) = True
    .ColHidden(1) = True
    .Rows = 1
    
    .TextMatrix(0, 2) = "Fecha"
    .TextMatrix(0, 3) = "Prov."
    .TextMatrix(0, 4) = "Ref."
    .TextMatrix(0, 5) = "Modelo"
    .TextMatrix(0, 6) = "Talla"
    .TextMatrix(0, 7) = "Color"
    .TextMatrix(0, 8) = "Temp"
    .TextMatrix(0, 9) = "Uds"
    .TextMatrix(0, 10) = "PVP"
    .TextMatrix(0, 11) = "PVP. Orig"
    .TextMatrix(0, 12) = "Dcto"
    .TextMatrix(0, 13) = "Sub."
    .TextMatrix(0, 14) = "Iva"
    .TextMatrix(0, 15) = "Total"
    .TextMatrix(0, 16) = "Ticket"
    .TextMatrix(0, 17) = "F. Cobro"
    .TextMatrix(0, 18) = "Depend."
    .TextMatrix(0, 19) = "Caja"
    .TextMatrix(0, 20) = "CBarras"
    
        
    If miRc.RecordCount <= 0 Then Exit Sub
    
    
    Do
             .Rows = .Rows + 1
    
        v_venta = n_venta
        n_venta = miRc.fields("CODVEN")
        
        If Not miRc.EOF Then
     
            conta_lineas = conta_lineas + 1
         
            'ID
            .TextMatrix(.Rows - 1, 0) = conta_lineas
            'CAJA
            .TextMatrix(.Rows - 1, 1) = miRc.fields("CODCAJA")
            
            T_Cabecera = devuelve_matriz("SELECT FMODI, FCOBRO, CODPER, IMP_PRIMERA, IMP_SEGUNDA FROM CABVENTA WHERE CODCAJA = " & miRc.fields("CODCAJA") & " AND CODIGO = " & miRc.fields("CODVEN"), locCnn)
            
            'FECHA (cogerla de la cabecera)
            .TextMatrix(.Rows - 1, 2) = T_Cabecera(0)
            
            t_articulo = devuelve_matriz("SELECT MODELO, REF, CODPROV, PREVEN FROM MAARTIC WHERE CODIGO = " & miRc.fields("CODART") & " AND TEMPOR = " & miRc.fields("TEMPOR"), locCnn)
            
            If IsArray(t_articulo) Then
            
            Dsp_Prov = Trim(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & t_articulo(2), locCnn))
            
            End If
             
            'Proveedor
            .TextMatrix(.Rows - 1, 3) = Dsp_Prov
            
            If IsArray(t_articulo) Then
            
            'Referencia
            .TextMatrix(.Rows - 1, 4) = Trim(t_articulo(1))
            
             
            'MODELO
            .TextMatrix(.Rows - 1, 5) = Format(miRc.fields("CODART"), "00000") & "-" & Trim(t_articulo(0))
            
            End If
            
            'TALLA
            .TextMatrix(.Rows - 1, 6) = Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & miRc.fields("CODTALLA")))
            
            'obtener el texto del color y su codigo de color (para colorear
            'la celda del grid)
            'COLOR
            If miRc.fields("CODCOL") > 0 Then
      
                tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & miRc.fields("CODCOL"))
                .TextMatrix(.Rows - 1, 7) = Trim(devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & miRc.fields("CODCOL")))
                .Col = 7
                .Row = .Rows - 1
                .CellBackColor = tmpcodcolor
                .Col = 2
        
            End If
            
            'TEMP
            .TextMatrix(.Rows - 1, 8) = Trim(devuelve_campo("SELECT ABREVIA FROM TEMPOR WHERE IDTEM = " & miRc.fields("TEMPOR"), locCnn))
            
            'UDS
            .TextMatrix(.Rows - 1, 9) = miRc.fields("UNIDADES")
            
            'precio de venta
            .TextMatrix(.Rows - 1, 10) = miRc.fields("PREVEN")
             
            If IsArray(t_articulo) Then
            
                totalpvp = totalpvp + (miRc.fields("UNIDADES") * miRc.fields("PREVEN"))
                totalpvporig = totalpvporig + (miRc.fields("UNIDADES") * t_articulo(3))
            
                'Precio venta Original
                .TextMatrix(.Rows - 1, 11) = t_articulo(3)
            
            'si el precio de venta original, y el aplicado son distintos, resaltar ambas celdas con el
            'fondo blanco
            If miRc.fields("PREVEN") <> t_articulo(3) Then
            
                .Row = .Rows - 1
                .Col = 10
                .CellForeColor = vbBlue
                .CellBackColor = vbWhite
                .CellFontBold = True
                .Col = 11
                .CellForeColor = vbBlue
                .CellBackColor = vbWhite
                .CellFontBold = True
                tmpcambios = tmpcambios + 1
            
            End If
            
            End If
             
                        
            
            'DCTO
            .TextMatrix(.Rows - 1, 12) = miRc.fields("DCTO") & " %"
            
            tmpdcto = ((miRc.fields("PREVEN") * miRc.fields("DCTO")) / 100)
            
            totaldcto = totaldcto + tmpdcto
              
             'SUBT
            .TextMatrix(.Rows - 1, 13) = (miRc.fields("PREVEN") - tmpdcto) * miRc.fields("UNIDADES")
            
            '"Iva"
            .TextMatrix(.Rows - 1, 14) = miRc.fields("IVA") & " %"
                        
             
             tmpiva = ((.TextMatrix(.Rows - 1, 13) * miRc.fields("IVA")) / 100)
             
             totiva = tmpiva + totiva
             
             'Total
             .TextMatrix(.Rows - 1, 15) = .TextMatrix(.Rows - 1, 13) + tmpiva
                          
             totalgeneral = totalgeneral + (.TextMatrix(.Rows - 1, 13) + tmpiva)
            
              'codven
             .TextMatrix(.Rows - 1, 16) = miRc.fields("CODVEN") & Format(miRc.fields("CODCAJA"), "000")
             
             
             If IsArray(t_articulo) Then
              'fcobro
             .TextMatrix(.Rows - 1, 17) = Trim(devuelve_campo("SELECT DESCRIPCION FROM FCOBRO WHERE CODIGO = " & T_Cabecera(1), locCnn))

              'dependiente
             .TextMatrix(.Rows - 1, 18) = Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & T_Cabecera(2), locCnn))
             
             End If
              
              'caja
             .TextMatrix(.Rows - 1, 19) = Trim(devuelve_campo("SELECT DESCRIPCION FROM CAJAS WHERE CODIGO = " & miRc.fields("CODCAJA"), locCnn))
                               
              'codigo de barras
              .TextMatrix(.Rows - 1, 20) = Conforma_CB(miRc.fields("CODART"), miRc.fields("TEMPOR"), miRc.fields("CODTALLA"), miRc.fields("CODCOL"))
                
            'estado
            '.TextMatrix(.Rows - 1, 10) = miRc.Fields("ESTADO")
            
     
        End If
    
    
        'si es distinto, romper ...
        If (v_venta <> n_venta) Then
        
            Select Case T_Cabecera(1)
            Case 4
            
            Case 5, 6
                totalreal = totalreal + T_Cabecera(3)
            Case Else
                totalreal = totalreal + (T_Cabecera(3) + T_Cabecera(4))
    
            End Select
        
        End If
        
        
    If Not miRc.EOF Then miRc.MoveNext
    
    'DoEvents
    
    Loop Until miRc.EOF
    
    numdev = (carga_grid_devoluciones)
    
    Debug.Print totalgeneral
          
        .SubtotalPosition = flexSTAbove
        .subtotal flexSTSum, -1, 9, , vbBlue, vbWhite, True
        
       
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 4) = ""
        .TextMatrix(1, 7) = "Uds: " & .TextMatrix(1, 9)
        .TextMatrix(1, 10) = Format(totalpvp, "Currency")
        .TextMatrix(1, 13) = "" & Format(totalpvp - tmpsubt, "Currency") ' .TextMatrix(1, 12)
        .TextMatrix(1, 11) = "Brutas:" & Format(totalpvporig, "Currency") ' .TextMatrix(1, 12)
        .TextMatrix(1, 12) = "" & Format(totaldcto, "Currency") ' .TextMatrix(1, 12)
        .TextMatrix(1, 14) = "" & Format(totiva, "Currency") ' .TextMatrix(1, 12)
        .TextMatrix(1, 15) = "Netas:" & Format(totalgeneral, "Currency") '.TextMatrix(1, 14)
        .TextMatrix(1, 18) = "Devoluciones(" & numdev & "): " & Format(tmpsubt * -1, "Currency")
        .TextMatrix(1, 19) = "Dif. PVP(" & tmpcambios & "): " & Format(totalpvp - totalpvporig, "Currency")
        
     '   .TextMatrix(1, 16) = "Arreglos: " & Format(importe_arreglos, "Currency") '.TextMatrix(1, 14)
     '   .TextMatrix(1, 17) = "Real: " & Format(totalreal - importe_arreglos, "Currency") '.TextMatrix(1, 14)
                        
    .AutoSize 1, .Cols - 1
    .Redraw = True

  End With
  
  tmpsubt = 0
   
   DoEvents
   
   On Error GoTo 0
   Exit Sub

carga_grid_Error:
   
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_grid de Formulario frmFlexArre"
 
 End Sub




Private Sub cbSECCION_GotFocus()
If Tab1.Tab <> 1 Then Tab1.Tab = 1
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
    
    linea1 = "Ventas:  F.Inicial: " & ioFECHAINI.Text & ". F.Final: " & ioFECHAFIN.Text & ". Almacén: " & tmpcaja
    linea2 = "Usuario: " & Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & UsuarioActual, locCnn)) & ". Fecha: " & Now
    
   Call PrintFlexGrid(fg, 1, 1, 2, linea1, linea2, 13, 2, 9)
    
   ' Call PrintGrid(fg, 1, 4, 2, 2, linea1, linea2)
    'fg.SaveGrid "c:\prueba.txt", flexFileCommaText
    'fg.PrintGrid "Transferencia " & rc.Fields("CODIGO"), , 2, 600, 600
    
    
    'actualizar los colores del grid, volivendolo a cargar
    'rc.Move 0


   On Error GoTo 0
   Exit Sub

chameleonButton2_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento chameleonButton2_Click de Formulario frmFlexVen"

End Sub

Private Sub fg_DblClick()

If miRc.State = 0 Then Exit Sub
If miRc.RecordCount <= 0 Then Exit Sub

    seleccionado = True
    
If fg.Rows <= 1 Then Exit Sub
    
If IsNumeric(fg.TextMatrix(fg.Row, 0)) Then
        'posicionarse en el registro
        miRc.Move (fg.TextMatrix(fg.Row, 0) - 1), 1
        
        If Desde_Devol Then
        
            'preguntar al usuario ...
            If MsgBox("¿Desea introducir el artículo actual como DEVOLUCION? " & Chr(13) & _
            "Modelo: " & fg.TextMatrix(fg.Row, 3) & Chr(13) & _
            "Talla: " & fg.TextMatrix(fg.Row, 4) & Chr(13) & _
            "Color: " & fg.TextMatrix(fg.Row, 5) & Chr(13) & _
            "Importe: " & fg.TextMatrix(fg.Row, 12) & Chr(13) & _
            "(se devolverán las unidades al almacén)", vbQuestion + vbYesNo, titulo) = vbYes Then
                
           'ASIGNAR VARIABLES:
           D_Codart = miRc.fields("CODART")
           D_Tempor = miRc.fields("TEMPOR")
           D_CodTalla = miRc.fields("CODTALLA")
           D_CodCol = miRc.fields("CODCOL")
           D_Unidades = miRc.fields("UNIDADES")
           D_Importe = CDbl(fg.TextMatrix(fg.Row, 12))
           D_Cancelado = False
                              
           miRc.Close
                        
           DoEvents
                
           Unload Me
         
           Exit Sub
           
           End If
           
           
        
        'si no entramos desde devoluciones
        Else
        
           'presentar la venta en el formulario de ventas y permitir su modificación.
           
           Me.WindowState = 1
           
           With frmCabVen
            .Desde_Grid = True
            .D_Codigo = miRc.fields("CODVEN")
            .D_CodCaja = miRc.fields("CODCAJA")
            .Show
           End With
           DoEvents
           
           'miRc.Close
           
'           Unload Me
                         
        End If
        
        DoEvents
                
End If
  
End Sub

Private Sub fg_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

Case 13
    
    Call fg_DblClick
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
          
  

        first = True
    End If
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rc_cabven As ADODB.Recordset
Dim rc_detven As ADODB.Recordset
Dim rc_arreglos As ADODB.Recordset

Select Case KeyCode

'Ir al grid, o regresar
Case vbKeyF5
    
    If fg.Rows > 1 Then
        If fg.TabStop Then
            fg.TabStop = False
            ioREF.SetFocus
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
    
Case vbKeyF4
    KeyCode = 0
    Call cbLista_click

'imprimir ticket
Case vbKeyF9
    KeyCode = 0
    
    If miRc.State = 0 Then Exit Sub
    If miRc.RecordCount <= 0 Then Exit Sub
    If miRc.EOF Then
        MsgBox "No se ha seleccionado ningún artículo para imprimir su ticket", vbExclamation, titulo
        Exit Sub
    End If
        
    Set rc_cabven = New ADODB.Recordset
    Set rc_detven = New ADODB.Recordset
    Set rc_arreglos = New ADODB.Recordset
    
    'abrir datos de cabecera, detalle y arreglos ...
    rc_cabven.Open "SELECT * FROM CABVENTA WHERE CODIGO = " & miRc.fields("CODVEN") & " AND CODCAJA = " & miRc.fields("CODCAJA"), locCnn, adOpenDynamic, adLockReadOnly
    rc_detven.Open "SELECT * FROM DETVENTA WHERE CODVEN = " & miRc.fields("CODVEN") & " AND CODCAJA = " & miRc.fields("CODCAJA"), locCnn, adOpenStatic, adLockReadOnly
    rc_arreglos.Open "SELECT * FROM ARREGLOS WHERE CODVEN = " & miRc.fields("CODVEN") & " AND CODCAJ = " & miRc.fields("CODCAJA"), locCnn, adOpenDynamic, adLockReadOnly
    
    'imprimir ticket
    Call Imprime_Ticket(0, rc_cabven, rc_detven, rc_arreglos)
    
    DoEvents

    rc_cabven.Close
    rc_detven.Close
    rc_arreglos.Close
    
    Set rc_cabven = Nothing
    Set rc_detven = Nothing
    Set rc_arreglos = Nothing

End Select

End Sub



'---------------------------------------------------------------------------------------
' Subrutina   : carga_grid_devoluciones
' Fecha/Hora  : 30/06/2004 01:22
' Autor       : JCASTILLO
' Propósito   : Cargar las devoluciones correspondientes a los artículo seleccionados
'                   devuelve el nº de devoluciones para sacar por el subtotal
'---------------------------------------------------------------------------------------
Private Function carga_grid_devoluciones() As Long
Dim tmpcodcolor As Variant
Dim conta_lineas As Long
Dim T_Cabecera As Variant
Dim t_articulo As Variant
Dim Dsp_Prov As String

Dim tmpdcto As Currency
Dim tmpiva As Currency

Dim totaldcto As Currency
Dim totiva As Currency
Dim totalpvp As Currency
'Dim totalgeneral As Currency
Dim totalreal As Currency

Dim timporte As Currency

Dim v_venta As Long
Dim n_venta As Long
   
   On Error GoTo carga_grid_devoluciones_Error

   With fg
   
        
    If drc.RecordCount <= 0 Then Exit Function
    
    
    Do
        .Rows = .Rows + 1
    
        v_venta = n_venta
        n_venta = drc.fields("CODVEN")
        
        If Not drc.EOF Then
        
            timporte = drc.fields("IMPORTE") * -1
     
            conta_lineas = conta_lineas + 1
         
            'ID
            .TextMatrix(.Rows - 1, 0) = conta_lineas
            'CAJA
            .TextMatrix(.Rows - 1, 1) = drc.fields("CODCAJA")
            
            'T_Cabecera = devuelve_matriz("SELECT FMODI, FCOBRO, CODPER, IMP_PRIMERA, IMP_SEGUNDA FROM CABVENTA WHERE CODCAJA = " & drc.fields("CODCAJA") & " AND CODIGO = " & drc.fields("CODVEN"), locCnn)
            
            'FECHA (cogerla de la cabecera)
            .TextMatrix(.Rows - 1, 2) = drc.fields("FMODI")
            
            t_articulo = devuelve_matriz("SELECT MODELO, REF, CODPROV, PREVEN FROM MAARTIC WHERE CODIGO = " & drc.fields("CODART") & " AND TEMPOR = " & drc.fields("TEMPOR"), locCnn)
                        
            
            If IsArray(t_articulo) Then
            
                Dsp_Prov = Trim(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & t_articulo(2), locCnn))
             
                'Proveedor
                .TextMatrix(.Rows - 1, 3) = Dsp_Prov
            
                'Referencia
                .TextMatrix(.Rows - 1, 4) = Trim(t_articulo(1))
            
                'MODELO
                .TextMatrix(.Rows - 1, 5) = Format(drc.fields("CODART"), "00000") & "-" & Trim(t_articulo(0))
            
                'TALLA
                .TextMatrix(.Rows - 1, 6) = Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & drc.fields("CODTALLA")))
            
                'obtener el texto del color y su codigo de color (para colorear
                'la celda del grid)
                'COLOR
                If drc.fields("CODCOL") > 0 Then
      
                    tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & drc.fields("CODCOL"))
                    
                    If tmpcodcolor <> "@" Then
                        .TextMatrix(.Rows - 1, 7) = Trim(devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & drc.fields("CODCOL")))
                        .Col = 7
                        .Row = .Rows - 1
                        .CellBackColor = tmpcodcolor
                        .Col = 2
                    End If
        
                End If
            
                'TEMP
                .TextMatrix(.Rows - 1, 8) = Trim(devuelve_campo("SELECT ABREVIA FROM TEMPOR WHERE IDTEM = " & drc.fields("TEMPOR"), locCnn))
            
                'UDS
                ' .TextMatrix(.Rows - 1, 9) =
            
                totalgeneral = totalgeneral - drc.fields("IMPORTE") ' timporte
                tmpsubt = tmpsubt + drc.fields("IMPORTE")
            
            
                'precio de venta
                .TextMatrix(.Rows - 1, 10) = timporte
             
          '  totalpvp = totalpvp + (timporte)
            
            'Precio venta Original
            '.TextMatrix(.Rows - 1, 11) = t_articulo(3)
            
            'DCTO
                .TextMatrix(.Rows - 1, 12) = 0
            
            'tmpdcto = ((drc.fields("PREVEN") * drc.fields("DCTO")) / 100)
            
            'totaldcto = totaldcto + tmpdcto
              
             'SUBT
            '.TextMatrix(.Rows - 1, 13) = (timporte - tmpdcto)
            
            '"Iva"
                .TextMatrix(.Rows - 1, 14) = 0 'drc.fields("IVA") & " %"
                        
             
            ' tmpiva = ((.TextMatrix(.Rows - 1, 13) * drc.fields("IVA")) / 100)
             
           '  totiva = tmpiva + totiva
             
             'Total
                .TextMatrix(.Rows - 1, 15) = timporte + tmpiva
             
                          
             'totalgeneral = totalgeneral + (.TextMatrix(.Rows - 1, 13) + tmpiva)
            
              'codven
                If drc.fields("CODVEN") > 0 Then
                    .TextMatrix(.Rows - 1, 16) = drc.fields("CODVEN") & Format(drc.fields("CODCAJA"), "000")
                End If
             
              'fcobro
                .TextMatrix(.Rows - 1, 17) = "* Devolución *"

              'dependiente
                .TextMatrix(.Rows - 1, 18) = Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & drc.fields("CODUSR"), locCnn))
              
              'caja
                .TextMatrix(.Rows - 1, 19) = Trim(devuelve_campo("SELECT DESCRIPCION FROM CAJAS WHERE CODIGO = " & drc.fields("CODCAJA"), locCnn))
                               
              'codigo de barras
                .TextMatrix(.Rows - 1, 20) = Conforma_CB(drc.fields("CODART"), drc.fields("TEMPOR"), drc.fields("CODTALLA"), drc.fields("CODCOL"))
                
            'estado
            '.TextMatrix(.Rows - 1, 10) = drc.Fields("ESTADO")
            
     
      
    
    
        'si es distinto, romper ...
     '   If (v_venta <> n_venta) Then
      '
      '      Select Case T_Cabecera(1)
     '       Case 4
            
      ''      Case 5, 6
      '          totalreal = totalreal + T_Cabecera(3)
      '      Case Else
      '          totalreal = totalreal + (T_Cabecera(3) + T_Cabecera(4))
    '
      '      End Select
      '
        End If
        
         End If
        
      
        
    If Not drc.EOF Then drc.MoveNext
    
    'DoEvents
    
    Loop Until drc.EOF
          
      '  .SubtotalPosition = flexSTAbove
       ' .subtotal flexSTSum, -1, 9, , vbBlue, vbWhite, True
       
       ' .TextMatrix(1, 0) = ""
       ' .TextMatrix(1, 4) = ""
       ' .TextMatrix(1, 7) = "Uds: " & .TextMatrix(1, 9)
       ' .TextMatrix(1, 13) = "" & Format(totalpvp, "Currency") ' .TextMatrix(1, 12)
       ' .TextMatrix(1, 12) = "" & Format(totaldcto, "Currency") ' .TextMatrix(1, 12)
       ' .TextMatrix(1, 14) = "" & Format(totiva, "Currency") ' .TextMatrix(1, 12)
       ' .TextMatrix(1, 15) = "" & Format(totalgeneral, "Currency") '.TextMatrix(1, 14)
       ' .TextMatrix(1, 16) = "Arreglos: " & Format(importe_arreglos, "Currency") '.TextMatrix(1, 14)
       ' .TextMatrix(1, 17) = "Real: " & Format(totalreal - importe_arreglos, "Currency") '.TextMatrix(1, 14)
       
       
    '  .TextMatrix(1, 18) = "Total(" & drc.RecordCount & ") Devoluciones: " & Format(tmpsubt * -1, "Currency")
                        
                        
    '.AutoSize 1, .Cols - 1
   ' .Redraw = True

  End With
   
   carga_grid_devoluciones = drc.RecordCount
   
   DoEvents

   On Error GoTo 0
   Exit Function

carga_grid_devoluciones_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_grid_devoluciones de Formulario frmFlexVen"
End Function


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
    DoEvents
  End With
  
  With ioFECHA
    .dspFormat = "dd/mm/yyyy"
   .LongMaxima = 10
   .Text = Date
  End With
  
  With ioFECHAINI
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
  End With
  
 With ioMODELO
   .LongMaxima = 30
 End With
  
 With ioREF
     .LongMaxima = 15
 End With
  
With cbCODPROV
    .ConexionString = locCnn
    .LenCodigo = 5
    .SQLString = "SELECT CODIGO, NOMBRE FROM MAPROV WHERE MBAJA = 0 ORDER BY CODIGO"
    .CodigoWidth = 800
    .carga
    DoEvents
End With

With cbFAMILIA
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM FAMILIAS WHERE MBAJA = 0 ORDER BY CODIGO"
    .LenCodigo = 5
    .CodigoWidth = 800
    .carga
End With

   With cbCATTALL
      .LenCodigo = 4
      .CodigoWidth = 670
      .ConexionString = locCnn
      .SQLString = "SELECT CODIGO, DESCRIPCION FROM CATTALL ORDER BY CODIGO"
      .carga
  End With
  
  With cbCOLOR
    .ConexionString = locCnn
    .LenCodigo = 3
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM COLORES WHERE (MBAJA = 0) AND (CODIGO > 0) ORDER BY CODIGO"
    .carga
    .CodigoWidth = 500
End With

With cbTALLA
    .ConexionString = locCnn
    .LenCodigo = 3
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM TALLAS WHERE (MBAJA = 0) AND (CODIGO > 0) ORDER BY CODIGO"
    .carga
    .CodigoWidth = 500
End With
  
With cbSUBFAM
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM SUBFAM WHERE MBAJA = 0 ORDER BY CODIGO"
    .LenCodigo = 5
    .CodigoWidth = 800
    .carga
End With

With cbSECCION
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM SECCIONES WHERE MBAJA = 0 ORDER BY CODIGO"
    .LenCodigo = 5
    .carga
    .CodigoWidth = 800
End With
 
With cbFCOBRO
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM FCOBRO WHERE MBAJA = 0 ORDER BY CODIGO"
    .LenCodigo = 1
    .CodigoWidth = 300
    .carga
End With

With cbCODPER
    .ConexionString = locCnn
    
    If TipoPermiso = 0 Then
        .SQLString = "SELECT CODIGO, NOMBRE FROM PERSONAL WHERE ((CODCEN = " & CentroActual & ") AND (MBAJA = 0) AND (TIPPERM = 0)) OR ((TIPPERM = 1) AND (MBAJA = 0)) ORDER BY TIPPERM, CODIGO"
    Else
        .SQLString = "SELECT CODIGO, NOMBRE FROM PERSONAL WHERE ((MBAJA = 0) AND (TIPPERM = 0)) OR ((TIPPERM = 1) AND (MBAJA = 0)) ORDER BY TIPPERM, CODIGO"
    End If
    
    .LenCodigo = 5
    .CodigoWidth = 800
    .carga
End With

'CAJA DE DESTINO (0=A, 1=B)
With cbEN
    .añade_item "0  A"
    .añade_item "1  B"
    .LenCodigo = 1
    .CodigoWidth = 300
End With

'0=PENDIENTE, 1=ACEPTADA, 2=CANCELADA
With cbESTADO
    .añade_item "0  PENDIENTE"
    .añade_item "1  ACEPTADA"
    .añade_item "2  CANCELADA"
    .LenCodigo = 1
    .CodigoWidth = 300
End With

With ioCODBAR
    .LongMaxima = LenCodBar
    .SoloNumeros = True
    .dspFormat = "0000000000000"
End With

With ioCODART
    .LongMaxima = 8
    .SoloNumeros = True
    .Alineacion = 1
End With

With cbTEMPOR
    .ConexionString = locCnn
    .SQLString = "SELECT IDTEM, ABREVIA FROM TEMPOR WHERE MBAJA = 0 ORDER BY ACTUAL DESC, IDTEM"
    .LenCodigo = 3
    .CodigoWidth = 500
    .DataField = "TEMPOR"
    .carga
End With

   Select Case TipoPermiso
   
   'usuario comun, ver solo los pedidos de su almacén
   Case 0
        cbCAJAS.Text = CajaActual
        cbCAJAS.Locked = True
        
   'supervisor, ver todos los pedidos
   Case 1
        cbCAJAS.Text = CajaActual
        cbCAJAS.Locked = False
           
   End Select
  
 artsql.AddTable "MAARTIC"
 artsql.AddField "(CONVERT(char(10), CODIGO) + CONVERT(char(3), TEMPOR))"
 
 DetVenSQL.AddTable "DETVENTA"
 
 'DetVenSQL.AddField "CONVERT(char(10), CODART) + CONVERT(char(3), TEMPOR) as COD"
 'DetVenSQL.AddField "CONVERT(char(10), CODVEN) + CONVERT(char(3), CODCAJA) as CODVENTA"

 DetVenSQL.AddField "CODVEN"
 DetVenSQL.AddField "CODCAJA"
 DetVenSQL.AddField "LINEA"
 DetVenSQL.AddField "CODART"
 DetVenSQL.AddField "TEMPOR"
 DetVenSQL.AddField "CODTALLA"
 DetVenSQL.AddField "CODCOL"
 DetVenSQL.AddField "UNIDADES"
 DetVenSQL.AddField "PREVEN"
 DetVenSQL.AddField "DCTO"
 DetVenSQL.AddField "IVA"
 DetVenSQL.AddField "RE"
 DetVenSQL.AddOrderClause "CODVEN"
 
 CabVenSQL.AddTable "CABVENTA"
 CabVenSQL.AddField "CONVERT(char(10), CODIGO) + CONVERT(char(3), CODCAJA)"
 'CabVenSQL.AddField "FCOBRO"
' CabVenSQL.AddField "FMODI"
 'CabVenSQL.AddField "CODPER"
 'CabVenSQL.AddField "CAJADES"
 'CabVenSQL.AddField "ESTADO"

 
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 '   tmpstrcombo = ""
    'Set nif = Nothing
        
    'si hemos establecido un filtro que no devuelve ningun registro,
    'borrar filtro para que no de error al volver al formulario
    'If miRc.EOF Then Call cbBorrar_click
    
    'No descargar desde aqui, descargar desde el formulario desde donde
    'se llame
    
    If miRc.State = 1 Then miRc.Close
    Set miRc = Nothing
    
    If Not seleccionado Then D_Cancelado = True
    
    If Not Desde_Devol Then Set frmFlexVen = Nothing
    Desde_Devol = False
        
End Sub

Private Sub ioCODBAR_GotFocus()
If Tab1.Tab <> 3 Then Tab1.Tab = 3
End Sub

'limpiar las otras fechas
Private Sub ioFECHA_Validate(Cancel As Boolean)
If ioFECHA.Text <> "" Then
    ioFECHAINI.Text = ""
    ioFECHAFIN.Text = ""
End If
End Sub

Private Sub ioFECHAFIN_GotFocus()
If Tab1.Tab <> 2 Then Tab1.Tab = 2
End Sub

Private Sub ioFECHAFIN_LostFocus()
ioCODBAR.SetFocus
End Sub

'si introduce finicio y ffin limpiar el campo de fecha
Private Sub ioFECHAINI_Validate(Cancel As Boolean)
    If ioFECHAINI.Text <> "" Then ioFECHA.Text = ""
End Sub


