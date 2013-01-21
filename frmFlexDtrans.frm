VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmFlexDtrans 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver Transferencias ..."
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11685
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
   ScaleWidth      =   11685
   StartUpPosition =   2  'CenterScreen
   Begin PCGestion.ucGrdBttn cbBorrar 
      Height          =   315
      Left            =   10935
      TabIndex        =   25
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
      Image           =   "frmFlexDtrans.frx":0000
   End
   Begin PCGestion.ucGrdBttn cbLista 
      Height          =   315
      Left            =   9690
      TabIndex        =   24
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
      Image           =   "frmFlexDtrans.frx":001C
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel4 
      Height          =   285
      Left            =   6195
      Top             =   1245
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
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
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   6015
      Left            =   15
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1530
      Visible         =   0   'False
      Width           =   11670
      _cx             =   20585
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
      FormatString    =   $"frmFlexDtrans.frx":0038
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
      Left            =   15
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   0
      Width           =   11655
      _ExtentX        =   20558
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
      TabPicture(0)   =   "frmFlexDtrans.frx":0116
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label14"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label16"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label17"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ioNUMPED"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ioCODTRN"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cbCODALMDEST"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ioMODELO"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "ioIMPORTE"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "ioREF"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "ioFECHA"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cbCODALMORIG"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Hoja 2"
      TabPicture(1)   =   "frmFlexDtrans.frx":0132
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
      TabPicture(2)   =   "frmFlexDtrans.frx":014E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label12"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label13"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label15"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label7"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label6"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cbCODPROV"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cbEN"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "cbESTADO"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "bsGradientLabel1"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "ioFECHAFIN"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "ioFECHAINI"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "Hoja 4"
      TabPicture(3)   =   "frmFlexDtrans.frx":016A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ioCODBAR"
      Tab(3).Control(1)=   "cbTEMPOR"
      Tab(3).Control(2)=   "ioCODART"
      Tab(3).Control(3)=   "cbTALLA"
      Tab(3).Control(4)=   "cbCATTALL"
      Tab(3).Control(5)=   "cbCOLOR"
      Tab(3).Control(6)=   "chameleonButton1"
      Tab(3).Control(7)=   "Label23"
      Tab(3).Control(8)=   "Label18"
      Tab(3).Control(9)=   "Label20"
      Tab(3).Control(10)=   "Label19"
      Tab(3).Control(11)=   "Label21"
      Tab(3).Control(12)=   "Label22"
      Tab(3).ControlCount=   13
      Begin PCGestion.miCombo cbCODALMORIG 
         Height          =   510
         Left            =   4080
         TabIndex        =   6
         Top             =   660
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
      Begin PCGestion.miText ioFECHA 
         Height          =   495
         Left            =   10215
         TabIndex        =   3
         Top             =   120
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
         Left            =   765
         TabIndex        =   0
         Top             =   120
         Width           =   2625
         _ExtentX        =   4630
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
         Left            =   8415
         TabIndex        =   2
         Top             =   120
         Width           =   1155
         _ExtentX        =   2037
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
         Left            =   4110
         TabIndex        =   1
         Top             =   120
         Width           =   3420
         _ExtentX        =   6033
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
      Begin PCGestion.miCombo cbSECCION 
         Height          =   525
         Left            =   -73155
         TabIndex        =   8
         Top             =   105
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
         Left            =   -67860
         TabIndex        =   9
         Top             =   120
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
         Left            =   -73170
         TabIndex        =   10
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
      Begin PCGestion.miCombo cbCODPER 
         Height          =   525
         Left            =   -67860
         TabIndex        =   11
         Top             =   615
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
         TabIndex        =   15
         Top             =   615
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
         TabIndex        =   16
         Top             =   615
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
      Begin PCGestion.bsGradientLabel bsGradientLabel1 
         Height          =   315
         Left            =   -67095
         Top             =   150
         Width           =   2475
         _ExtentX        =   4366
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
         Left            =   -71670
         TabIndex        =   14
         Top             =   600
         Width           =   2385
         _ExtentX        =   4207
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
      Begin PCGestion.miCombo cbEN 
         Height          =   525
         Left            =   -74085
         TabIndex        =   13
         Top             =   600
         Width           =   1620
         _ExtentX        =   2858
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
      Begin PCGestion.miCombo cbCODALMDEST 
         Height          =   525
         Left            =   8310
         TabIndex        =   7
         Top             =   660
         Width           =   3270
         _ExtentX        =   5768
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
      Begin PCGestion.miCombo cbCODPROV 
         Height          =   510
         Left            =   -74070
         TabIndex        =   12
         Top             =   90
         Width           =   4785
         _ExtentX        =   8440
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
      Begin PCGestion.miText ioCODTRN 
         Height          =   480
         Left            =   765
         TabIndex        =   4
         Top             =   660
         Width           =   1005
         _ExtentX        =   1773
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
      Begin PCGestion.miText ioNUMPED 
         Height          =   480
         Left            =   2370
         TabIndex        =   5
         Top             =   660
         Width           =   1020
         _ExtentX        =   1799
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
      Begin PCGestion.miText ioCODBAR 
         Height          =   480
         Left            =   -74505
         TabIndex        =   17
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
         Left            =   -68790
         TabIndex        =   19
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
         Left            =   -70980
         TabIndex        =   18
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
         Left            =   -74100
         TabIndex        =   21
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
         Left            =   -66270
         TabIndex        =   20
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
         Left            =   -70305
         TabIndex        =   22
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
      Begin PCGestion.chameleonButton chameleonButton1 
         Height          =   555
         Left            =   -66360
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
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
         MICON           =   "frmFlexDtrans.frx":0186
         PICN            =   "frmFlexDtrans.frx":01A2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CB"
         Height          =   330
         Left            =   -74925
         TabIndex        =   49
         Top             =   165
         Width           =   390
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TEMP."
         Height          =   285
         Left            =   -69420
         TabIndex        =   48
         Top             =   180
         Width           =   675
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CODIGO ART."
         Height          =   570
         Left            =   -71970
         TabIndex        =   47
         Top             =   30
         Width           =   960
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "COLOR"
         Height          =   330
         Left            =   -71100
         TabIndex        =   46
         Top             =   705
         Width           =   735
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TALLA"
         Height          =   330
         Left            =   -74820
         TabIndex        =   45
         Top             =   690
         Width           =   675
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CAT. TALLA"
         Height          =   555
         Left            =   -67125
         TabIndex        =   44
         Top             =   15
         Width           =   855
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PEDIDO"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   43
         Top             =   735
         Width           =   705
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CODIGO TRANSF."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -75
         TabIndex        =   42
         Top             =   585
         Width           =   885
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PROV."
         Height          =   330
         Left            =   -74730
         TabIndex        =   41
         Top             =   180
         Width           =   660
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DESTINO"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7410
         TabIndex        =   40
         Top             =   750
         Width           =   915
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "EN"
         Height          =   330
         Left            =   -74520
         TabIndex        =   39
         Top             =   690
         Width           =   405
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO"
         Height          =   330
         Left            =   -72510
         TabIndex        =   38
         Top             =   705
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA FINAL"
         Height          =   330
         Left            =   -66375
         TabIndex        =   37
         Top             =   660
         Width           =   1350
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA INICIAL"
         Height          =   330
         Left            =   -69270
         TabIndex        =   36
         Top             =   675
         Width           =   1530
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "USUARIO"
         Height          =   330
         Left            =   -68910
         TabIndex        =   35
         Top             =   690
         Width           =   1020
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SUBFAMILIA"
         Height          =   330
         Left            =   -74475
         TabIndex        =   34
         Top             =   675
         Width           =   1290
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SECCION"
         Height          =   330
         Left            =   -74100
         TabIndex        =   33
         Top             =   180
         Width           =   930
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FAMILIA"
         Height          =   330
         Left            =   -68715
         TabIndex        =   32
         Top             =   195
         Width           =   840
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "REF."
         Height          =   285
         Left            =   180
         TabIndex        =   31
         Top             =   210
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9465
         TabIndex        =   30
         Top             =   225
         Width           =   705
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ORIGEN"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3360
         TabIndex        =   29
         Top             =   735
         Width           =   675
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IMPORTE"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7470
         TabIndex        =   28
         Top             =   210
         Width           =   885
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "MODELO"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         TabIndex        =   27
         Top             =   225
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmFlexDtrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Módulo       : frmFlexDtrans
' Fecha/Hora : 03/03/2004 11:51
' Autor         : JCastillo
' Propósito    :  Mostrar detalle de transferencias ...
'---------------------------------------------------------------------------------------
Option Explicit

Dim first As Boolean

'Dim tmpstrcombo As String

Dim CabPtrnSQL As New clsSmartSQL
Dim DetPtrnSQL As New clsSmartSQL
Dim artsql As New clsSmartSQL

Dim miRc As New ADODB.Recordset
Dim seleccionado As Boolean

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
'Public Desde_Devol As Boolean
'Public D_Codart As Long
'Public D_Tempor As Byte
'Public D_CodTalla As Integer
'Public D_CodCol As Integer
'Public D_Unidades As Single
'Public D_Importe As Single
'Public D_Cancelado As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Sub cbBorrar_click()

ioREF.Text = ""
ioMODELO.Text = ""
ioIMPORTE.Text = ""
ioCODTRN.Text = ""
ioFECHA.Text = Date
cbCODALMORIG.Text = AlmacenActual
cbCODALMDEST.Text = ""
cbCODPROV.Text = ""
cbEN.Text = ""
cbSECCION.Text = ""
cbFAMILIA.Text = ""
cbSUBFAM.Text = ""
cbCODPER.Text = ""
ioFECHAINI.Text = ""
ioFECHAFIN.Text = ""
ioNUMPED.Text = ""

fg.Clear
fg.Rows = 1

ioREF.SetFocus

End Sub

Private Sub cbCATTALL_Validate(Cancel As Boolean)

   On Error GoTo cbCATTALL_Validate_Error

If cbCATTALL.Text = "" Then Exit Sub

With cbTALLA
    .ConexionString = locCnn
    .LenCodigo = 3
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM TALLAS WHERE (MBAJA = 0) AND (CODIGO > 0) AND CATTALL = " & cbCATTALL.Text & " ORDER BY CODIGO"
    .carga
    .CodigoWidth = 500
End With

   On Error GoTo 0
   Exit Sub

cbCATTALL_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbCATTALL_Validate de Formulario frmFlexDtrans"

End Sub

Private Sub cbCODPROV_GotFocus()

If Tab1.Tab <> 2 Then Tab1.Tab = 2

End Sub

Private Sub cbCOLOR_lostfocus()
Tab1.Tab = 0
ioREF.SetFocus
End Sub

Private Sub cbLista_click()
Dim usa_where As Boolean
Dim usa_artic As Boolean
Dim nuefech As String
Dim miCod As MiCodBar

   On Error GoTo cbLista_click_Error


  With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
  End With

CabPtrnSQL.ClearWhereClause
DetPtrnSQL.ClearWhereClause
artsql.ClearWhereClause


'si fechas estan en blanco, poner la de hoy por defecto
If (ioFECHA.Text = "") And (ioFECHAINI.Text = "") Then ioFECHA.Text = Date

'comprobar si quiere solo la fecha de hoy
If (ioFECHA.Text <> "") And IsDate(ioFECHA.Text) Then

    'CabPtrnSQL.AddSimpleWhereClause "FALTA", ioFECHA.Text, , CLAUSE_GREATERTHANOREQUAL
    'CabPtrnSQL.AddSimpleWhereClause "FALTA", CStr(DateAdd("d", 1, ioFECHA.Text)), , CLAUSE_LESSTHAN, LOGIC_AND
    'CabPtrnSQL.AddComplexWhereClause "Year(FALTA IN (" & DetPtrnSQL.SQL & ")", LOGIC_AND
    
    '>= q el dia actual
    '< que el dia siguiente
    nuefech = DateAdd("d", 1, ioFECHA.Text)
    CabPtrnSQL.AddComplexWhereClause "(FMODI >= '" & Format(Year((ioFECHA.Text)), "0000") & Format(Month((ioFECHA.Text)), "00") & Format(Day((ioFECHA.Text)), "00") & "' AND FMODI < '" & Format(Year((nuefech)), "0000") & Format(Month((nuefech)), "00") & Format(Day((nuefech)), "00") & "')", LOGIC_AND
    usa_where = True
   
'comprobar si quiere un rango de fechas

ElseIf (ioFECHAINI.Text <> "" And ioFECHAINI.Text <> "") And (IsDate(ioFECHAINI.Text) And IsDate(ioFECHAINI.Text)) Then
    Debug.Print Format(Day((ioFECHAFIN.Text)), "0000")
    CabPtrnSQL.AddComplexWhereClause "FMODI >= '" & Format(Year((ioFECHAINI.Text)), "0000") & Format(Month((ioFECHAINI.Text)), "00") & Format(Day((ioFECHAINI.Text)), "00") & "' AND FMODI <= '" & Format(Year((ioFECHAFIN.Text)), "0000") & Format(Month((ioFECHAFIN.Text)), "00") & Format(Day((ioFECHAFIN.Text)), "00") & "'", LOGIC_AND
    usa_where = True
End If

If cbESTADO.Text <> "" Then
    CabPtrnSQL.AddSimpleWhereClause "ESTADO", CLng(cbESTADO.Text), , , LOGIC_AND
    usa_where = True
End If

If ioCODBAR.Text <> "" Then

    miCod = Descompone_CBAR(ioCODBAR.Text)
    DetPtrnSQL.AddSimpleWhereClause "CODART", CLng(miCod.CODIGO_ART), , , LOGIC_AND
    DetPtrnSQL.AddSimpleWhereClause "TEMPOR", CByte(miCod.TEMPORADA_ART), , , LOGIC_AND
    DetPtrnSQL.AddSimpleWhereClause "CODTALLA", CInt(miCod.TALLA_ART), , , LOGIC_AND
    DetPtrnSQL.AddSimpleWhereClause "CODCOL", CInt(miCod.COLOR_ART), , , LOGIC_AND
    
End If

If ioCODART.Text <> "" Then
    DetPtrnSQL.AddSimpleWhereClause "CODART", CLng(ioCODART.Text), , , LOGIC_AND
End If

If cbTEMPOR.Text <> "" Then
    DetPtrnSQL.AddSimpleWhereClause "TEMPOR", CLng(cbTEMPOR.Text), , , LOGIC_AND
End If

If cbTALLA.Text <> "" Then
    DetPtrnSQL.AddSimpleWhereClause "CODTALLA", CLng(cbTALLA.Text), , , LOGIC_AND
End If

If cbCOLOR.Text <> "" Then
    DetPtrnSQL.AddSimpleWhereClause "CODCOL", CLng(cbCOLOR.Text), , , LOGIC_AND
End If

If ioNUMPED.Text <> "" Then
    CabPtrnSQL.AddSimpleWhereClause "NUMPED", CLng(ioNUMPED.Text), , , LOGIC_AND
    usa_where = True
End If

If (ioCODTRN.Text <> "") And (cbCODALMORIG.Text <> "") Then
    CabPtrnSQL.AddComplexWhereClause "((CODIGO = " & ioCODTRN.Text & ") AND (CODALMORIG = " & cbCODALMORIG.Text & "))", LOGIC_AND
    usa_where = True
End If
 
If ioIMPORTE.Text <> "" Then
    If CDbl(ioIMPORTE.Text) > 0 Then
        artsql.AddSimpleWhereClause "PRECOM", CDbl(ioIMPORTE.Text), , , LOGIC_AND
        usa_artic = True
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

If cbCODALMORIG.Text <> "" Then
    CabPtrnSQL.AddSimpleWhereClause "CODALMORIG", cbCODALMORIG.Text, , , LOGIC_AND
    usa_where = True
End If

If cbCODALMDEST.Text <> "" Then
    CabPtrnSQL.AddSimpleWhereClause "CODALMDEST", cbCODALMDEST.Text, , , LOGIC_AND
    usa_where = True
End If

If cbCODPER.Text <> "" Then
    CabPtrnSQL.AddSimpleWhereClause "CODUSR", cbCODPER.Text, , , LOGIC_AND
    usa_where = True
End If

'EN, A o B
If cbEN.Text <> "" Then

If cbEN.Text = "0" Then
        artsql.AddSimpleWhereClause "IVACOM", 0, , CLAUSE_GREATERTHAN, LOGIC_AND
        usa_artic = True
    'mostrar aquellos que estan en B (IVACOM = 0)
Else
        artsql.AddSimpleWhereClause "IVACOM", 0, , CLAUSE_EQUALS, LOGIC_AND
        usa_artic = True
End If

End If

If cbESTADO.Text <> "" Then
    CabPtrnSQL.AddSimpleWhereClause "ESTADO", cbESTADO.Text, , , LOGIC_AND
    usa_where = True
End If

'si deja todo en blanco, no mostrar ningun registro
If (Not usa_where) And (Not usa_artic) Then
    fg.Clear
    ioREF.SetFocus
    Exit Sub
End If


If usa_artic Then DetPtrnSQL.AddComplexWhereClause "(CONVERT(char(10), CODART) + CONVERT(char(3), TEMPOR)) IN (" & artsql.SQL & ")", LOGIC_AND

'If usa_artic Then DetPtrnSQL.AddComplexWhereClause "COD IN (" & ArtSql.SQL & ")", LOGIC_AND
If usa_where Then DetPtrnSQL.AddComplexWhereClause "(CONVERT(char(10), CODIGO) + CONVERT(char(3), CODALM)) IN (" & CabPtrnSQL.SQL & ")", LOGIC_AND

'If ioNOMBRE.Text <> "" Then CabPtrnSQL.AddComplexWhereClause "CODCOST IN (" & DetPtrnSQL.SQL & ")", LOGIC_AND

If miRc.State = 1 Then miRc.Close
miRc.Open DetPtrnSQL.SQL, locCnn, adOpenStatic, adLockOptimistic

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

   On Error GoTo carga_grid_Error
   
   With fg
   
    .Clear
    .Cols = 22
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
    .TextMatrix(0, 10) = "Pre.Com."
    .TextMatrix(0, 11) = "PVP"
    .TextMatrix(0, 12) = "Dcto"
    .TextMatrix(0, 13) = "Sub."
    .TextMatrix(0, 14) = "R.E."
    .TextMatrix(0, 15) = "Total Com."
    .TextMatrix(0, 16) = "Codigo"
    .TextMatrix(0, 17) = "Pedido"
    .TextMatrix(0, 18) = "Origen"
    .TextMatrix(0, 19) = "Destino"
    .TextMatrix(0, 20) = "En"
    .TextMatrix(0, 21) = "CBarras"
        
    If miRc.RecordCount <= 0 Then Exit Sub
        
    Do
             .Rows = .Rows + 1
    
        If Not miRc.EOF Then
     
            conta_lineas = conta_lineas + 1
         
            'ID
            .TextMatrix(.Rows - 1, 0) = conta_lineas
            'CODALMORIG
            .TextMatrix(.Rows - 1, 1) = miRc.fields("CODALM")
            
             T_Cabecera = devuelve_matriz("SELECT CODALMDEST, DCTO, NUMPED FROM PTRANS WHERE CODALMORIG = " & miRc.fields("CODALM") & " AND CODIGO = " & miRc.fields("CODIGO"), locCnn)
            
            'FECHA (cogerla de la cabecera)
            .TextMatrix(.Rows - 1, 2) = miRc.fields("FMODI")
            
            t_articulo = devuelve_matriz("SELECT MODELO, REF, CODPROV, PRECOM, IVACOM, PREVEN FROM MAARTIC WHERE CODIGO = " & miRc.fields("CODART") & " AND TEMPOR = " & miRc.fields("TEMPOR"), locCnn)
            Dsp_Prov = Trim(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & t_articulo(2), locCnn))
             
            'Proveedor
            .TextMatrix(.Rows - 1, 3) = Dsp_Prov
            
            'Referencia
            .TextMatrix(.Rows - 1, 4) = Trim(t_articulo(1))
            
            'MODELO
            .TextMatrix(.Rows - 1, 5) = Format(miRc.fields("CODART"), "00000") & "-" & Trim(t_articulo(0))
            
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
            'precio de compra
            .TextMatrix(.Rows - 1, 10) = t_articulo(3)
            
            'precio de venta
            .TextMatrix(.Rows - 1, 11) = t_articulo(5)
            
            'DCTO
            .TextMatrix(.Rows - 1, 12) = T_Cabecera(1) & " %"
            
            'SUBT
            .TextMatrix(.Rows - 1, 13) = (t_articulo(3) - ((t_articulo(3) * T_Cabecera(1)) / 100)) * miRc.fields("UNIDADES")
            
            'RE (solo mostrar a los supervisores)
            If TipoPermiso = 1 Then
            .TextMatrix(.Rows - 1, 14) = miRc.fields("RE") & " %"
            'Else
            '.TextMatrix(.Rows - 1, 14) = ""
            End If
            
           'Total
           .TextMatrix(.Rows - 1, 15) = .TextMatrix(.Rows - 1, 13) + ((.TextMatrix(.Rows - 1, 13) * miRc.fields("RE")) / 100)
            
           'codigo
           .TextMatrix(.Rows - 1, 16) = miRc.fields("CODIGO")
           
           'numero de pedido
           .TextMatrix(.Rows - 1, 17) = T_Cabecera(2)
             
           'origen
           .TextMatrix(.Rows - 1, 18) = Trim(devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & miRc.fields("CODALM"), locCnn))
              
           'destino
           .TextMatrix(.Rows - 1, 19) = Trim(devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & T_Cabecera(0), locCnn))
              
           'en
           If t_articulo(4) > 0 Then
           .TextMatrix(.Rows - 1, 20) = "A"
           Else
           .TextMatrix(.Rows - 1, 20) = "B"
           End If
           
        '12345 678 90 123
        '00000-000-00-000
        .TextMatrix(.Rows - 1, 21) = Conforma_CB(miRc.fields("CODART"), miRc.fields("TEMPOR"), miRc.fields("CODTALLA"), miRc.fields("CODCOL"))
             
      
                        
           'estado
           '.TextMatrix(.Rows - 1, 10) = miRc.Fields("ESTADO")
            
        End If
    
    If Not miRc.EOF Then miRc.MoveNext
    
    'DoEvents
    
    Loop Until miRc.EOF
          
        .SubtotalPosition = flexSTAbove
        .subtotal flexSTCount, -1, 6, , vbBlue, vbWhite, True
        .subtotal flexSTSum, -1, 9, , vbBlue, vbWhite, True
        .subtotal flexSTSum, -1, 10, , vbBlue, vbWhite, True
      '  .subtotal flexSTSum, -1, 13, , vbBlue, vbWhite, True
        .subtotal flexSTSum, -1, 15, , vbBlue, vbWhite, True
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 5) = "Nº Artículos: (" & .TextMatrix(1, 6) & ")"
        .TextMatrix(1, 6) = ""
        .TextMatrix(1, 9) = "Uds: " & Round(.TextMatrix(1, 9))
        '.TextMatrix(1, 13) = "Sub: " & .TextMatrix(1, 13)
        .TextMatrix(1, 15) = "Total: " & .TextMatrix(1, 15)
        
    .AutoSize 1, .Cols - 1
    .Redraw = True

  End With
   
   DoEvents
   
   On Error GoTo 0
   Exit Sub

carga_grid_Error:
   
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_grid de Formulario frmFlexArre"
 
 End Sub

Private Sub cbSECCION_GotFocus()
If Tab1.Tab <> 1 Then Tab1.Tab = 1
End Sub

Private Sub chameleonButton1_Click()

Dim linea1 As String
Dim linea2 As String
Dim tmpalm As String
         
   
   On Error GoTo chameleonButton1_Click_Error

    DoEvents

    If cbCODALMORIG.Text <> "" Then
        tmpalm = devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & cbCODALMORIG.Text, locCnn)
        If tmpalm = "@" Then tmpalm = ""
    End If
    
    linea1 = "Transferencias:  F.Inicial: " & ioFECHA.Text & ". F.Final: " & ioFECHAFIN.Text & ". Almacén: " & tmpalm
    linea2 = "Usuario: " & Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & UsuarioActual, locCnn)) & ". Fecha: " & Now
    
    Call PrintFlexGrid(fg, 1, 1, 2, linea1, linea2, 13, 2, 10)
    'fg.SaveGrid "c:\prueba.txt", flexFileCommaText
    'fg.PrintGrid "Transferencia " & rc.Fields("CODIGO"), , 2, 600, 600
    
    
    'actualizar los colores del grid, volivendolo a cargar
    'rc.Move 0


   On Error GoTo 0
   Exit Sub

chameleonButton1_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento chameleonButton1_Click de Formulario frmFlexDtrans"

End Sub

Private Sub fg_DblClick()

If miRc.State = 0 Then Exit Sub
If miRc.RecordCount <= 0 Then Exit Sub

    seleccionado = True
    
If fg.Rows <= 1 Then Exit Sub
    
If IsNumeric(fg.TextMatrix(fg.Row, 0)) Then
        'posicionarse en el registro
        miRc.Move (fg.TextMatrix(fg.Row, 0) - 1), 1
        
End If
        
      '  If Desde_Devol Then
     '
    '        'preguntar al usuario ...
    '        If MsgBox("¿Desea introducir el artículo actual como DEVOLUCION? " & Chr(13) & _
    '        "Modelo: " & fg.TextMatrix(fg.Row, 3) & Chr(13) & _
    '        "Talla: " & fg.TextMatrix(fg.Row, 4) & Chr(13) & _
    '        "Color: " & fg.TextMatrix(fg.Row, 5) & Chr(13) & _
    '        "Importe: " & fg.TextMatrix(fg.Row, 12) & Chr(13) & _
    '        "(se devolverán las unidades al almacén)", vbQuestion + vbYesNo, titulo) = vbYes Then
                
    '       'ASIGNAR VARIABLES:
    '       D_Codart = miRc.fields("CODART")
    '       D_Tempor = miRc.fields("TEMPOR")
    '       D_CodTalla = miRc.fields("CODTALLA")
    '       D_CodCol = miRc.fields("CODCOL")
    '       D_Unidades = miRc.fields("UNIDADES")
    '       D_Importe = CDbl(fg.TextMatrix(fg.Row, 12))
    '       D_Cancelado = False
    '
    '       miRc.Close
    '
    '       DoEvents
    '
    '       Unload Me
    '
    '       Exit Sub
    '
    '       End If
           
           
        
        'si no entramos desde devoluciones
    '    Else
    '
    '       'presentar la venta en el formulario de ventas y permitir su modificación.
    '
    '       Me.WindowState = 1
    '
    '       With frmCabVen
    '        .Desde_Grid = True
    '        .D_Codigo = miRc.fields("CODVEN")
    '        .D_CodCaja = miRc.fields("CODCAJA")
    '        .Show
    '       End With
    '       DoEvents
           
           'miRc.Close
           
'           Unload Me
                         
    '    End If
    '
    '    DoEvents
                
'End If
  
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
  
  With cbCODALMORIG
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM CAJAS WHERE MBAJA = 0 ORDER BY CODIGO"
    .LenCodigo = 1
    .CodigoWidth = 500
    .carga
    .Refresh
    DoEvents
  End With
  
  With cbCODALMDEST
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM CAJAS WHERE MBAJA = 0 ORDER BY CODIGO"
    .LenCodigo = 1
    .CodigoWidth = 500
    .carga
    .Refresh
    DoEvents
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
  
  With ioCODBAR
    .LongMaxima = LenCodBar
    .SoloNumeros = True
    .dspFormat = "0000000000000"
  End With
  
  With ioFECHA
    .dspFormat = "dd/mm/yyyy"
   .LongMaxima = 10
   .Text = Date
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
 
With cbCODPER
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, NOMBRE FROM PERSONAL WHERE ((CODCEN = " & CentroActual & ") AND (MBAJA = 0) AND (TIPPERM = 0)) OR ((TIPPERM = 1) AND (MBAJA = 0)) ORDER BY TIPPERM, CODIGO"
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

With ioNUMPED
    .Alineacion = 1
    .SoloNumeros = True
    .LongMaxima = 10
End With

'0=EN CREACION, 1=PENDIENTE, 2=ACEPTADO, 3=CANCELADO
With cbESTADO
    .añade_item "0  EN CREACION", 1
    .añade_item "1  PENDIENTE", 2
    .añade_item "2  ACEPTADA", 3
    .añade_item "3  CANCELADA", 4
    .LenCodigo = 1
    .CodigoWidth = 300
End With

   Select Case TipoPermiso
   
   'usuario comun, ver solo los pedidos de su almacén
   Case 0
        cbCODALMORIG.Text = AlmacenActual
        cbCODALMORIG.Locked = True
        
   'supervisor, ver todos los pedidos
   Case 1
        cbCODALMORIG.Text = AlmacenActual
        cbCODALMORIG.Locked = False
           
   End Select
  
 artsql.AddTable "MAARTIC"
 artsql.AddField "(CONVERT(char(10), CODIGO) + CONVERT(char(3), TEMPOR))"
 
 DetPtrnSQL.AddTable "DETTRANS"
 
 'DetPtrnSQL.AddField "CONVERT(char(10), CODART) + CONVERT(char(3), TEMPOR) as COD"
 'DetPtrnSQL.AddField "CONVERT(char(10), CODVEN) + CONVERT(char(3), CODCAJA) as CODVENTA"

 DetPtrnSQL.AddField "CODIGO"
 DetPtrnSQL.AddField "CODALM"
 DetPtrnSQL.AddField "ID"
 DetPtrnSQL.AddField "CODART"
 DetPtrnSQL.AddField "TEMPOR"
 DetPtrnSQL.AddField "CODTALLA"
 DetPtrnSQL.AddField "CODCOL"
 DetPtrnSQL.AddField "UNIDADES"
 DetPtrnSQL.AddField "RE"
 DetPtrnSQL.AddField "FMODI"
 
 CabPtrnSQL.AddTable "PTRANS"
 CabPtrnSQL.AddField "CONVERT(char(10), CODIGO) + CONVERT(char(3), CODALMORIG)"
 'CabPtrnSQL.AddField "FCOBRO"
' CabPtrnSQL.AddField "FMODI"
 'CabPtrnSQL.AddField "CODPER"
 'CabPtrnSQL.AddField "CAJADES"
 'CabPtrnSQL.AddField "ESTADO"

 
  
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
    
    'If Not seleccionado Then D_Cancelado = True
    
    'If Not Desde_Devol Then Set frmFlexVen = Nothing
    'Desde_Devol = False
    Set frmFlexDtrans = Nothing
    
        
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




'si introduce finicio y ffin limpiar el campo de fecha
Private Sub ioFECHAINI_Validate(Cancel As Boolean)
    If ioFECHAINI.Text <> "" Then ioFECHA.Text = ""
End Sub



