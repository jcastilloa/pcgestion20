VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmPedProv 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedidos a Proveedores"
   ClientHeight    =   7380
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   10785
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   10785
   Begin TabDlg.SSTab Tab1 
      Height          =   1380
      Left            =   1800
      TabIndex        =   37
      Top             =   15
      Width           =   8970
      _ExtentX        =   15822
      _ExtentY        =   2434
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "Hoja 1"
      TabPicture(0)   =   "frmPedProv.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label19"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cbListaProv"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cbPLAZOE"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cBFPAGO"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cbCODPROV"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ioFECHA"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Hoja 2"
      TabPicture(1)   =   "frmPedProv.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ioDCTOPP"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cbCODALM"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "ioPORTES"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "ioGASTOS"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "ioTRNSPORTI"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "ioALBARAN"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "ioFACTURA"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "ioSUCODIGO"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label21"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label20"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label18"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label4"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label7"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label6"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label10"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label3"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).ControlCount=   16
      Begin PCGestion.miText ioFECHA 
         Height          =   495
         Left            =   705
         TabIndex        =   0
         Top             =   360
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9.75
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
         Height          =   495
         Left            =   3150
         TabIndex        =   1
         Top             =   360
         Width           =   5370
         _ExtentX        =   10213
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
      Begin PCGestion.miCombo cBFPAGO 
         Height          =   480
         Left            =   690
         TabIndex        =   2
         Top             =   840
         Width           =   3840
         _ExtentX        =   6773
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
      Begin PCGestion.miCombo cbPLAZOE 
         Height          =   480
         Left            =   5280
         TabIndex        =   3
         Top             =   840
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   847
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin PCGestion.miText ioDCTOPP 
         Height          =   525
         Left            =   -70245
         TabIndex        =   5
         Top             =   360
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9.75
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
      Begin PCGestion.miCombo cbCODALM 
         Height          =   480
         Left            =   -74115
         TabIndex        =   4
         Top             =   345
         Width           =   3300
         _ExtentX        =   5821
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
      Begin PCGestion.miText ioPORTES 
         Height          =   480
         Left            =   -68805
         TabIndex        =   6
         Top             =   360
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   847
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9.75
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
      Begin PCGestion.miText ioGASTOS 
         Height          =   525
         Left            =   -67080
         TabIndex        =   7
         Top             =   360
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9.75
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
      Begin PCGestion.miText ioTRNSPORTI 
         Height          =   525
         Left            =   -67635
         TabIndex        =   11
         Top             =   840
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9.75
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
      Begin PCGestion.miText ioALBARAN 
         Height          =   525
         Left            =   -71910
         TabIndex        =   9
         Top             =   840
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9.75
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
         Left            =   -69885
         TabIndex        =   10
         Top             =   840
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9.75
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
      Begin PCGestion.miText ioSUCODIGO 
         Height          =   525
         Left            =   -74100
         TabIndex        =   8
         Top             =   840
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9.75
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
      Begin PCGestion.chameleonButton cbListaProv 
         Height          =   390
         Left            =   8550
         TabIndex        =   55
         TabStop         =   0   'False
         ToolTipText     =   "Seleccionar Proveedor ..."
         Top             =   405
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   688
         BTYPE           =   3
         TX              =   "Prv"
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
         MICON           =   "frmPedProv.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA FACTURA"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   -74925
         TabIndex        =   49
         Top             =   780
         Width           =   825
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FACTURA"
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
         Left            =   -70710
         TabIndex        =   48
         Top             =   945
         Width           =   825
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ALBARAN"
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
         Left            =   -72735
         TabIndex        =   47
         Top             =   930
         Width           =   825
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TRANSPOR."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -68700
         TabIndex        =   46
         Top             =   945
         Width           =   1080
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "GASTOS"
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
         Left            =   -67770
         TabIndex        =   45
         Top             =   435
         Width           =   705
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PORTES"
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
         Left            =   -69555
         TabIndex        =   44
         Top             =   435
         Width           =   765
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ALMACEN"
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
         Left            =   -75015
         TabIndex        =   43
         Top             =   465
         Width           =   915
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DCTO PP %"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   -70815
         TabIndex        =   42
         Top             =   315
         Width           =   570
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
         Height          =   240
         Left            =   120
         TabIndex        =   41
         Top             =   435
         Width           =   585
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PROVEEDOR"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2085
         TabIndex        =   40
         Top             =   450
         Width           =   1065
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FORMA PAGO"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -45
         TabIndex        =   39
         Top             =   780
         Width           =   765
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PLAZO    ENT."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   4515
         TabIndex        =   38
         Top             =   795
         Width           =   885
      End
   End
   Begin TabDlg.SSTab Tab2 
      Height          =   4005
      Left            =   30
      TabIndex        =   24
      Top             =   1110
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   7064
      _Version        =   393216
      Style           =   1
      Tabs            =   2
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
      TabCaption(0)   =   "Artículos"
      TabPicture(0)   =   "frmPedProv.frx":0054
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fg"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Totales"
      TabPicture(1)   =   "frmPedProv.frx":0070
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label11"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label12"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label14"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label16"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label17"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label23"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label24"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "bsGradientLabel1"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lblIVA"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lblSubTotal"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lblTotal"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lblGastos"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lblPortes"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "vsTotales"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).ControlCount=   15
      Begin VSFlex8LCtl.VSFlexGrid vsTotales 
         Height          =   2745
         Left            =   -74655
         TabIndex        =   29
         ToolTipText     =   "Doble Click para ajustar el IVA desde Sub Total"
         Top             =   900
         Width           =   6105
         _cx             =   10769
         _cy             =   4842
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
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
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
         FormatString    =   $"frmPedProv.frx":008C
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
      Begin VSFlex8Ctl.VSFlexGrid fg 
         Height          =   3615
         Left            =   30
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Doble Click para editar los artículos del pedido"
         Top             =   360
         Width           =   10680
         _cx             =   18838
         _cy             =   6376
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
         Rows            =   1
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPedProv.frx":0161
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
      Begin PCGestion.bsGradientLabel lblPortes 
         Height          =   465
         Left            =   -66210
         Top             =   1770
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   820
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
         CaptionColour   =   192
         Colour1         =   14737632
         Colour2         =   12632256
         CaptionAlignment=   2
      End
      Begin PCGestion.bsGradientLabel lblGastos 
         Height          =   465
         Left            =   -66210
         Top             =   2310
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   820
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
         CaptionColour   =   192
         Colour1         =   14737632
         Colour2         =   12632256
         CaptionAlignment=   2
      End
      Begin PCGestion.bsGradientLabel lblTotal 
         Height          =   780
         Left            =   -67065
         Top             =   2850
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   1376
         Caption         =   ""
         BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   18.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionColour   =   16711680
         Colour1         =   14737632
         Colour2         =   12632256
         CaptionAlignment=   2
      End
      Begin PCGestion.bsGradientLabel lblSubTotal 
         Height          =   465
         Left            =   -66540
         Top             =   705
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   820
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
         Colour1         =   14737632
         Colour2         =   12632256
         CaptionAlignment=   2
      End
      Begin PCGestion.bsGradientLabel lblIVA 
         Height          =   465
         Left            =   -66210
         Top             =   1230
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   820
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
         CaptionColour   =   192
         Colour1         =   14737632
         Colour2         =   12632256
         CaptionAlignment=   2
      End
      Begin PCGestion.bsGradientLabel bsGradientLabel1 
         Height          =   345
         Left            =   -74655
         Top             =   540
         Width           =   6090
         _ExtentX        =   10742
         _ExtentY        =   609
         Caption         =   "NETO : SUBTOTAL - DCTO - DCTOPP + IVA + GASTOS + PORTES"
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
         Colour1         =   15640462
         Colour2         =   7177785
         CaptionAlignment=   1
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IVA"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68145
         TabIndex        =   54
         Top             =   1320
         Width           =   870
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   735
         Left            =   -66690
         TabIndex        =   53
         Top             =   1110
         Width           =   435
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   735
         Left            =   -66690
         TabIndex        =   35
         Top             =   2175
         Width           =   435
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   735
         Left            =   -66690
         TabIndex        =   34
         Top             =   1650
         Width           =   435
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NETO"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68325
         TabIndex        =   33
         Top             =   810
         Width           =   1080
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68100
         TabIndex        =   32
         Top             =   3090
         Width           =   870
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "GASTOS"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68100
         TabIndex        =   31
         Top             =   2415
         Width           =   870
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PORTES"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68100
         TabIndex        =   30
         Top             =   1860
         Width           =   870
      End
   End
   Begin PCGestion.chameleonButton cmdPrevious 
      Height          =   600
      Left            =   1065
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5940
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
      MICON           =   "frmPedProv.frx":023F
      PICN            =   "frmPedProv.frx":025B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cmdFirst 
      Height          =   600
      Left            =   30
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5940
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
      MICON           =   "frmPedProv.frx":0F2D
      PICN            =   "frmPedProv.frx":0F49
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
      Height          =   600
      Left            =   8625
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5940
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
      MICON           =   "frmPedProv.frx":1C7F
      PICN            =   "frmPedProv.frx":1C9B
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
      Height          =   600
      Left            =   9705
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5940
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
      MICON           =   "frmPedProv.frx":296D
      PICN            =   "frmPedProv.frx":2989
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
      Height          =   735
      Left            =   30
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6600
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
      MICON           =   "frmPedProv.frx":36BF
      PICN            =   "frmPedProv.frx":36DB
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
      Height          =   735
      Left            =   1125
      TabIndex        =   12
      Top             =   6600
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
      MICON           =   "frmPedProv.frx":43B5
      PICN            =   "frmPedProv.frx":43D1
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
      Height          =   735
      Left            =   2355
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6600
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
      MICON           =   "frmPedProv.frx":4CAB
      PICN            =   "frmPedProv.frx":4CC7
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
      Height          =   735
      Left            =   7620
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6600
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
      MICON           =   "frmPedProv.frx":5525
      PICN            =   "frmPedProv.frx":5541
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbImprimir 
      Height          =   735
      Left            =   8595
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6600
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "&Imprimir"
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
      MICON           =   "frmPedProv.frx":5E1B
      PICN            =   "frmPedProv.frx":5E37
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
      Height          =   735
      Left            =   9705
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6600
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
      MICON           =   "frmPedProv.frx":6B11
      PICN            =   "frmPedProv.frx":6B2D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton dtAgregar 
      Height          =   390
      Left            =   30
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Añadir un nuevo artículo"
      Top             =   5145
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1032
      BTYPE           =   9
      TX              =   "Agregar"
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
      MICON           =   "frmPedProv.frx":7807
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton dtEdicion 
      Height          =   390
      Left            =   780
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Editar linea de pedido"
      Top             =   5145
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1032
      BTYPE           =   9
      TX              =   "Edición"
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
      MICON           =   "frmPedProv.frx":7823
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblstatus 
      Height          =   345
      Left            =   30
      Top             =   5565
      Width           =   10725
      _ExtentX        =   16854
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
   Begin PCGestion.bsGradientLabel lblStatusD 
      Height          =   375
      Left            =   3120
      Top             =   5145
      Width           =   3075
      _ExtentX        =   5424
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
   Begin PCGestion.chameleonButton cbSeleccionaArticulo 
      Height          =   390
      Left            =   10035
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "Añadir artículo existente"
      Top             =   5145
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   688
      BTYPE           =   9
      TX              =   "Buscar"
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
      MICON           =   "frmPedProv.frx":783F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.miCombo cbESTADO 
      Height          =   480
      Left            =   7065
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5100
      Width           =   2955
      _ExtentX        =   5212
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
   Begin PCGestion.bsGradientLabel lblTemporadaWrk 
      Height          =   375
      Left            =   3930
      ToolTipText     =   "Doble Click para cambiar la temporada de trabajo actual"
      Top             =   6045
      Width           =   2655
      _ExtentX        =   4683
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
      Colour1         =   16777152
      Colour2         =   16744576
      CaptionAlignment=   1
   End
   Begin PCGestion.bsGradientLabel lblExisteCom 
      Height          =   330
      Left            =   45
      Top             =   405
      Visible         =   0   'False
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   661
      Caption         =   "Comentario"
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
      Colour1         =   8454143
      Colour2         =   49152
      CaptionAlignment=   1
   End
   Begin PCGestion.ucGrdBttn cmComentario 
      Height          =   345
      Left            =   30
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   750
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   609
      Caption         =   "Modificar Comentario"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "frmPedProv.frx":785B
   End
   Begin PCGestion.miCombo Temporada 
      Height          =   480
      Left            =   4755
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   6435
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   847
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
   Begin PCGestion.chameleonButton btA 
      Height          =   405
      Left            =   6210
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   5130
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "A"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedProv.frx":7877
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton btB 
      Height          =   405
      Left            =   6645
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   5130
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "B"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedProv.frx":7893
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cmEntrada 
      Height          =   405
      Left            =   2145
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   6015
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "&Etiquetas"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedProv.frx":78AF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cmDesacerLinea 
      Height          =   405
      Left            =   6615
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   6015
      Visible         =   0   'False
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "&Deshacer Linea"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedProv.frx":78CB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cmHistorico 
      Height          =   405
      Left            =   3990
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   6930
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "&Pasar a Histórico"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedProv.frx":78E7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cmCambTem 
      Height          =   390
      Left            =   7170
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   6480
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   688
      BTYPE           =   3
      TX              =   "C"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPedProv.frx":7903
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cmEditarPorRef 
      Height          =   390
      Left            =   1530
      TabIndex        =   62
      TabStop         =   0   'False
      ToolTipText     =   "Editar linea de pedido, buscando por referencia de artículo"
      Top             =   5145
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   688
      BTYPE           =   9
      TX              =   "Buscar Referencia"
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
      MICON           =   "frmPedProv.frx":791F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TEMPORADAS"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3345
      TabIndex        =   52
      Top             =   6510
      Width           =   1380
   End
   Begin VB.Label ioNUMERO 
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
      Left            =   465
      TabIndex        =   36
      Top             =   30
      Width           =   1305
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NUM"
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
      Left            =   0
      TabIndex        =   25
      Top             =   60
      Width           =   435
   End
End
Attribute VB_Name = "frmPedProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Módulo     : frmPedProv
' Fecha/Hora : 04/08/2004 16:40
' Autor      : JCastillo
' Propósito  :
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' Module     : frmPedProv
' DateTime : 10/10/2003 17:40
' Author     :  José Castillo
' Purpose   : Pedidos a Proveedores.
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' Notas:
'---------------------------------------------------------------------------------------
' Para esta versión se utiliza un recordset jerarquico.
' Primario (Maestro)    :  adoPrimaryRS
' Secundario (Detalle) :  adoDetalleRS

' poner el código de validación de los campos en detalle en:
'
' * fg_ValidateEdit (evento del grid)
'
' y si fuera necesario, también poner código en:
' * adoDetalleRS_WillChangeRecord (evento del recordset)

' poner replaces (.EditText = Replace(.EditText, ",", ".")) para aquellos campos
' que sean con decimales. Tambien para poner solo_numericos:

' * fg_ChangeEdit
'
' Poner los DataSource, DataFields y miCombos en
' Private Sub Asigna_Campos()

' Los formatos y ColCombos para el grid poner en
' Private Sub Asigna_Grid()
'---------------------------------------------------------------------------------------
' ESTADOS DEL PEDIDO:
' 1 -> en creacion
' 2 -> parcial
' 3 -> en historico
' 4 -> transferido

Option Explicit

Public WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Public WithEvents adoDetalleRS As Recordset
Attribute adoDetalleRS.VB_VarHelpID = -1

Dim creando As Boolean

Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim editagrid As Boolean

Dim cabSQL As String

Dim lblValue As String
Dim lblOriginalValue As String
Dim Casignados As Boolean
Dim Gasignados As Boolean

Dim NoAvisar As Boolean

'Dim se_ha_desecho As Boolean

'para almacenar el IDTEM de la temporada de trabajo.
'Dim tmptemporada As Long
Dim oSQL As New clsSmartSQL
Dim sqlSHAPE As String


Dim tmptempor As String
'Dim tmpcodigo As String
Dim tmptalla As String
Dim tmpcolor As String

'subtotal (sin IVA)
Dim tmpsubtot As Double

'Temporada de trabajo para los pedidos (puede ser distinta a la temporada
'actual
Public TemporadaWrk As Byte

Public Linea_Creada As Long

'-----------------------------------------------------------------------------------------------
'si venimos del formulario TrabaDatos, mostrar el numero de pedido seleccionado
'para trabajar con el
Public trabajar_con_pedido As Boolean
Public NUMERO_PEDIDO As Long
Public codigo_almacen As Byte
'-----------------------------------------------------------------------------------------------


Private Sub adoDetalleRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
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
' Procedure : Asigna_Campos
' DateTime  : 22/10/2003 21:38
' Author    : Administrador
' Purpose   : Configura los Datafield y Datasources de los campos
'             asi como los miCombos si fueran necesarios
'---------------------------------------------------------------------------------------
Private Sub Asigna_Campos()
   
  With ioNUMERO
    Set .DataSource = adoPrimaryRS
    .DataField = "NUMERO"
   .Alignment = 2
    '.Alineacion = 1
   ' .Locked = True  'que entre como bloqueado y solo desbloquear
                    'para añadir un nuevo registro
End With
  
With ioFECHA
    .LongMaxima = 10
    Set .DataSource = adoPrimaryRS
    .dspFormat = "dd/mm/yyyy"
    .DataField = "FECHA"
End With
  
With ioFACTURA
    .LongMaxima = 10
    Set .DataSource = adoPrimaryRS
    .DataField = "FACTURA"
End With
  
With ioSUCODIGO
    .LongMaxima = 10
    Set .DataSource = adoPrimaryRS
    .DataField = "SUCODIGO"
    .dspFormat = "dd/mm/yyyy"
End With
  
With ioALBARAN
    .LongMaxima = 10
    Set .DataSource = adoPrimaryRS
    .DataField = "ALBARAN"
End With
  
With ioTRNSPORTI
    .LongMaxima = 20
    Set .DataSource = adoPrimaryRS
    .DataField = "TRNSPORTI"
End With
  
With ioDCTOPP
    .Alineacion = 1
    '.dspFormat = "00.00"
    .LongMaxima = 5
    .SoloNumeros = True
End With
    
With ioGASTOS
    .Alineacion = 1
    .dspFormat = "Currency"
    .SoloNumeros = True
End With
   
With ioPORTES
    .Alineacion = 1
    .dspFormat = "Currency"
    .SoloNumeros = True
End With
 
   
With cbCODPROV
    .ConexionString = locCnnSP
    .LenCodigo = 5
    .SQLString = "SELECT CODIGO, NOMBRE FROM MAPROV WHERE MBAJA = 0 ORDER BY CODIGO"
    .DataField = "CODPROV"
    .carga
    Set .DataSource = adoPrimaryRS
    .CodigoWidth = 800
End With

With cBFPAGO
    .ConexionString = locCnnSP
    .LenCodigo = 3
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM FPAGO WHERE MBAJA = 0 ORDER BY CODIGO"
    .CodigoWidth = 500
    .DataField = "FPAGO"
    .carga
    Set .DataSource = adoPrimaryRS
End With

With cbCODALM
    .ConexionString = locCnnSP
    .LenCodigo = 3
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM ALMACENES WHERE MBAJA = 0 ORDER BY CODIGO"
    .CodigoWidth = 500
    .DataField = "CODALM"
    .carga
    Set .DataSource = adoPrimaryRS
End With

With cbPLAZOE
    .ConexionString = locCnnSP
    .LenCodigo = 3
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM PLAZOE WHERE MBAJA = 0 ORDER BY CODIGO"
    .CodigoWidth = 500
    .DataField = "PLAZOE"
    .carga
    Set .DataSource = adoPrimaryRS
    
End With
  

      
    mbDataChanged = False
  
  Casignados = True
  
  If Not adoPrimaryRS.EOF Then
  If Not IsNull(adoPrimaryRS.fields("NUMERO").Value) Then
  'If Not IsNull(adoPrimaryRS("ChildCMD").UnderlyingValue) Then
    Call Asigna_Grid
  End If
  End If
 
End Sub



'---------------------------------------------------------------------------------------
' Subrutina   : refresca_grid_externo
' Fecha/Hora  : 17/11/2003 21:51
' Autor       : JCASTILLO
' Propósito   : Para refrescar el grid desde fuera de este formulario
'               (frmDPedPro por ejemplo).
'               Ir_a_Registro = TRUE  -> vuelve al registro actual en el
'               momento de llamar a la rutina.
'               Ir_a_Registro = FALSE -> ir al principio
'---------------------------------------------------------------------------------------
Public Sub refresca_grid_externo(Ir_a_Registro As Boolean)
Dim tmpnump As Long

  On Error GoTo refresca_grid_externo_Error

    'adoPrimaryRS.Requery
    'adoDetalleRS.Requery
    
If creando Then Exit Sub
If adoPrimaryRS.State = 0 Then Exit Sub
If (adoPrimaryRS.EOF And adoPrimaryRS.BOF) Then Exit Sub
If IsNull(adoPrimaryRS.fields("NUMERO").Value) Then Exit Sub
If adoPrimaryRS.fields("NUMERO").Value = 0 Then Exit Sub

    
    If Ir_a_Registro Then tmpnump = adoPrimaryRS.fields("NUMERO")
    
   ' With locCnnSP
   '     If .State <> 0 Then
   '      .Close
   '      .Open strLocCnnSP
   '   End If
   ' End With
    
    With adoPrimaryRS
        
        DoEvents
       ' If .State <> 0 Then
       '    DoEvents
       '    .Close
       ' End If
        
        .Requery
        '.Open cabSQL, locCnnSP, adOpenDynamic, adLockBatchOptimistic
        '.Open sqlSHAPE, locCnnSP, adOpenDynamic, adLockBatchOptimistic
        
        DoEvents
        
        If Ir_a_Registro Then 'volver al registro
            If tmpnump <> 0 Then
                If Not .BOF Then .MoveFirst
                .Find "NUMERO = " & tmpnump
            End If
        Else 'no volver al registro, ir al principio
                If Not .BOF Then .MoveFirst
        End If
        
    End With

    DoEvents
    
    'si no hay registros, limpiar grids y labels de totales y salir
    If adoPrimaryRS.RecordCount = 0 Then
        
        fg.Clear
        vsTotales.Clear
        
        lblSubTotal.Caption = ""
        lblIVA.Caption = ""
        lblPortes.Caption = ""
        lblGastos.Caption = ""
        lblTotal.Caption = ""
        
       ' Call des_enlaza_campos
       ' Call Asigna_Campos
                
                
        Exit Sub
        
    End If
        
If adoDetalleRS.State = 1 Then adoDetalleRS.Close
adoDetalleRS.Open "SELECT * FROM DETPEDPRO WHERE NUMERO = " & adoPrimaryRS.fields("NUMERO").Value & " AND ALMORIG = " & AlmacenActual, locCnnSP, adOpenDynamic, adLockOptimistic
'Set adoDetalleRS = adoPrimaryRS("ChildCMD").UnderlyingValue

'Call des_enlaza_campos
'Call Asigna_Campos
Call Re_Asigna_Grid
    
   On Error GoTo 0
   Exit Sub

refresca_grid_externo_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento refresca_grid_externo de Formulario frmPedProv"
End Sub


Private Sub des_enlaza_campos()
 
With ioNUMERO
    Set .DataSource = Nothing
    .DataField = ""
    '.Text = ""
End With
  
With ioFECHA
    Set .DataSource = Nothing
    .DataField = ""
    .Text = ""
End With
  
With ioFACTURA
    Set .DataSource = Nothing
    .DataField = ""
    .Text = ""
End With
  
With ioSUCODIGO
    Set .DataSource = Nothing
    .DataField = ""
    .Text = ""
End With
  
With ioALBARAN
    Set .DataSource = Nothing
    .DataField = ""
    .Text = ""
End With
  
With ioTRNSPORTI
    Set .DataSource = Nothing
    .DataField = ""
    .Text = ""
End With
    
With cbCODPROV
    .DataField = ""
    Set .DataSource = Nothing
    .Text = ""
End With

With cBFPAGO
    .DataField = ""
    Set .DataSource = Nothing
    .Text = ""
End With

With cbCODALM
    .DataField = ""
    Set .DataSource = Nothing
    .Text = ""
End With

With cbPLAZOE
    .DataField = ""
    Set .DataSource = Nothing
    .Text = ""
End With
  
With cbESTADO
   .DataField = ""
   Set .DataSource = Nothing
   .Text = ""
End With

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Asigna_Grid
' DateTime  : 22/10/2003 21:45
' Author    : Administrador
' Purpose   : Asigna los datos a adoDetalleRs y al grid
'---------------------------------------------------------------------------------------
Private Sub Re_Asigna_Grid()
Dim tmpcodart As Integer
Dim tmpuds As Double
Dim totuds As Double
Dim tmpimpor As Double
Dim tmpsubtot As Double
Dim totsubtot As Double
Dim totimpor As Double
Dim tmpdctopp As Double
Dim tmpimpdcto As Double
Dim fila As Long
Dim t_articulo As Variant


'1 NUMERO
'2 linea
'3 MODELO
'3 CODART
'4 TEMPOR
'5 PRECOM
'6 CODTALLA
'7 CODCOL
'8 unidades
'9 DCTO1
'10 DCTO2
'11 IVA
'12 RE
'13 FMODI
'14 METIDO

   On Error GoTo Re_Asigna_Grid_Error

 fg.Clear
 fg.Rows = 1
 
 If (adoDetalleRS Is Nothing) Then
  
       'cargar el grid de los totales de la tab totales
        Call carga_grid_totales
        Exit Sub
        
        'Insertar totales en sus campos
        With adoPrimaryRS
            If .EditMode <> adEditNone Then
                .fields("TOTALIVA") = 0
                .fields("TOTALNET") = 0
            End If
        End With
  
 ElseIf (adoDetalleRS.EOF And adoDetalleRS.BOF) Or (adoDetalleRS.RecordCount = 0) Then
  
        'cargar el grid de los totales de la tab totales
        Call carga_grid_totales
        Exit Sub
        
        'Insertar totales en sus campos
        With adoPrimaryRS
            If .EditMode <> adEditNone Then
                .fields("TOTALIVA") = 0
                .fields("TOTALNET") = 0
            End If
        End With

 End If
  
 With fg
    .Rows = 1

    .AllowBigSelection = False
    .AllowSelection = True
    .SelectionMode = flexSelectionByRow
    .AllowUserResizing = flexResizeColumns
    .HighLight = flexHighlightWithFocus
    .FocusRect = flexFocusSolid
 End With
 
 With adoDetalleRS
 
    fg.Cols = .fields.Count + 2
    fg.ColAlignment(2) = flexAlignLeftCenter

    If Not .BOF Then .MoveFirst
    fg.Redraw = flexRDNone
    
    'If tmpcodart = 0 Then 'asignar el codigo si es la primera vez
    tmpcodart = .fields("CODART")
   ' End If
   
    fila = .AbsolutePosition
     
    Do Until .EOF
    
        
    fg.Rows = fg.Rows + 1
    
   ' If Not IsNull(.Fields("NUMERO")) Then ' fg.TextMatrix(.AbsolutePosition, 1) = .Fields("NUMERO")
    
   
    '.TextMatrix(0, 1) = "Línea"
    '.TextMatrix(0, 2) = "Ref"
    '.TextMatrix(0, 3) = "Artículo"
    '.TextMatrix(0, 4) = "Temporada"
    '.TextMatrix(0, 5) = "Talla"
    '.TextMatrix(0, 6) = "Color"
    '.TextMatrix(0, 7) = "Precio Com"
    '.TextMatrix(0, 8) = "Uds."
    '.TextMatrix(0, 9) = "Dcto."
    '.TextMatrix(0, 10) = "Sub.tot. Art."
    '.TextMatrix(0, 11) = "IVA"
    '.TextMatrix(0, 12) = "RE"
    '.TextMatrix(0, 13) = "Total. Art."
    '.TextMatrix(0, 14) = "Ul.Modif."
   
    If tmpcodart = .fields("CODART") Then  'romper por codigo
        
        If Not IsNull(.fields("LINEA")) Then fg.TextMatrix(fila, 1) = .fields("LINEA")
                
        t_articulo = devuelve_matriz("SELECT MODELO, REF FROM MAARTIC WHERE CODIGO = " & .fields("CODART") & " AND TEMPOR = " & .fields("TEMPOR"), locCnnSP)
        
        If Not IsNull(t_articulo(1)) Then fg.TextMatrix(fila, 2) = Trim(t_articulo(1))
                
        If Not IsNull(.fields("CODART")) And Not IsNull(.fields("TEMPOR")) Then
            tmpcodart = .fields("CODART") 'asignar nuevo codigo
            fg.TextMatrix(fila, 3) = Format(.fields("CODART"), "00000") & " " & Trim(t_articulo(0))
            fg.TextMatrix(fila, 4) = .fields("TEMPOR")
        End If
        
        If Not IsNull(.fields("CODTALLA")) And .fields("CODTALLA") <> 0 Then fg.TextMatrix(fila, 5) = .fields("CODTALLA")
        If Not IsNull(.fields("CODCOL")) And .fields("CODCOL") <> 0 Then
            fg.TextMatrix(fila, 6) = .fields("CODCOL")
            fg.Col = 6
            fg.Row = fila
            fg.CellBackColor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & .fields("CODCOL"), locCnnSP)
        End If
                
        If Not IsNull(.fields("PRECOM")) Then fg.TextMatrix(fila, 7) = .fields("PRECOM")
        
        If Not IsNull(.fields("UNIDADES")) Then fg.TextMatrix(fila, 8) = .fields("unidades")
    
        If Not IsNull(.fields("DCTO")) Then fg.TextMatrix(fila, 9) = .fields("DCTO")
                
        'SubTotal 1
        If Not IsNull(.fields("PRECOM")) And Not IsNull(.fields("UNIDADES")) And Not IsNull(.fields("DCTO")) Then
               
               'sacar dctopp
               'If ioDCTOPP.Text <> "" Then tmpdctopp = Replace(ioDCTOPP.Valor, ".", ",")
               tmpimpdcto = ((.fields("PRECOM") * .fields("UNIDADES")) * .fields("DCTO") / 100)
               fg.TextMatrix(fila, 10) = (.fields("PRECOM") * .fields("UNIDADES")) - tmpimpdcto     ''- '_
             '  ((.fields("PRECOM") * .fields("UNIDADES") - tmpimpdcto) * tmpdctopp / 100)
                              
               'sacar dctopp
               'If ioDCTOPP.Text <> "" Then tmpdctopp = Replace(ioDCTOPP.Valor, ".", ",")
               'tmpimpdcto = ((.fields("PRECOM") * .fields("UNIDADES")) * .fields("DCTO") / 100)
               'fg.TextMatrix(fila, 10) = '_
               '(.fields("PRECOM") * .fields("UNIDADES")) - tmpimpdcto -'_
               '((.fields("PRECOM") * .fields("UNIDADES") - tmpimpdcto) * tmpdctopp / 100)
               
        End If
                  
        If Not IsNull(.fields("IVA")) Then fg.TextMatrix(fila, 11) = .fields("IVA")
    
        If Not IsNull(.fields("RE")) Then fg.TextMatrix(fila, 12) = .fields("RE")
            
        'Total Articulo
        If Not IsNull(.fields("IVA")) Then  'sumarle el importe correspondiente al IVA
               fg.TextMatrix(fila, 13) = fg.TextMatrix(fila, 10) + (fg.TextMatrix(fila, 10) * .fields("IVA") / 100)
               fg.Col = 13
               fg.Row = fila
               fg.CellFontBold = True
               
        'de lo contrario poner sin iva
        Else
               fg.TextMatrix(fila, 13) = fg.TextMatrix(fila, 10)
               fg.Col = 13
               fg.Row = fila
               fg.CellFontBold = True
                       
        End If
        
        
        
        If Not IsNull(.fields("FMODI")) Then fg.TextMatrix(fila, 14) = .fields("FMODI")
        If Not IsNull(.fields("METIDO")) Then fg.TextMatrix(fila, 15) = .fields("METIDO")
               
        tmpuds = tmpuds + .fields("UNIDADES")  'sumar unidades
        tmpimpor = tmpimpor + fg.TextMatrix(fila, 13)
        tmpsubtot = tmpsubtot + fg.TextMatrix(fila, 10)
    
       .MoveNext
        fila = fila + 1

        
    'romper por articulo
    ElseIf (tmpcodart <> .fields("CODART") Or .EOF) Then
    
    'insertar fila de subtotal
        Call imprime_Subtotal(tmpuds, tmpimpor, tmpsubtot, fila, "Total Artículo:", False)
    
        tmpcodart = .fields("CODART")
        fila = fila + 1
    
        totuds = totuds + tmpuds
        totimpor = totimpor + tmpimpor
        totsubtot = totsubtot + tmpsubtot
        tmpsubtot = 0
    
        tmpuds = 0
        tmpimpor = 0
        
    End If
    Loop
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    fg.Rows = fg.Rows + 1
    totuds = totuds + tmpuds
    totimpor = totimpor + tmpimpor
    totsubtot = totsubtot + tmpsubtot
    'insertar fila de subtotal
    
    Call imprime_Subtotal(tmpuds, tmpimpor, tmpsubtot, fila, "Total Artículo:", False)
    
    'añadir una linea al principio para el total general
    fg.AddItem "", 1
    'insertar el total general
    Call imprime_Subtotal(totuds, totimpor, totsubtot, 1, "Total General:", True)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
        
 End With
 
 
 With fg
    
   ' .Subtotal flexSTClear
    .TextMatrix(0, 1) = "Línea"
    .TextMatrix(0, 2) = "Ref"
    .TextMatrix(0, 3) = "Artículo"
    .TextMatrix(0, 4) = "Temporada"
    .TextMatrix(0, 5) = "Talla"
    .TextMatrix(0, 6) = "Color"
    .TextMatrix(0, 7) = "Precio Com"
    .TextMatrix(0, 8) = "Uds."
    .TextMatrix(0, 9) = "Dcto."
    .TextMatrix(0, 10) = "Sub.tot. Art."
    .TextMatrix(0, 11) = "IVA"
    .TextMatrix(0, 12) = "RE"
    .TextMatrix(0, 13) = "Total. Art."
    .TextMatrix(0, 14) = "Ul.Modif."
    .TextMatrix(0, 15) = "En Almacén"
    
    '.Subtotal flexSTSum, 0, 7, , vbBlue, vbWhite, True, "Total Uds. General:"
    '.SubtotalPosition = flexSTAbove
    '.Subtotal flexSTSum, 2, 7, , 16744576, , True, "Total por Articulo:"


    .ColFormat(1) = "000"
    
    .ColComboList(4) = tmptempor
    .ColComboList(5) = tmptalla
    .ColComboList(6) = tmpcolor
    
    .ColFormat(7) = "Currency"
    .ColFormat(10) = "Currency"
    .ColFormat(13) = "Currency"
    
  '  .ColDataType(14) = flexDTBoolean
    '.ColFormat(14) = "No;Si"

    .AutoSize 1, .Cols - 1
    
  End With
  
  With adoDetalleRS
  
    If Not .BOF Then .MoveFirst
  
  End With

  fg.Redraw = True
  
  'cargar el grid de los totales de la tab totales
  Call carga_grid_totales

   On Error GoTo 0
   Exit Sub

Re_Asigna_Grid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Re_Asigna_Grid de Formulario frmPedProv"
  
End Sub


'---------------------------------------------------------------------------------------
' Subrutina   : imprime_Subtotal
' Fecha/Hora  : 16/11/2003 21:08
' Autor       : JCASTILLO
' Propósito   : Imprimir la linea de subtotales en el flexgrid
'
'---------------------------------------------------------------------------------------
Private Sub imprime_Subtotal(unidades As Double, Importe As Double, subtotal As Double, fila As Long, texto As String, invertido As Boolean)
Dim tmpcol As Byte

   On Error GoTo imprime_Subtotal_Error
    
    fg.Row = fila
    
        For tmpcol = 1 To fg.Cols - 1
            fg.Col = tmpcol
            
            If Not invertido Then
            
                fg.CellBackColor = vbWhite
                fg.CellForeColor = vbBlue
            
            Else
            
                fg.CellBackColor = vbBlue
                fg.CellForeColor = vbWhite
            
            End If
            
            fg.CellFontBold = True
        Next
    
    fg.TextMatrix(fila, 2) = texto
    fg.TextMatrix(fila, 3) = fg.TextMatrix(fila - 1, 3) 'poner la descripción del articulo

    fg.TextMatrix(fila, 8) = unidades
    fg.TextMatrix(fila, 10) = subtotal
    fg.TextMatrix(fila, 13) = Importe
    
           

   On Error GoTo 0
   Exit Sub

imprime_Subtotal_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento imprime_Subtotal de Formulario TmpMaDet"
End Sub

Private Sub Asigna_Grid()
'On Error Resume Next
Dim tmprc As New ADODB.Recordset

   On Error GoTo Asigna_Grid_Error

  If Not adoPrimaryRS.EOF Then
  
 If Not IsNull(adoPrimaryRS.fields("NUMERO")) Then
        If adoDetalleRS.State = 1 Then adoDetalleRS.Close
        adoDetalleRS.Open "SELECT * FROM DETPEDPRO WHERE NUMERO = " & adoPrimaryRS.fields("NUMERO").Value & " AND ALMORIG = " & AlmacenActual, locCnnSP, adOpenDynamic, adLockOptimistic
 End If

  'If Not IsNull(adoPrimaryRS("ChildCMD").UnderlyingValue) Then Set adoDetalleRS = adoPrimaryRS("ChildCMD").UnderlyingValue
  
'01. numberformat
'02. numberformat
'03. numberformat
'04. temporada
'06. talla
'07. color
'09. formato %
'10. formato %
'11. formato %
'12. formato %


    With tmprc
        .Open "SELECT IDTEM, ABREVIA FROM TEMPOR WHERE MBAJA = 0 ORDER BY IDTEM", locCnnSP, adOpenDynamic, adLockReadOnly
        tmptempor = fg.BuildComboList(tmprc, "ABREVIA", "IDTEM", vbBlue)
        .Close
        .Open "SELECT CODIGO, DESCRIPCION FROM TALLAS WHERE MBAJA = 0 ORDER BY CODIGO", locCnnSP, adOpenDynamic, adLockReadOnly
        tmptalla = fg.BuildComboList(tmprc, "DESCRIPCION", "CODIGO", vbBlue)
        .Close
        .Open "SELECT CODIGO, DESCRIPCION FROM COLORES WHERE MBAJA = 0 ORDER BY CODIGO", locCnnSP, adOpenDynamic, adLockReadOnly
        tmpcolor = fg.BuildComboList(tmprc, "DESCRIPCION", "CODIGO", vbBlue)
        .Close
  '      .Open "SELECT CODIGO, MODELO FROM MAARTIC WHERE MBAJA = 0 AND TEMPOR = " & TemporadaActual & " ORDER BY CODIGO", locCnn, adOpenDynamic, adLockReadOnly
'        tmpcodigo = fg.BuildComboList(tmprc, "MODELO", "CODIGO", vbBlue)
      '  .Close
    End With
    
    
    
     'Set fg.DataSource = adoPrimaryRS("ChildCMD").UnderlyingValue
    'fg.DataMode = flexDMBoundBatch
    
    Call Re_Asigna_Grid
  
  'formato para el grid
 
  Gasignados = True
  
  End If


Set tmprc = Nothing

   On Error GoTo 0
   Exit Sub

Asigna_Grid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Asigna_Grid de Formulario frmPedProv"
End Sub







Private Sub btA_click()
'Dim tmpcodigo As Long

   On Error GoTo btA_click_Error

    If adoPrimaryRS.RecordCount < 0 Then Exit Sub
    
    'tmpcodigo = adoPrimaryRS.Fields("CODIGO")

    btB.Value = False
    
    If adoPrimaryRS.fields("DESTINO") <> 0 Then
        locCnnSP.Execute "UPDATE CABPEDPRO SET DESTINO = 0 WHERE NUMERO = " & adoPrimaryRS.fields("NUMERO") & " AND ALMORIG = " & adoPrimaryRS.fields("ALMORIG")
    End If
    
    DoEvents

    Call refresca_grid_externo(True)
    
    'adoPrimaryRS.Find "CODIGO = " & tmpcodigo, , adSearchForward, 1
    
    lblStatusD.Caption = "PEDIDO EN A"
    

   On Error GoTo 0
   Exit Sub

btA_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento btA_click de Formulario frmPedProv"
    
End Sub

Private Sub btB_click()
'Dim tmpcodigo As Long

   On Error GoTo btB_click_Error

   If adoPrimaryRS.RecordCount < 0 Then Exit Sub
   
   'tmpcodigo = adoPrimaryRS.Fields("CODIGO")

   btA.Value = False
   If adoPrimaryRS.fields("DESTINO") <> 1 Then
    locCnnSP.Execute "UPDATE CABPEDPRO SET DESTINO = 1 WHERE NUMERO = " & adoPrimaryRS.fields("NUMERO") & " AND ALMORIG = " & adoPrimaryRS.fields("ALMORIG")
   End If
   DoEvents
   
   Call refresca_grid_externo(True)
    
   'adoPrimaryRS.Find "CODIGO = " & tmpcodigo, , adSearchForward, 1
    
   'adoPrimaryRS.Requery
   lblStatusD.Caption = "PEDIDO EN B"

   On Error GoTo 0
   Exit Sub

btB_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento btB_click de Formulario frmPedProv"

End Sub





Private Sub cbPLAZOE_GotFocus()
    If Tab1.Tab = 1 Then
        Tab1.Tab = 0
        DoEvents
    End If
End Sub

Private Sub cbPLAZOE_Validate(Cancel As Boolean)
    Tab1.Tab = 1
End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : cmCambTem_Click
' Fecha/Hora  : 03/12/2003 21:46
' Autor       : JCASTILLO
' Propósito   : Cambio de la temporada de trabajo actual (variable TemporadaWrk).
'---------------------------------------------------------------------------------------
Private Sub cmCambTem_Click()
Dim tmptext As String

   On Error GoTo cmCambTem_Click_Error

If Temporada.Text <> "" Then
    tmptext = devuelve_campo("SELECT AÑO + '  ' + TEMPORADA FROM TEMPOR WHERE IDTEM = " & Temporada.Text, locCnnSP)
    
    If NoAvisar = False Then
        If MsgBox("¿Desea cambiar la temporada actual de TRABAJO? (solo para esta sesion), y poner la siguiente: " & Chr(13) & tmptext, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    TemporadaWrk = CByte(Temporada.Text)
    lblTemporadaWrk.Caption = tmptext
    lblstatus.Caption = ""
Else
    lblstatus.Caption = "Temporada no puede estar en blanco"
End If

tmptext = ""

   On Error GoTo 0
   Exit Sub

cmCambTem_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmCambTem_Click de Formulario frmPedProv"

End Sub

Private Sub cmCamTem_Click()

End Sub

Private Sub cmComentario_Click()

FrmInicio.Editor.carga "Comentario de Pedidos [" & ioNUMERO.Caption & "]", adoPrimaryRS.fields("COMENTARIO"), ""

End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : cbListaProv_Click
' Fecha/Hora  : 29/12/2003 00:38
' Autor       : JCASTILLO
' Propósito   : Mostrar el grid de proveedores para seleccionar un proveedor para este
'               pedido
'---------------------------------------------------------------------------------------
Private Sub cbListaProv_Click()
Dim tmpsql As New clsSmartSQL

Dim tmprc As New ADODB.Recordset

   On Error GoTo cbListaProv_Click_Error
    
    If Not mbAddNewFlag Then Exit Sub

    tmpsql.AddTable "MAPROV"
    tmpsql.AddSimpleWhereClause "MBAJA", 0
    tmpsql.AddOrderClause "CODIGO"
    
With frmFlexProv
    
    tmprc.Open tmpsql.SQL, locCnnSP, adOpenDynamic, adLockReadOnly
    
    Set .miosql = tmpsql
    
    With .fg
             Set frmFlexProv.miRc = tmprc
    End With
    
    .Caption = "Proveedores ..."
    
    .Show 1
        
    cbCODPROV.Text = tmprc.fields("CODIGO")
    cBFPAGO.SetFocus
    
    tmprc.Close
    Set tmprc = Nothing
    Set tmpsql = Nothing
    
End With

   On Error GoTo 0
   Exit Sub

cbListaProv_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbListaProv_Click de Formulario frmPedProv"

End Sub

Private Sub cmDeshacerLinea_Click()

End Sub

Private Sub cmEditarPorRef_Click()
Dim tmpref As String
Dim fila As Single

   On Error GoTo cmEditarPorRef_Click_Error

 tmpref = InputBox("Introduzca referencia:", "Buscar por referencia")
 
 If Trim(tmpref) = "" Then Exit Sub
 
 fila = fg.FindRow(tmpref, , 2)
 
 If fila < 0 Then
    lblstatus.Caption = "No se ha encontrado la referencia: " & tmpref
    Exit Sub
 End If
 
 adoDetalleRS.Move fila - 1, 0
 fg.Row = fila
 
 DoEvents

 Call fg_dblClick

   On Error GoTo 0
   Exit Sub

cmEditarPorRef_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmEditarPorRef_Click de Formulario frmPedProv"

End Sub

Private Sub fg_KeyDown(KeyCode As Integer, Shift As Integer)

'ir al artículo con intro
If KeyCode = 13 Then

 If fg.Rows > 1 Then
    Call fg_dblClick
    
 End If

KeyCode = 0
End If

End Sub



Private Sub ionumero_Change()
ioNUMERO.Caption = Format(ioNUMERO.Caption, "00000000")
End Sub

Private Sub cbSeleccionaArticulo_Click()
Dim artsql As New clsSmartSQL
Dim mrc As New ADODB.Recordset

If adoPrimaryRS.RecordCount = 0 Then Exit Sub
   
If Temporada.Text = "" Then
    lblstatus.Caption = "Seleccione TEMPORADA"
    Temporada.SetFocus
    Exit Sub
End If


With artsql
    .AddTable "MAARTIC"
    .AddOrderClause "CODIGO"
    .AddSimpleWhereClause "MBAJA", "0"
    .AddSimpleWhereClause "HIST", "0"
End With

    mrc.Open artsql.SQL, locCnnSP, adOpenDynamic, adLockReadOnly

    With frmFlexArt
    
    Set .miosql = artsql
    Set .miRc = mrc
    'el recordset de este formulario
    
    Set .rc_detalle = adoDetalleRS
    .NumeroPedido = adoPrimaryRS.fields("NUMERO")
    .add_en_detalle = True
    .tmpcodprov = cbCODPROV.Text
    .Temporada_Defecto = TemporadaWrk
    
    .Show 1
    DoEvents
    
    
    
    End With
    
    mrc.Close
    Set mrc = Nothing
    
   'posicionarse y editar la ultima linea insertada
   If Not adoPrimaryRS.RecordCount = 0 Then
     If Not adoDetalleRS.BOF And Linea_Creada > 0 Then
        
        'abrir para editar la ultima linea añadida
        adoDetalleRS.MoveFirst
        adoDetalleRS.Find "LINEA = " & Linea_Creada
               
   With frmDPedPro
      .NumeroPedido = ioNUMERO.Caption
      If Not adoPrimaryRS.EOF Then
      Set .rc = adoDetalleRS
      End If
    .Show
   End With
   
       Linea_Creada = 0
     
    End If
   End If

Set artsql = Nothing

End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : cmDesacerLinea_Click
' Fecha/Hora    : 19/11/2003 16:27
' Autor         : JCastillo
' Propósito     :  Saca el artículo de STOCK (descontando las unidades previamente
'                      introducidas, y lo marca como METIDO = 0. Esto permite volver a
'                      editar el artículo.
'---------------------------------------------------------------------------------------
Private Sub cmDesacerLinea_Click()
Dim miConn As New ADODB.Connection
Dim rccnn As New ADODB.Recordset
Dim tmptransf As Double
Dim nuetrans As Long
Dim posicion_Reg As Long
Dim posicion_fg As Long

   
   On Error GoTo cmDesacerLinea_Click_Error

'   se_ha_desecho = False

   If adoPrimaryRS.RecordCount = 0 Then
    lblstatus.Caption = "No se encuentran artículos para Deshacer."
    Exit Sub
   End If
   

   If fg.TextMatrix(fg.Row, 1) = "" Or Not IsNumeric(fg.TextMatrix(fg.Row, 1)) Then
    lblstatus.Caption = ""
   Exit Sub
   End If
   
   DoEvents
 
    If MsgBox("¿Desea deshacer el artículo seleccionado? " & Chr(13) & "[" & fg.TextMatrix(fg.Row, 2) & "].  Uds:[" & fg.TextMatrix(fg.Row, 7) & "] Talla:[" & devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & fg.TextMatrix(fg.Row, 5), locCnnSP) & "] Color:[" & devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & fg.TextMatrix(fg.Row, 6), locCnnSP) & "]" & Chr(13) & _
    "Nota: Se descontarán las unidades y quedará el stock como si no se hubiera introducido.", vbYesNo + vbQuestion) = vbNo Then Exit Sub
   
    
    tmptransf = adoPrimaryRS.fields("CODPTRN").Value
   
    'ir al registro especificado
    With adoDetalleRS
        .MoveFirst
        .Find "LINEA = " & fg.TextMatrix(fg.Row, 1), , adSearchForward
        
        If .fields("METIDO") = 0 Then
            lblstatus.Caption = "El artículo no se encuentra en almacén, imposible deshacer"
            Exit Sub
        End If
        
    End With
    
    posicion_Reg = adoDetalleRS.AbsolutePosition
    'posicion_fg = fg.TextMatrix(fg.Row, 1)
    posicion_fg = fg.Row
    
     
     With locCnnSP
                Set adoPrimaryRS.ActiveConnection = Nothing
                Set adoDetalleRS.ActiveConnection = Nothing
                
        If .State <> 0 Then
                 .Close

                .CursorLocation = adUseServer
                .Open strLocCnn
                DoEvents
                .BeginTrans
        End If
    End With

    'realizar los cambios
        'miConn.CursorLocation = adUseServer
        'miConn.Open strLocCnnSP
        'miConn.BeginTrans
        
                '====================================================
                'SI ES PARA EL ALMACEN ACTUAL, HACER LOS CAMBIOS EN STOCK DIRECTAMENTE
                '====================================================
                If cbCODALM.Text = AlmacenActual Then
                  
                    Call stock(adoDetalleRS.fields("CODART"), adoDetalleRS.fields("TEMPOR"), adoDetalleRS.fields("CODTALLA"), adoDetalleRS.fields("CODCOL"), AlmacenActual, adoDetalleRS.fields("UNIDADES"), False, locCnnSP)
                
                Else
                '====================================================
                'SI ES PARA OTRO ALMACEN, IR A LA TRANSFERENCIA CORRESPONDIENTE
                'Y BORRAR LA LINEA. CODPTRN
                '====================================================
                locCnnSP.Execute "DELETE FROM DETTRANS WHERE (CODIGO = " & tmptransf & ") AND (CODART = " & adoDetalleRS.fields("CODART") & ")" & _
                      " AND (CODCOL = " & adoDetalleRS.fields("CODCOL") & ") AND (CODTALLA = " & adoDetalleRS.fields("CODTALLA") & ") AND (TEMPOR = " & adoDetalleRS.fields("TEMPOR") & ")" & _
                      " AND (UNIDADES = " & adoDetalleRS.fields("UNIDADES") & ")"
                                                   
                End If
                
                'ADEMAS DESCONTAR LAS UNIDADES DEL CENTRAL
                'Call stock(adoDetalleRS.Fields("CODART"), adoDetalleRS.Fields("TEMPOR"), adoDetalleRS.Fields("CODTALLA"), adoDetalleRS.Fields("CODCOL"), AlmacenActual, adoDetalleRS.Fields("UNIDADES"), False, locCnnSP)
   
        With locCnnSP
            .CommitTrans
            .Close
            .CursorLocation = adUseClient
            .Open strLocCnnSP
        End With
        
        Set adoPrimaryRS.ActiveConnection = locCnnSP
        Set adoDetalleRS.ActiveConnection = locCnnSP
         
        adoDetalleRS.fields("METIDO") = 0
        adoDetalleRS.UpdateBatch adAffectAll
        
        'Call refresca_grid_externo(False)
   

        'miConn.CommitTrans
        'miConn.Close
        'Set miConn = Nothing
                       
        DoEvents
        
        Call refresca_grid_externo(True)
        
        DoEvents
        
        
        adoDetalleRS.Move posicion_Reg, 0
        fg.Row = posicion_fg
        fg.Select posicion_fg, 1, posicion_fg, fg.Cols - 1
        fg.RowSel = posicion_fg
        fg.ColSel = 1
        'posicion_fg = fg.FindRow(posicion_fg, , 1)
        
  ' se_ha_desecho = True
   
   On Error GoTo 0
      
   Exit Sub

cmDesacerLinea_Click_Error:

   With locCnnSP
   If .CursorLocation = adUseServer Then
        .RollbackTrans
        .Close
        .CursorLocation = adUseClient
        .Open strLocCnnSP
        Call refresca_grid_externo(True)
   End If
   End With

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmDesacerLinea_Click de Formulario frmPedProv"
End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : cmEntrada_Click
' Fecha/Hora    : 19/11/2003 13:31
' Autor         : JCastillo
' Propósito     : Dar entrada de mercancia. Añadir las unidades correspondientes
'                     en la tabla stock, marcar como METIDO = 1 en detpedpro y
'                     imprimir las etiquetas
'---------------------------------------------------------------------------------------
Private Sub cmEntrada_Click()
Dim rcSrv As New ADODB.Recordset
Dim tmpnump As Double
Dim tmptransf As Double
Dim tansfped As Double
Dim tmpprecom As Double
Dim entrans As Boolean

'si hay alguna combinacion talla o color en blanco salir
On Error GoTo cmEntrada_Click_Error

If comprueba_tallas_colores Then Exit Sub

If cmEntrada.Enabled = False Then Exit Sub
cmEntrada.Enabled = False

lblstatus.Caption = "Registrando pedido en el sistema ..."

tmpnump = adoPrimaryRS.fields("NUMERO").Value
tansfped = adoPrimaryRS.fields("CODPTRN").Value



'abrir un recordset desconectado
'rcSrv.Open "SELECT * FROM DETPEDPRO WHERE (NUMERO = " & tmpnump & " and ALMORIG = " & AlmacenActual & ") ORDER BY LINEA ", locCnnSP, adOpenDynamic, adLockBatchOptimistic
  
'With locCnnSP
'    If .State <> 0 Then
'       .Close
       
'       Set rcSrv.ActiveConnection = Nothing
       
'       .CursorLocation = adUseServer
'       .Open strLocCnn
'       DoEvents
'       .BeginTrans
'       entrans = True
'    End If
'End With

'With rcSrv

 ' .Open "SELECT NUMERO, LINEA, CODART, TEMPOR, PRECOM, CODTALLA, CODCOL, UNIDADES, METIDO, PRECOM, DCTO, IVA, RE FROM DETPEDPRO WHERE NUMERO = " & tmpnump & " AND ALMORIG = " & AlmacenActual & "  ORDER BY LINEA", locCnnSP, adOpenDynamic, adLockOptimistic
 ' DoEvents
    
   'dar una primera pasada para poner los PRECOM en MAARTIC
    'Do Until .EOF
    
       'quitarle el dcto
    '  tmpprecom = .fields("PRECOM") - ((.fields("PRECOM") * .fields("DCTO")) / 100)
       'sumarle el IVA
    '  tmpprecom = tmpprecom + ((tmpprecom * .fields("IVA")) / 100)
                     
      'actualizar el precio de compra y IVACOM en el maestro ...
    '  locCnnSP.Execute "UPDATE MAARTIC SET PRECOM = " & Replace(tmpprecom, ",", ".") & ", IVACOM = " & CLng(.fields("IVA")) & " WHERE CODIGO = " & .fields("CODART") & " AND TEMPOR = " & .fields("TEMPOR")
      
    '  DoEvents
       
    '  .MoveNext
      
    'Loop
       

    'If Not .BOF Then .MoveFirst

  'DoEvents
  
  frmImpEtiq.Show 1
  
  Call CreateDBEtiquetas(adoDetalleRS, frmImpEtiq.digitos, frmImpEtiq.saltar)
  
  DoEvents
    
 'dar entrada a los artículos
  'If .EOF Then Exit Sub
  'If Not .BOF Then .MoveFirst
  
    
    'si es para otro almacén
   ' If cbCODALM.Text <> AlmacenActual Then
    
        'si no existe ya una transferencia para este pedido
   '     If tansfped = 0 Then
            'creamos un registro de cabecera de transferencias
    '        tmptransf = crear_transferencia(AlmacenActual, CLng(cbCODALM.Text), locCnnSP, True, tmpnump)
        
            'asociar esta transferencia con su pedido
            'locCnnSP.Execute "UPDATE CABPEDPRO SET CODPTRN = " & tmptransf & " WHERE NUMERO = " & tmpnump & " AND ALMORIG = " & AlmacenActual
        
        'si existe, asignar la transferencia existente
     '   Else
            
   '         tmptransf = tansfped
            
   '     End If
        
   ' End If
               
        
    'If Not .BOF Then .MoveFirst
    'Do Until .EOF
       
    '    If .fields("METIDO") = 0 Then
             
                '====================================================
                'SI ES PARA EL ALMACEN ACTUAL, HACER LOS CAMBIOS EN STOCK DIRECTAMENTE
                '====================================================
    '            If cbCODALM.Text = AlmacenActual Then
                  
     '               Call stock(.fields("CODART"), .fields("TEMPOR"), .fields("CODTALLA"), .fields("CODCOL"), AlmacenActual, .fields("UNIDADES"), True, locCnnSP)
     '               DoEvents
      '
'                Else
                '====================================================
                'SI ES PARA OTRO ALMACEN, AÑADIR UNA NUEVATRANSFERENCIA y  GUARDAR EL CODIGO
                'DE TRANSFERENCIA PARA REALIZAR LOS CAMBIOS SOBRE LA MISMA. (y asegurarse de no
                'crear varias transferencias para un mismo pedido). CODPTRN
                '====================================================
                        
                'primero pasarlo al stock del almacen inicial para poder hacer la trasnferencia al final
              '  Call stock(.Fields("CODART"), .Fields("TEMPOR"), .Fields("CODTALLA"), .Fields("CODCOL"), AlmacenActual, .Fields("UNIDADES"), True, locCnnSP)
                
                'añadimos una nueva linea a la transferencia q cremos antes
     '           Call crear_linea_transferencia(tmptransf, AlmacenActual, .fields("codart"), .fields("TEMPOR"), .fields("CODTALLA"), .fields("CODCOL"), .fields("UNIDADES"), locCnnSP, True)
                                                
      '          End If
                              
            ' .Fields("NTRANS") = crear_nueva_transferencia(.Fields("CODART"), .Fields("TEMPOR"), .Fields("CODTALLA"), .Fields("CODCOL"), 0, cbCODALM.Text, True, False, .Fields("UNIDADES"), locCnnSP)
             'Call stock(.Fields("CODART"), .Fields("TEMPOR"), .Fields("CODTALLA"), .Fields("CODCOL"), cbCODALM.Text, .Fields("UNIDADES"), True, locCnnSP)
             
        
       '       DoEvents
             
        '      .fields("METIDO") = 1
        '     .Update
             
             
             
      '  End If
        
    '    .MoveNext
               
   ' Loop
    
'End With



lblstatus.Caption = "Generando etiquetas  ..."
Call generar_etiquetas(locCnnSP)

lblstatus.Caption = "Operación finalizada correctamente"

'Call refresca_grid_externo(True)

adoPrimaryRS.Find "NUMERO =" & tmpnump

Call Re_Asigna_Grid

   If rcSrv.State <> 0 Then
    rcSrv.Close
   End If
   
Set rcSrv = Nothing

'entrans = False

cmEntrada.Enabled = True

Exit Sub
   
cmEntrada_Click_Error:
   
'   On Error Resume Next
   
   If rcSrv.State <> 0 Then
    rcSrv.Close
   End If
   
   Set rcSrv = Nothing
   
   With locCnnSP
    If entrans Then .RollbackTrans
    .Close
    .CursorLocation = adUseClient
    .Open strLocCnnSP
    Call refresca_grid_externo(True)
   End With
   
   cmEntrada.Enabled = True

   lblstatus.Caption = ""
   MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmEntrada_Click de Formulario frmPedProv"

End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : cmHistorico_Click
' Fecha/Hora  : 26/11/2003 21:25
' Autor       : JCASTILLO
' Propósito   : Pasar el pedido ACTUAL al historico (estado = 2)
'---------------------------------------------------------------------------------------
Private Sub cmHistorico_Click()
Dim tmprc As New ADODB.Recordset
Dim tmptrans As Double
Dim tmpdias As Variant

On Error GoTo cmHistorico_Click_Error


'ir al modulo de aceptar pedido
Call ExecCmdNoFocus("UWACEPTAPEDIDOS.EXE")

'////////////////// Cortar aqui el codigo anterior //////////////////////////////////
Exit Sub

If adoPrimaryRS.RecordCount = 0 Then Exit Sub

If MsgBox("¿Desea pasar el pedido actual (nº " & ioNUMERO.Caption & ") al histórico de pedidos?", vbYesNo + vbQuestion) = vbNo Then Exit Sub

tmptrans = adoPrimaryRS.fields("CODPTRN")

With tmprc
    
    .Open "SELECT count(linea) FROM DETPEDPRO WHERE NUMERO = " & adoPrimaryRS.fields("NUMERO").Value & " AND ALMORIG = " & adoPrimaryRS.fields("ALMORIG") & " AND METIDO = 0", locCnnSP, adOpenDynamic, adLockReadOnly
    
    If .fields(0).Value > 0 Then
        MsgBox "No se puede pasar el pedido al HISTORICO mientras existan artículos pendientes:" & Chr(13) & .fields(0).Value & " artículo/s. pulse RECIBIR PEDIDO para añadirlos al almacén", vbInformation
        .Close
        Set tmprc = Nothing
        Exit Sub
    
    Else
    
    'Set adoPrimaryRS.ActiveConnection = Nothing
    'Set adoDetalleRS.ActiveConnection = Nothing
    
    Set .ActiveConnection = Nothing
    
    With locCnnSP
      '  If .State = 1 Then .Close
      '  .CursorLocation = adUseServer
      '  .Open locCnnSP
      '  .BeginTrans
        'pasar al historico (estado = 3)
        .Execute "UPDATE CABPEDPRO SET ESTADO = 3 WHERE NUMERO =" & adoPrimaryRS.fields("NUMERO").Value & " AND ALMORIG = " & adoPrimaryRS.fields("ALMORIG")
        'pasar transferencia a PENDIENTE (estado = 1)
        .Execute "UPDATE PTRANS SET ESTADO = 1 WHERE CODIGO = " & tmptrans & " and CODALMORIG =" & adoPrimaryRS.fields("ALMORIG")
        
      '  .CommitTrans
       ' .Close
       ' .CursorLocation = adUseClient
       ' .Open locCnnSP
    End With
    
    'Set adoPrimaryRS.ActiveConnection = locCnnSP
    'Set adoDetalleRS.ActiveConnection = locCnnSP
    
    End If
    
    .Close
    Set tmprc = Nothing
    
End With


   '//////////////////////////////////////////////////////////////////////////////////////////////////////////
   ' Introducir el PAGO en la tabla de PAGOS
   tmpdias = devuelve_campo("SELECT DIAS FROM FPAGO WHERE CODIGO = " & adoPrimaryRS.fields("FPAGO"), locCnnSP)
   If tmpdias = "@" Then tmpdias = 0
   
   'solo si la forma de pago tiene dias ...
   If tmpdias > 0 Then
    Call introduce_pago(CajaActual, adoPrimaryRS.fields("TOTALNET") + adoPrimaryRS.fields("TOTALIVA") + adoPrimaryRS.fields("GASTOS") + adoPrimaryRS.fields("PORTES"), CInt(tmpdias), adoPrimaryRS.fields("CODPROV"), adoPrimaryRS.fields("NUMERO"), Trim(adoPrimaryRS.fields("FACTURA")), adoPrimaryRS.fields("FPAGO"), locCnnSP)
   End If
   '//////////////////////////////////////////////////////////////////////////////////////////////////////////
 
  
   'le ponemos como parametro false para q no intente volver al registro
   'actual pues lo hemos pasado al historico
   Call refresca_grid_externo(False)
  
  
  
   Call cmdLast_Click
   
     
   
   DoEvents
   locCnnSP.Execute "BACKUP LOG LOCAL WITH TRUNCATE_ONLY"

   On Error GoTo 0
   Exit Sub

cmHistorico_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmHistorico_Click de Formulario frmPedProv"
 
End Sub


Private Sub dtAgregar_Click()

If adoPrimaryRS.RecordCount = 0 Then Exit Sub

If Temporada.Text = "" Then
    lblstatus.Caption = "Seleccione TEMPORADA"
    Temporada.SetFocus
    Exit Sub
End If

''**********************************************************
' ATENCION
' Poner en el detalle de pedidos como clave 2 campos:
' NUMERO
' LINEA
' donde linea empieza a contar desde 1 en cada pedido
''**********************************************************
If Not Gasignados Then Call Asigna_Grid
 
DoEvents

With FrmMntArt
    .NumeroPedido = adoPrimaryRS.fields("NUMERO")
    .add_en_detalle = True
    .ProveedPedido = CLng(cbCODPROV.Text)
    Set .rc_detalle = adoDetalleRS
    .TemporadaWrk = TemporadaWrk
    .Show
    DoEvents
End With

End Sub


Private Sub introduce_pago(codigo_caja As Byte, Importe As Currency, dias As Integer, codigo_proveedor As Long, NUMERO_PEDIDO As Long, factura As String, forma_pago As Integer, conexion As ADODB.Connection)
Dim tmpcodigo As Variant

   On Error GoTo introduce_pago_Error

      
    If Not IsNumeric(factura) Then factura = 0
    
    'obtener el ultimo pago
    tmpcodigo = devuelve_campo("SELECT MAX(CODIGO) + 1 FROM PAGOS WHERE CODCAJA = " & codigo_caja, conexion)
    
    If tmpcodigo = "@" Then tmpcodigo = 1
    
    conexion.Execute "INSERT INTO PAGOS (CODIGO, CODCAJA, TIPOPAGO, CODPROV, IMPORTE, PAGADO, MESES, CUOTA, NUMPED, FACTURA, FPAGO, ESTADO, DESCRIPCION) " & _
                     "VALUES (" & tmpcodigo & ", " & codigo_caja & ", " & "1" & ", " & codigo_proveedor & ", " & Replace(Importe, ",", ".") & ", " & "0" & ", " & Round(dias / 30) & ", " & Replace(Importe / (dias / 30), ",", ".") & ", " & NUMERO_PEDIDO & ", " & factura & ", " & forma_pago & ", " & "0" & ", " & "'Desde Recibir Mercancía'" & ")"
    
    DoEvents
    
   On Error GoTo 0
   Exit Sub

introduce_pago_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento introduce_pago de Formulario frmPedProv"
End Sub


Private Sub dtEdicion_Click()

Call fg_dblClick

End Sub

Private Sub fg_dblClick()
'Dim tmpmetido As Variant


   If adoPrimaryRS.RecordCount = 0 Then
    lblstatus.Caption = "No se encuentran artículos para editar."
    Exit Sub
   End If
   
   
If Temporada.Text = "" Then
    lblstatus.Caption = "Seleccione TEMPORADA"
    Temporada.SetFocus
    Exit Sub
End If
   

   With fg

   If .TextMatrix(.Row, 1) = "" Or Not IsNumeric(.TextMatrix(.Row, 1)) Then
    lblstatus.Caption = ""
   Exit Sub
   End If
  
   
   lblstatus.Caption = ""
   DoEvents
           
    'ir al registro especificado
    adoDetalleRS.MoveFirst
    adoDetalleRS.Find "LINEA = " & .TextMatrix(.Row, 1), , adSearchForward
    
          'si ya se ha introducido, no dejar modificar hasta que se haga un deshacer artículo.
   If adoDetalleRS.fields("METIDO") = True Then
   
        lblstatus.Caption = "El artículo ya se ha introducido en STOCK. Debe deshacer el pedido para editar."
        Exit Sub
    
   End If
    
     '   Call cmDesacerLinea_Click
     '   DoEvents
        
        
        'si ha contestado que no deshacer (el artículo sigue estando
        'como metido, salir
        'tmpmetido = devuelve_campo("SELECT METIDO FROM DETPEDPRO WHERE LINEA = " & adoDetalleRS.fields("LINEA"), locCnnSP)
        
        'If tmpmetido <> 0 Then Exit Sub
        
        
        'If adoDetalleRS.fields("METIDO") <> 0 Then Exit Sub
        
        
   '     Call fg_dblClick

   
   
   End With
    
   If Not Gasignados Then Call Asigna_Grid
   
   With frmDPedPro
      .NumeroPedido = ioNUMERO.Caption
      If Not adoPrimaryRS.EOF Then
      Set .rc = adoDetalleRS
      End If
    .Show
   End With
   
   
   
End Sub

Private Sub Form_Load()
Dim tem As String
 
  
  Dim detSQL As String
  
  'Detalle de pedidos
  
  Move (Screen.Width - Width) \ 2, Separacion_MDIForm
  
  With oSQL
  
    .AddTable "DETPEDPRO"
    .AddOrderClause "CODART"
    .AddOrderClause "LINEA"
  
  detSQL = .SQL

  'Set oSQL = Nothing
  'DoEvents
  
  'Cabecera de pedidos
    .ClearFromClause
    .ClearOrderClause
    .ClearWhereClause
    .ClearFields
    
   .AddTable "CABPEDPRO"
   
   'si viene de trabajar con pedido, mostrar solo el pedido actual
   If trabajar_con_pedido Then
        .AddSimpleWhereClause "NUMERO", NUMERO_PEDIDO, , , LOGIC_AND
        .AddSimpleWhereClause "ALMORIG", codigo_almacen, , , LOGIC_AND
   Else
   
        .AddComplexWhereClause "(ESTADO = 1 OR ESTADO = 2)"
   
        Select Case TipoPermiso
           'usuario comun, ver solo los pedidos de su almacén
        Case 0
                .AddSimpleWhereClause "ALMORIG", AlmacenActual, , , LOGIC_AND
            'supervisor, ver todos los pedidos
            ' Case 1
        End Select
   
   End If
   '.AddSimpleWhereClause "ESTADO", 1
   '.AddSimpleWhereClause "ESTADO", 2, , , LOGIC_OR
     
   .AddOrderClause "NUMERO"
    
  cabSQL = .SQL
   
   'Debug.Print .SQL
   
  End With
    
  With locCnnSP
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnnSP
    End If
   End With
     
  Set adoPrimaryRS = New Recordset
  Set adoDetalleRS = New Recordset
  
  'sqlSHAPE = "SHAPE {" & cabSQL & "} AS ParentCMD APPEND ({" & detSQL & "} AS ChildCMD RELATE NUMERO TO NUMERO) AS ChildCMD"
  'adoPrimaryRS.Open sqlSHAPE, locCnnSP, adOpenDynamic, adLockBatchOptimistic
  
  'adoPrimaryRS.Open cabSQL, locCnnSP, adOpenDynamic, adLockBatchOptimistic
  
   adoPrimaryRS.Open cabSQL, locCnnSP, adOpenStatic, adLockOptimistic

  cbCODPROV.CodigoWidth = 800

  'obtener el IDTEM de la temporada actual
  'tmptemporada = devuelve_campo("select IDTEM from TEMPOR where ACTUAL = 1", locCnnSP)



  'temporada separado de los otros para que no me lo este borrando
  With Temporada
    .ConexionString = locCnnSP
    .LenCodigo = 3
    .SQLString = "SELECT IDTEM, ABREVIA FROM TEMPOR WHERE MBAJA = 0 ORDER BY IDTEM"
    .CodigoWidth = 500
    '.DataField = "CODALM"
    .carga
    'Set .DataSource = adoPrimaryRS
   End With
   
   With cbESTADO
    .DataField = "ESTADO"
    .LenCodigo = 1
    .CodigoWidth = 300
    .borra_combo
    .añade_item "1  PENDIENTE", 1
    .añade_item "2  PARCIAL", 2
    .añade_item "3  RECIBIDO", 3
    DoEvents
   Set .DataSource = adoPrimaryRS
 End With
 
  'si no hay registros
  If adoPrimaryRS.EOF Then Exit Sub
  'adoPrimaryRS.MoveFirst
  
  Call cmdFirst_Click
  DoEvents
  
  'cargar campos y grid
  Call Asigna_Campos
  
    fg.AllowSelection = True
    fg.HighLight = flexHighlightAlways
    fg.SelectionMode = flexSelectionByRow
    
  'asignar la temporada de trabajo como temporada actual
  TemporadaWrk = TemporadaActual
  'poner el texto
  lblTemporadaWrk.Caption = devuelve_campo("SELECT AÑO + '  ' + TEMPORADA FROM TEMPOR WHERE IDTEM = " & TemporadaWrk, locCnnSP)
    
    
  fg.SelectionMode = flexSelectionByRow
  fg.HighLight = flexHighlightAlways
  fg.AllowBigSelection = False
  fg.AllowUserFreezing = flexFreezeColumns

    
  Call SetButtons(True)
  
  DoEvents
  Show
  
  Do
  tem = InputBox("Introduzca temporada de trabajo", "Introducir temporada", TemporadaActual)
  
  Loop Until tem <> "" And IsNumeric(Trim(tem))
  
  Temporada.Text = tem
  NoAvisar = True
  Call cmCambTem_Click
  NoAvisar = False
    

End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      
    If editagrid = False Then cbcerrar_Click
    Case vbKeyEnd
      If editagrid = False Then cmdLast_Click
    Case vbKeyHome
      If editagrid = False Then cmdFirst_Click
    
 '   Case vbKeyPageUp
 '
  '    If editagrid = False Then
  '
  '      If Shift = vbCtrlMask Then
  '          cmdFirst_Click
  '      Else
  '          cmdPrevious_Click
  '      End If
  '
  '    End If
  '
  '    Case vbKeyPageDown
   '
   '   If editagrid = False Then
   '
   '   If Shift = vbCtrlMask Then
   '     cmdLast_Click
   '   Else
   '     cmdNext_Click
   ''   End If
      
    '  End If
    
    
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
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error Resume Next

tmptempor = ""
tmptalla = ""
tmpcolor = ""
sqlSHAPE = ""

If Not adoDetalleRS Is Nothing Then
   If adoDetalleRS.State = 1 Then adoDetalleRS.Close
   Set adoDetalleRS = Nothing
End If

If Not adoPrimaryRS Is Nothing Then
   If adoPrimaryRS.State = 1 Then adoPrimaryRS.Close
   Set adoPrimaryRS = Nothing
End If

If locCnnSP.State = 1 Then locCnnSP.Close

'If Me.MDIChild = True Then frmMenuTactil.Show
Set frmPedProv = Nothing
Set oSQL = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub ver_contenido_registro()

   On Error GoTo ver_contenido_registro_Error

If (adoPrimaryRS.EOF And adoPrimaryRS.BOF) Then Exit Sub
If IsNull(adoPrimaryRS.fields("NUMERO").Value) Then Exit Sub
If adoPrimaryRS.fields("NUMERO").Value = 0 Then Exit Sub

 fg.Clear
 
 With adoPrimaryRS
       If .State = 0 Then
       Unload Me
       Exit Sub
       End If
 
 If .AbsolutePosition > 0 Then
 lblstatus.Caption = "Registro: " & CStr(adoPrimaryRS.AbsolutePosition)
  
    ioDCTOPP.Text = .fields("DCTOPP")
    ioGASTOS.Text = .fields("GASTOS")
    ioPORTES.Text = .fields("PORTES")
    
    DoEvents
    
    If Not IsNull(.fields("COMENTARIO")) Then
        lblExisteCom.Visible = True
    Else
        lblExisteCom.Visible = False
    End If
    
   'si es para la caja A
   If .fields("DESTINO") = 0 Then
        lblStatusD.Caption = "PEDIDO EN A"
        btA.Value = True
        btB.Value = False
   'o para la caja B
   Else
        lblStatusD.Caption = "PEDIDO EN B"
        btB.Value = True
        btA.Value = False
   End If
   

 End If
 
 End With
 
 
If Not (adoPrimaryRS.EOF Or adoPrimaryRS.BOF) Then
    If adoDetalleRS.State = 1 Then adoDetalleRS.Close
    adoDetalleRS.Open "SELECT * FROM DETPEDPRO WHERE NUMERO = " & adoPrimaryRS.fields("NUMERO").Value & " AND ALMORIG = " & AlmacenActual, locCnnSP, adOpenDynamic, adLockOptimistic
Else
    Exit Sub
End If

Call Re_Asigna_Grid

   On Error GoTo 0
   Exit Sub

ver_contenido_registro_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ver_contenido_registro de Formulario frmPedProv"

End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  
Exit Sub
  
  'Esto mostrará la posición de registro actual para este Recordset
If (adoPrimaryRS.EOF And adoPrimaryRS.BOF) Then Exit Sub
If IsNull(adoPrimaryRS.fields("NUMERO").Value) Then Exit Sub
If adoPrimaryRS.fields("NUMERO").Value = 0 Then Exit Sub

 fg.Clear
 
 With adoPrimaryRS
       If .State = 0 Then
       Unload Me
       Exit Sub
       End If
 
 If .AbsolutePosition > 0 Then
 lblstatus.Caption = "Registro: " & CStr(adoPrimaryRS.AbsolutePosition)
  
    ioDCTOPP.Text = .fields("DCTOPP")
    ioGASTOS.Text = .fields("GASTOS")
    ioPORTES.Text = .fields("PORTES")
    
    DoEvents
    
    If Not IsNull(.fields("COMENTARIO")) Then
        lblExisteCom.Visible = True
    Else
        lblExisteCom.Visible = False
    End If
    
   'si es para la caja A
   If .fields("DESTINO") = 0 Then
        lblStatusD.Caption = "PEDIDO EN A"
        btA.Value = True
        btB.Value = False
   'o para la caja B
   Else
        lblStatusD.Caption = "PEDIDO EN B"
        btB.Value = True
        btA.Value = False
   End If
   

 End If
 
 End With
 
 
If Not (adoPrimaryRS.EOF Or adoPrimaryRS.BOF) Then
    If adoDetalleRS.State = 1 Then adoDetalleRS.Close
    adoDetalleRS.Open "SELECT * FROM DETPEDPRO WHERE NUMERO = " & adoPrimaryRS.fields("NUMERO").Value & " AND ALMORIG = " & AlmacenActual, locCnnSP, adOpenDynamic, adLockOptimistic
Else
    Exit Sub
End If

Call Re_Asigna_Grid
  
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
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
  
  Call Re_Asigna_Grid
  
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cbAgregar_Click()
Dim tmpcodigo As Variant
       
  On Error GoTo cbAgregar_Click_Error

  If Not Casignados Then Call Asigna_Campos
  If Not Gasignados Then Call Asigna_Grid
  
  'Call Asigna_Campos
 
  tmpcodigo = devuelve_campo("select max(NUMERO) + 1 from CABPEDPRO WHERE ALMORIG = " & AlmacenActual, locCnnSP)
 
  If tmpcodigo = "@" Then tmpcodigo = "1"

  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
      If .EditMode <> 0 Then .CancelUpdate
    End If
    
    fg.Clear
    fg.Rows = 1
    
    
    .AddNew
    lblstatus.Caption = "Agregar registro"
    mbAddNewFlag = True
    SetButtons False
    
    .fields("ALMORIG") = AlmacenActual
    .fields("NUMERO") = tmpcodigo
    .fields("ESTADO") = 1
    
    ioNUMERO.Caption = tmpcodigo
    'ioNUMERO.Locked = True
    ioFECHA.Text = Format(Date, "dd/mm/yyyy")
    
    cbCODPROV.SetFocus
  End With
  
  btA.Value = True

Set tmpcodigo = Nothing


   On Error GoTo 0
   Exit Sub

cbAgregar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbAgregar_Click de Formulario frmPedProv"

End Sub

Private Sub cbImprimir_Click()
Dim linea1 As String
Dim linea2 As String
         
    On Error GoTo cbImprimir_Click_Error

    DoEvents
    
    If adoPrimaryRS.RecordCount = 0 Then Exit Sub

    linea1 = "Informe de Pedido. Codigo: " & adoPrimaryRS.fields("NUMERO") & ". Proveedor: " & Trim(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & cbCODPROV.Text, locCnnSP)) & ". Destino: " & Trim(devuelve_campo("select descripcion from almacenes where codigo = " & cbCODALM.Text, locCnnSP))
    linea2 = "Usuario: " & Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & UsuarioActual, locCnnSP)) & ". Fecha: " & adoPrimaryRS.fields("FMODI")
    Call PrintFlexGrid(fg, 1, 1, 2, linea1, linea2, 12, 2)
  
   'With adoPrimaryRS
   '.Delete
   '.MoveNext
   'If .EOF Then .MoveLast
   'End With
   'Exit Sub

   On Error GoTo 0
   Exit Sub

cbImprimir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbImprimir_Click de Formulario frmPedProv"

End Sub





Private Sub cbedicion_Click()
  On Error GoTo EditErr

  If adoPrimaryRS.EOF Then Exit Sub
  


  lblstatus.Caption = "Modificar registro"
  mbEditFlag = True
  SetButtons False
  
  cbCODALM.Enabled = False
  cbCODALM.Locked = True
  
  cBFPAGO.SetFocus
  Exit Sub
  
EditErr:
  MsgBox Err.Description
End Sub
Private Sub cbCancelar_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False
  
  Tab1.Tab = 0
  

End Sub

Private Sub cbactualizar_Click()

   'On Error GoTo cbactualizar_Click_Error

   creando = True

    With adoPrimaryRS
        If .EOF Then Exit Sub
    
        '  If ioDCTO.Text = "" Then ioDCTO.Text = 0
        If ioDCTOPP.Text = "" Then ioDCTOPP.Text = 0
        If ioGASTOS.Text = "" Then ioGASTOS.Text = 0
        If ioPORTES.Text = "" Then ioPORTES.Text = 0
    
        
      '  .Fields("DCTO") = ioDCTO.Text
        .fields("DCTOPP") = ioDCTOPP.Text
        .fields("GASTOS") = ioGASTOS.Text
        .fields("PORTES") = ioPORTES.Text
    
    End With

    With ioFECHA
       If .Text = "" Then
         lblstatus.Caption = "No se permite fecha en blanco"
         .CancelarValidacion
         .SetFocus
         Exit Sub
       End If
    End With
    
    With cbCODPROV
       If .Text = "" Then
         lblstatus.Caption = "No se permite PROVEEDOR en blanco"
         .SetFocus
         Exit Sub
       End If
    End With
    
    With cBFPAGO
       If .Text = "" Then
         lblstatus.Caption = "No se permite FORMA DE PAGO en blanco"
         .SetFocus
         Exit Sub
       End If
    End With
        
    With cbCODALM
       If .Text = "" Then
         lblstatus.Caption = "No se permite ALMACEN en blanco"
         .SetFocus
         Exit Sub
       End If
    End With
    
    With adoPrimaryRS
        .fields("CODUSR") = UsuarioActual
        '.UpdateBatch adAffectAll
        .Update 'Batch adAffectAll
    End With
     
    DoEvents
    adoPrimaryRS.Requery
  'ioNUMERO.Locked = True
  
  If mbAddNewFlag Then
    adoPrimaryRS.fields("ESTADO") = 2  'parcial
    
    Call refresca_grid_externo(True)
    
    DoEvents
    Call cmdLast_Click
    'adoPrimaryRS.MoveLast              'va al nuevo registro
    
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  
  Tab1.Tab = 0
  
  DoEvents
  
  creando = False

   On Error GoTo 0
   Exit Sub

cbactualizar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbactualizar_Click de Formulario frmPedProv"

End Sub

Private Sub cbcerrar_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()

   On Error GoTo cmdFirst_Click_Error

  If adoPrimaryRS.BOF Or adoPrimaryRS.EOF Then Exit Sub
  adoPrimaryRS.MoveFirst
  
  adoPrimaryRS.Move 0
  mbDataChanged = False
  
  Call ver_contenido_registro
  
  'Call Re_Asigna_Grid

   On Error GoTo 0
   Exit Sub

cmdFirst_Click_Error:

    If Err.Number = -2147217842 Then Exit Sub
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmdFirst_Click de Formulario frmPedProv"
  
End Sub

Private Sub cmdLast_Click()
  
   On Error GoTo cmdLast_Click_Error

  If adoPrimaryRS.BOF Or adoPrimaryRS.EOF Then Exit Sub
  adoPrimaryRS.MoveLast
  mbDataChanged = False
  
  Call ver_contenido_registro
  'Call Re_Asigna_Grid

   On Error GoTo 0
   Exit Sub

cmdLast_Click_Error:
    
    If Err.Number = -2147217842 Then Exit Sub
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmdLast_Click de Formulario frmPedProv"

End Sub

Private Sub cmdNext_Click()
  
   On Error GoTo cmdNext_Click_Error

  If adoPrimaryRS.BOF Or adoPrimaryRS.EOF Then Exit Sub
  adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
     'ha sobrepasado el final; vuelva atrás
    adoPrimaryRS.MoveLast
  End If
  'muestra el registro actual
  mbDataChanged = False
  
  Call ver_contenido_registro
  'Call Re_Asigna_Grid
  
   On Error GoTo 0
   Exit Sub

cmdNext_Click_Error:
    
    If Err.Number = -2147217842 Then Exit Sub
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmdNext_Click de Formulario frmPedProv"

End Sub

Private Sub cmdPrevious_Click()
 
   On Error GoTo cmdPrevious_Click_Error

  If adoPrimaryRS.BOF Or adoPrimaryRS.EOF Then Exit Sub
  adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'ha sobrepasado el final; vuelva atrás
    adoPrimaryRS.MoveFirst
  End If
  'muestra el registro actual
  mbDataChanged = False
  
   Call ver_contenido_registro
   'Call Re_Asigna_Grid

   On Error GoTo 0
   Exit Sub

cmdPrevious_Click_Error:

    If Err.Number = -2147217842 Then Exit Sub
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmdPrevious_Click de Formulario frmPedProv"

End Sub


Private Sub SetButtons(bVal As Boolean)
  cbAgregar.Visible = bVal
  cbEdicion.Visible = bVal
  
'  cmEtiquetas.Visible = bVal
  cmEntrada.Visible = bVal
  
  cbActualizar.Visible = Not bVal
  cbCancelar.Visible = Not bVal
  cbImprimir.Visible = bVal
  cbCerrar.Visible = bVal
  
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
  
  If mbEditFlag Then
    cbCODPROV.Locked = True
  Else
    cbCODPROV.Locked = False
  End If
  
  cBFPAGO.Locked = bVal
  cbPLAZOE.Locked = bVal
  cbCODALM.Locked = bVal
  cbESTADO.Locked = bVal
  
  Tab1.Enabled = Not bVal
  cmComentario.Enabled = Not bVal
  
  
'  Call Habilita_miTextNum(Me, Not bVal)
  
End Sub

Private Sub ioDCTOPP_Validate(Cancel As Boolean)

If mbEditFlag Or mbAddNewFlag Then
    
With ioDCTOPP

    If Trim(.Text) <> "" Then
        If CDbl(Replace(.Text, ".", ",")) >= 100 Then
            lblstatus.Caption = "No se permite un descuento por pronto pago del 100%"
            .CancelarValidacion
            Cancel = True
            DoEvents
            Exit Sub
        
        Else
            lblstatus.Caption = ""
            DoEvents
        End If
    End If

End With

End If
 
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CreateDBEtiquetas
' DateTime  : 10/11/2003 20:36
' Author    : Administrador
' Purpose   : Rutina que crea la base de datos temporal donde se alma
'             cenaran los registros que se van al imprimir como etiquetas
'             (un registro por cada unidad articulo/talla/color.
'---------------------------------------------------------------------------------------
'
Private Sub CreateDBEtiquetas(rc_detped As ADODB.Recordset, digitos As String, saltar As Byte)
On Error GoTo ErrorCreateDB

Dim Cat     As New ADOX.Catalog
Dim Tbl(6) As ADOX.Table
Dim Idx()   As ADOX.Index
Dim msgErrR As Integer
Dim sCnn    As String
Const fichero = "c:\TempEtiquetasDB.mdb"
Dim etiqrc As New ADODB.Recordset
Dim nveces As Long
Dim tmpean As String * LenCodBar
Dim tmpvar As Byte

Dim tmp_precom As Long

ChDir ("c:\")

'si existe uno previo, borrar
If Dir(fichero) <> "" Then Kill fichero
sCnn = strCnnMdb & fichero
'sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Engine Type=5;Data Source=" & fichero

Cat.Create sCnn

  '----------* Table Definition of ETIQUETAS *----------
  Set Tbl(0) = New ADOX.Table
  Tbl(0).ParentCatalog = Cat
  With Tbl(0)
    .Name = "ETIQUETAS"
    .Columns.Append "ABREVIA", adVarWChar, 20
    .Columns.Append "CODCOLOR", adVarWChar, 3
      .Columns("CODCOLOR").Properties("Default").Value = "0"
    .Columns.Append "CODIGO", adVarWChar, 5
      .Columns("CODIGO").Properties("Default").Value = "0"
    .Columns.Append "CODTALLA", adVarWChar, 2
      .Columns("CODTALLA").Properties("Default").Value = "0"
    .Columns.Append "DESCOLOR", adVarWChar, 15
    .Columns.Append "DESCTALLA", adVarWChar, 15
    .Columns.Append "MODELO", adVarWChar, 30
    .Columns.Append "REFERENCIA", adVarWChar, 15
    .Columns.Append "Id", adInteger
      .Columns("Id").Properties("AutoIncrement").Value = True
      .Columns("Id").Properties("Nullable").Value = False
    
    .Columns.Append "PRECOM", adVarWChar, 6
   
    
    .Columns.Append "PROVEEDOR", adVarWChar, 3
    '.Columns("PROVEEDOR").Properties("Default").Value = ""
      
    .Columns.Append "PVP", adCurrency
      .Columns("PVP").Properties("Default").Value = "0"
    .Columns.Append "TEMPOR", adVarWChar, 3
      .Columns("TEMPOR").Properties("Default").Value = "0"
      
      .Columns.Append "IMAGEN", adLongVarBinary
      .Columns("IMAGEN").Properties("Description").Value = "IMAGEN"

  End With
  '----------* Index Defini4tions of ETIQUETAS *----------
  ReDim Idx(0)
  Set Idx(0) = New ADOX.Index
    Idx(0).Name = "PrimaryKey"
    Idx(0).PrimaryKey = True
    Idx(0).Unique = True
      Idx(0).Columns.Append "Id"
  Tbl(0).Indexes.Append Idx(0)

  Cat.tables.Append Tbl(0)

  Set Cat = Nothing
  
'tempor
'codart
'codtalla
'codcol
'preven
  
  etiqrc.Open "SELECT * FROM ETIQUETAS", sCnn, adOpenDynamic, adLockOptimistic
  
  
  'BarCodefrm.BarCode1.DataField = "IMAGEN"
  'Set BarCodefrm.BarCode1.DataSource = etiqrc
  
  'meterle etiquetas en blanco si hay para saltar
  
  If saltar > 0 Then
  
    For tmpvar = 0 To saltar - 1
        
        etiqrc.AddNew
            etiqrc.fields("ABREVIA") = " "
            etiqrc.fields("TEMPOR") = " "
            etiqrc.fields("CODIGO") = " "
            etiqrc.fields("CODTALLA") = " "
            etiqrc.fields("CODCOLOR") = " "
            etiqrc.fields("PRECOM") = " "
            etiqrc.fields("PVP") = "0"
            etiqrc.fields("PROVEEDOR") = " "
            etiqrc.fields("MODELO") = " "
            etiqrc.fields("DESCOLOR") = " "
            etiqrc.fields("DESCTALLA") = " "
            etiqrc.fields("REFERENCIA") = " "
            GuardarArchivo etiqrc.fields("IMAGEN"), App.Path & "\Blanca.bmp"
           
        etiqrc.Update
    
    Next tmpvar
  
  End If
   

 ' frmBarcode.Show
  DoEvents
  

  'ahora meterle los datos ....
  With rc_detped
    
    If Not .BOF Then .MoveFirst
    
    Do Until .EOF
 
        'si estan pendientes de meter en el almacén
        If .fields("METIDO") = 0 Then
            'un registro por cada unidad para el mismo articulo
            For nveces = 1 To .fields("UNIDADES").Value
    
            'CODIGO DE BARRAS 13 DIGITOS:
            
            'ARTICULO: 5 digitos
            'TEMPORADA:3 digitos
            'TALLA:    2 digitos
            'COLOR:    3 digitos
                        
                        
            tmpean = Format(.fields("CODART"), "00000") & Format(.fields("TEMPOR"), "000") & Format(.fields("CODTALLA"), "00") & Format(.fields("CODCOL"), "000")
   
            'Debug.Print tmpean
            
            'PaintCode BarCodefrm, Mid$(tmpean, 1, 1), Mid$(tmpean, 2, 6), Mid$(tmpean, 8, 6)

           ' frmBarcode.txtData.Text = tmpean
           ' frmBarcode.DrawBarCode "128"
           ' frmBarcode.cmdBMP_Click
           ' DoEvents
            
            BarCodefrm.BarCode1.DataToEncode = tmpean
            DoEvents
            Set BarCodefrm.Picture1.Picture = BarCodefrm.BarCode1.Picture
            DoEvents
            'BarCodefrm.barcode1.Picture
            
            SavePicture BarCodefrm.Picture1.Image, App.Path & "\Barcode.bmp"
            'BarCodefrm.BarCode1.SaveBarCode "c:\pruebasav.wmf"
            DoEvents
            
            'Set BarCodefrm.Picture1.Picture = LoadPicture("c:\pruebasav.wmf")
            
            DoEvents

            'Set BarCodefrm.Picture1.Picture = BarCodefrm.BarCode1.Picture
            
            etiqrc.AddNew
            
            GuardarArchivo etiqrc.fields("IMAGEN"), App.Path & "\Barcode.bmp"
            'GuardarBinary etiqrc.Fields("IMAGEN"), BarCodefrm.Picture1
            'GuardarBinary etiqrc.Fields("IMAGEN"), BarCodefrm.BarCode1.Picture
            etiqrc.fields("ABREVIA") = devuelve_campo("SELECT ABREVIA FROM TEMPOR WHERE IDTEM =" & .fields("TEMPOR"), locCnnSP)
            etiqrc.fields("TEMPOR") = Format(.fields("TEMPOR"), "000")
            etiqrc.fields("CODIGO") = Format(.fields("CODART"), "00000")
            etiqrc.fields("CODTALLA") = Format(.fields("CODTALLA"), "00")
            etiqrc.fields("CODCOLOR") = Format(.fields("CODCOL"), "000")
            
            'tmp_precom = Obtiene_Precom_Pedido(.Fields("CODART"), .Fields("TEMPOR"), .Fields("CODTALLA"), .Fields("CODCOL"), locCnnSP)
            
            tmp_precom = devuelve_campo("SELECT PRECOM FROM MAARTIC WHERE CODIGO =" & .fields("CODART") & " and TEMPOR = " & .fields("TEMPOR"), locCnnSP)
            
            etiqrc.fields("PRECOM") = Mid(digitos, 1, 2) & Format(tmp_precom, "00")
            etiqrc.fields("PVP") = devuelve_campo("SELECT PREVEN FROM MAARTIC WHERE CODIGO =" & .fields("CODART") & " and TEMPOR = " & .fields("TEMPOR"), locCnnSP)
            etiqrc.fields("PROVEEDOR") = Mid(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & cbCODPROV.Text, locCnnSP), 1, 3)
            etiqrc.fields("MODELO") = devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO =" & .fields("CODART") & " and TEMPOR = " & .fields("TEMPOR"), locCnnSP)
            etiqrc.fields("DESCOLOR") = devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO =" & .fields("CODCOL"), locCnnSP)
            etiqrc.fields("DESCTALLA") = devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO =" & .fields("CODTALLA"), locCnnSP)
            etiqrc.fields("REFERENCIA") = devuelve_campo("SELECT REF FROM MAARTIC WHERE CODIGO =" & .fields("CODART") & " and TEMPOR = " & .fields("TEMPOR"), locCnnSP)
            etiqrc.Update
        
            Next
         End If
         
        .MoveNext
    
    Loop

  End With
  
  etiqrc.Close
  Set etiqrc = Nothing
  
  tmpean = ""
  
  Exit Sub

ErrorCreateDB:
    msgErrR = MsgBox("    Error No. " & Err & " " & vbCrLf & Error, vbCritical + vbAbortRetryIgnore, "Code Gen Error")
    Select Case msgErrR
      Case Is = vbAbort
      If Not (Cat Is Nothing) Then
        Set Cat = Nothing
      End If
      Exit Sub
     Case Is = vbRetry
       Resume Next
     Case Is = vbIgnore
       Resume
    End Select

End Sub


'---------------------------------------------------------------------------------------
' Procedure : generar_etiquegtas
' DateTime  : 09/11/2003 21:45
' Author    : Administrador
' Purpose   : Genera las etiquetas y el informe Crystal Report
'---------------------------------------------------------------------------------------
'
Private Sub generar_etiquetas(conn As ADODB.Connection)
'Dim tmprcped As New ADODB.Recordset

   On Error GoTo generar_etiquegtas_Error

   'crear la base de datos etiquetas'
   
    'Me.refresca_grid_externo
    'If adoPrimaryRS.RecordCount = 0 Then Exit Sub
      
   'If MsgBox("¿Desea generar las etiquetas para el pedido actual nº " & ioNUMERO.Caption, vbQuestion + vbYesNo) = vbYes Then
    
        'tmprcped.Open "SELECT * FROM DETPEDPRO WHERE NUMERO = " & ioNUMERO.Caption & " AND ALMORIG = " & AlmacenActual, conn
    
        DoEvents
        Call procesa_informes(1, True)
        DoEvents
        
           
   'End If
  
      '  With locCnnSP
      '      .CommitTrans
       '     .Close
            'entrans = False
       '     .CursorLocation = adUseClient
       '     .Open strLocCnnSP
       ' End With
        
'        adoPrimaryRS.Open
         
      '  Call refresca_grid_externo(False)
        
      '  adoPrimaryRS.Open cabSQL, locCnnSP, adOpenStatic, adLockOptimistic
      '  adoPrimaryRS.Requery
       ' DoEvents
      '  Call Asigna_Grid
    
        'tmprcped.Close
       ' Set tmprcped = Nothing
   
   
   
  
  ' Set tmprcped = Nothing
   On Error GoTo 0
   Exit Sub

generar_etiquegtas_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generar_etiquetas of Formulario TmpMaDet"

End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : comprueba_tallas_colores
' Fecha/Hora    : 19/11/2003 12:55
' Autor         : JCastillo
' Propósito     : Comprueba q no haya tallas o colores en blanco o nulo antes
'                     de sacar etiquetas o dar entrada. Devuelve TRUE si hay alguna combinación
'                    codtalla o codcol = 0 o igual a NULL
'---------------------------------------------------------------------------------------
'
Private Function comprueba_tallas_colores() As Boolean
Dim rc As New ADODB.Recordset

   On Error GoTo comprueba_tallas_colores_Error

        rc.Open "SELECT * FROM DETPEDPRO WHERE  (NUMERO = " & ioNUMERO.Caption & " and ALMORIG = " & AlmacenActual & ") and ((CODTALLA = 0 or CODTALLA IS NULL) or (CODCOL = 0 or CODCOL IS NULL))", locCnnSP, adOpenDynamic, adLockReadOnly
                   
        If rc.EOF Then
            
            
            comprueba_tallas_colores = False
        
        Else
        
            lblstatus.Caption = "Faltan por asignar tallas y colores. (Articulo: " & Format(rc.fields("CODART").Value, "00000") & ". Linea: " & Format(rc.fields("LINEA").Value, "00000") & ")"
            comprueba_tallas_colores = True
        
        End If
        
        rc.Close

   On Error GoTo 0
   Exit Function

comprueba_tallas_colores_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento comprueba_tallas_colores de Formulario frmPedProv"
End Function

'---------------------------------------------------------------------------------------
' Procedimiento : carga_grid_totales
' Fecha/Hora     : 27/11/2003 16:34
' Autor             : JCastillo
' Propósito       : Carga el GRID de los totales y Subtotales (tab totales)
'---------------------------------------------------------------------------------------
Private Sub carga_grid_totales()
Dim tmprc As New ADODB.Recordset
Dim total_general As Double
Dim total_neto As Double
Dim total_iva As Double
Dim tmpdctopp As Double
Dim tmpgastos As Double
Dim tmpportes As Double

On Error GoTo carga_grid_totales_Error

If ioNUMERO.Caption = "" Then Exit Sub

tmpsubtot = 0

'obtener el dcto pp
If ioDCTOPP.Text <> "" Then
    tmpdctopp = Replace(ioDCTOPP.Valor, ".", ",")
    tmprc.Open "select IVA, sum((unidades * precom) - ((unidades * precom)* dcto/100) - ((unidades * precom) - ((unidades * precom) * dcto/100)) * " & tmpdctopp & "/100) as NETO from detpedpro where (numero = " & ioNUMERO.Caption & " AND ALMORIG = " & AlmacenActual & ") group by iva", locCnnSP, adOpenDynamic, adLockReadOnly
Else
    tmprc.Open "select IVA, sum((unidades * precom) - ((unidades * precom)* dcto/100) as NETO from detpedpro where (numero = " & ioNUMERO.Caption & " and ALMORIG = " & AlmacenActual & ") group by iva", locCnnSP, adOpenDynamic, adLockReadOnly
End If
'group by iva

With vsTotales

        .Rows = 1
        .Cols = 5
        .Clear
        .TextMatrix(0, 1) = "IVA %"
        .TextMatrix(0, 2) = "NETO"
        .TextMatrix(0, 3) = "IMPORTE IVA"
        .TextMatrix(0, 4) = "NETO + IVA"
        

    Do Until tmprc.EOF
    
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 1) = tmprc.fields("IVA")
        .TextMatrix(.Rows - 1, 2) = tmprc.fields("NETO")  'base imponible
        .TextMatrix(.Rows - 1, 3) = (tmprc.fields("NETO") * tmprc.fields("IVA")) / 100  'importe de IVA
        .TextMatrix(.Rows - 1, 4) = tmprc.fields("NETO") + ((tmprc.fields("NETO") * tmprc.fields("IVA")) / 100) 'importe mas iva
        
        total_general = total_general + tmprc.fields("NETO") + ((tmprc.fields("NETO") * tmprc.fields("IVA")) / 100)
        total_neto = total_neto + tmprc.fields("NETO")
        total_iva = total_iva + (tmprc.fields("NETO") * tmprc.fields("IVA")) / 100
          
        tmprc.MoveNext
        
    Loop
    
    .subtotal flexSTSum, , 2, , vbBlue, vbWhite
    .subtotal flexSTSum, , 3, , vbBlue, vbWhite
    .subtotal flexSTSum, , 4, , vbBlue, vbWhite, True
    
    .ColFormat(2) = "Currency"
    .ColFormat(3) = "Currency"
    .ColFormat(4) = "Currency"
    
   ' .TextMatrix(1, 3) = "Total"
    .AutoSize 1, .Cols - 1
    
'If ioDCTOPP.Text <> "" Then
'    lblDctoPP.Caption = Replace(ioDCTOPP.Text, ".", ",") & " %"
'End If

If ioPORTES.Text <> "" Then
    If ioPORTES.Valor <> "" Then tmpportes = Replace(ioPORTES.Valor, ".", ",")
    lblPortes.Caption = ioPORTES.Text
End If

If ioGASTOS.Text <> "" Then
    If ioGASTOS.Valor <> "" Then tmpgastos = Replace(ioGASTOS.Valor, ".", ",")
    lblGastos.Caption = ioGASTOS.Text
End If
    
    'base imponible
    lblSubTotal.Caption = Format(total_neto, "Currency")
    'iva
    lblIVA.Caption = Format(total_iva, "Currency")
    'ahora sumar PORTES y IVA para hallar el total
    lblTotal.Caption = Format(total_neto + total_iva + tmpportes + tmpgastos, "Currency") ' sumar portes y gastos

End With


'' Insertar totales en sus campos
With adoPrimaryRS

    'If mbEditFlag Or mbAddNewFlag Then
    
        'si no estan los totales correctamente en el registro ...
        If (.fields("TOTALIVA") <> Round(total_iva, 2)) Or (.fields("TOTALNET") <> Round(total_neto, 2)) Then
            
            'mandar un update
            locCnnSP.Execute "UPDATE CABPEDPRO SET TOTALIVA = " & Replace(Round(total_iva, 2), ",", ".") & " , TOTALNET = " & Replace(Round(total_neto, 2), ",", ".") & " WHERE NUMERO = " & adoPrimaryRS.fields("NUMERO") & " AND ALMORIG = " & adoPrimaryRS.fields("ALMORIG")

        End If
        
        DoEvents
        'Call cbactualizar_Click
        '.Update
   ' Else
       ' .Fields("TOTALIVA") = Round(total_iva, 2)
       ' .Fields("TOTALNET") = Round(total_neto, 2)
       ' adoPrimaryRS.UpdateBatch adAffectAll
   ' End If
    
End With

tmprc.Close
Set tmprc = Nothing

tmpsubtot = total_neto 'base imponible

   On Error GoTo 0
   Exit Sub

carga_grid_totales_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_grid_totales de Formulario frmPedProv"
End Sub



'---------------------------------------------------------------------------------------
' Subrutina   : ioSUCODIGO_Validate
' Fecha/Hora  : 02/12/2003 21:09
' Autor       : JCASTILLO
' Propósito   : Valida que no exista ya un codigo de usuario igual en la DB
'---------------------------------------------------------------------------------------
Private Sub ioSUCODIGO_Validate(Cancel As Boolean)

   On Error GoTo ioSUCODIGO_Validate_Error

If Not mbEditFlag And Not mbAddNewFlag Then Exit Sub

'With ioSUCODIGO
'If Trim(.Text <> "") Then

If Not IsDate(ioSUCODIGO.Text) Then
    ioSUCODIGO.CancelarValidacion
    Cancel = True
Else
    ioSUCODIGO.Text = Format(ioSUCODIGO.Text, "dd/mm/yyyy")
End If

    'comprueba si existe ...
  '  If devuelve_campo("SELECT NUMERO from CABPEDPRO WHERE SUCODIGO = '" & Trim(ioSUCODIGO.Text) & "' AND NUMERO <> " & adoPrimaryRS.fields("NUMERO") & " AND ALMORIG = " & AlmacenActual, locCnnSP) <> "@" Then

  '      lblstatus.Caption = "YA EXISTE ese codigo en la base de datos"
 '       .SetFocus
  '      .CancelarValidacion
  '      Cancel = True
    
  '  Else
    
    '    lblstatus.Caption = ""
    
  '  End If

'End If
'End With

   On Error GoTo 0
   Exit Sub

ioSUCODIGO_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioSUCODIGO_Validate de Formulario frmPedProv"

End Sub

Private Sub ioTRNSPORTI_Validate(Cancel As Boolean)
Tab1.Tab = 0
End Sub



Private Sub Temporada_Validate(Cancel As Boolean)
If Temporada.Text <> "" Then
    Call cmCambTem_Click
End If
End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : vsTotales_Click
' Fecha/Hora    : 03/12/2003 13:46
' Autor         : JCastillo
' Propósito     : Llamar al formulario de calculo de IVA, y cambiar el iva para
'                     este pedido si es necesario
'---------------------------------------------------------------------------------------
Private Sub vsTotales_dblClick()

On Error GoTo vsTotales_Click_Error

With frmCalcIVA
    .Total_General_Sin_IVA = tmpsubtot
    .Show 1
    
    If .Porcen_IVA > 0 Then
    
      'un update para cambiar los tipos de IVA del pedido actual
      locCnnSP.Execute "UPDATE DETPEDPRO SET IVA = " & .Porcen_IVA & " WHERE NUMERO = " & adoPrimaryRS.fields("NUMERO")
            
      'refrescar los datos del grid
      Call refresca_grid_externo(True)
        
    End If
    
End With

   On Error GoTo 0
   Exit Sub

vsTotales_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento vsTotales_Click de Formulario frmPedProv"

End Sub
