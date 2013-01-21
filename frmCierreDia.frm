VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmCierreDia 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cálculo de datos del cierre"
   ClientHeight    =   7185
   ClientLeft      =   2145
   ClientTop       =   1770
   ClientWidth     =   11520
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   11520
   Begin PCGestion.chameleonButton cbCancelar 
      Height          =   795
      Left            =   6225
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6360
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
      MICON           =   "frmCierreDia.frx":0000
      PICN            =   "frmCierreDia.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblEfectivo 
      Height          =   405
      Left            =   3855
      Top             =   600
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   714
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
      Colour1         =   11513775
      Colour2         =   16558731
      CaptionAlignment=   2
   End
   Begin PCGestion.bsGradientLabel lblVales 
      Height          =   345
      Left            =   3855
      Top             =   1410
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   609
      Caption         =   ""
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
      Colour1         =   11513775
      Colour2         =   16558731
      CaptionAlignment=   2
   End
   Begin PCGestion.bsGradientLabel lblTarjeta 
      Height          =   375
      Left            =   3855
      Top             =   1020
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   661
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
      Colour1         =   11513775
      Colour2         =   16558731
      CaptionAlignment=   2
   End
   Begin PCGestion.bsGradientLabel lblDevol 
      Height          =   360
      Left            =   60
      Top             =   1995
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   635
      Caption         =   ""
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
      Colour1         =   11513775
      Colour2         =   16558731
      CaptionAlignment=   2
   End
   Begin PCGestion.bsGradientLabel lblArreglos 
      Height          =   390
      Left            =   30
      Top             =   5025
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   688
      Caption         =   ""
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
      Colour1         =   11513775
      Colour2         =   16558731
      CaptionAlignment=   2
   End
   Begin PCGestion.bsGradientLabel lblIngresos 
      Height          =   375
      Left            =   3855
      Top             =   2670
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   661
      Caption         =   ""
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
      Colour1         =   11513775
      Colour2         =   16558731
      CaptionAlignment=   2
   End
   Begin PCGestion.bsGradientLabel lblPagos 
      Height          =   390
      Left            =   3855
      Top             =   1770
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   688
      Caption         =   ""
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
      Colour1         =   11513775
      Colour2         =   16558731
      CaptionAlignment=   2
   End
   Begin PCGestion.bsGradientLabel lblDeudCli 
      Height          =   360
      Left            =   60
      Top             =   2370
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   635
      Caption         =   ""
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
      Colour1         =   11513775
      Colour2         =   16558731
      CaptionAlignment=   2
   End
   Begin PCGestion.miText ioFECHA 
      Height          =   495
      Left            =   810
      TabIndex        =   0
      Top             =   30
      Width           =   1425
      _extentx        =   2514
      _extenty        =   873
      font            =   "frmCierreDia.frx":08F6
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.chameleonButton cmCalcular 
      Height          =   420
      Left            =   2250
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   570
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   741
      BTYPE           =   3
      TX              =   "Calcular"
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
      MICON           =   "frmCierreDia.frx":0922
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblEnCaja 
      Height          =   435
      Left            =   30
      Top             =   5445
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   767
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   16558731
      Colour2         =   14457707
      CaptionAlignment=   2
      ShadowDKColour  =   16711680
      TextShadowColour=   16776960
      TextShadow      =   -1  'True
      TextShadowYOffset=   1
      TextShadowXOffset=   1
   End
   Begin PCGestion.miText ioB5 
      Height          =   495
      Left            =   9915
      TabIndex        =   10
      Top             =   480
      Width           =   1425
      _extentx        =   2514
      _extenty        =   873
      font            =   "frmCierreDia.frx":093E
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioB10 
      Height          =   495
      Left            =   9915
      TabIndex        =   11
      Top             =   960
      Width           =   1425
      _extentx        =   2514
      _extenty        =   873
      font            =   "frmCierreDia.frx":096A
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioB20 
      Height          =   495
      Left            =   9915
      TabIndex        =   12
      Top             =   1440
      Width           =   1425
      _extentx        =   2514
      _extenty        =   873
      font            =   "frmCierreDia.frx":0996
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioB50 
      Height          =   495
      Left            =   9915
      TabIndex        =   13
      Top             =   1920
      Width           =   1425
      _extentx        =   2514
      _extenty        =   873
      font            =   "frmCierreDia.frx":09C2
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioB100 
      Height          =   495
      Left            =   9915
      TabIndex        =   14
      Top             =   2400
      Width           =   1425
      _extentx        =   2514
      _extenty        =   873
      font            =   "frmCierreDia.frx":09EE
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioB200 
      Height          =   495
      Left            =   9915
      TabIndex        =   15
      Top             =   2880
      Width           =   1425
      _extentx        =   2514
      _extenty        =   873
      font            =   "frmCierreDia.frx":0A1A
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioB500 
      Height          =   495
      Left            =   9915
      TabIndex        =   16
      Top             =   3360
      Width           =   1425
      _extentx        =   2514
      _extenty        =   873
      font            =   "frmCierreDia.frx":0A46
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioM1 
      Height          =   495
      Left            =   7665
      TabIndex        =   2
      Top             =   495
      Width           =   1425
      _extentx        =   2514
      _extenty        =   873
      font            =   "frmCierreDia.frx":0A72
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioM2 
      Height          =   495
      Left            =   7665
      TabIndex        =   3
      Top             =   960
      Width           =   1425
      _extentx        =   2514
      _extenty        =   873
      font            =   "frmCierreDia.frx":0A9E
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioM5 
      Height          =   495
      Left            =   7665
      TabIndex        =   4
      Top             =   1425
      Width           =   1425
      _extentx        =   2514
      _extenty        =   873
      font            =   "frmCierreDia.frx":0ACA
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioM10 
      Height          =   495
      Left            =   7665
      TabIndex        =   5
      Top             =   1890
      Width           =   1425
      _extentx        =   2514
      _extenty        =   873
      font            =   "frmCierreDia.frx":0AF6
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioM20 
      Height          =   495
      Left            =   7665
      TabIndex        =   6
      Top             =   2340
      Width           =   1425
      _extentx        =   2514
      _extenty        =   873
      font            =   "frmCierreDia.frx":0B22
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioM50 
      Height          =   495
      Left            =   7665
      TabIndex        =   7
      Top             =   2790
      Width           =   1425
      _extentx        =   2514
      _extenty        =   873
      font            =   "frmCierreDia.frx":0B4E
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioM100 
      Height          =   495
      Left            =   7665
      TabIndex        =   8
      Top             =   3240
      Width           =   1425
      _extentx        =   2514
      _extenty        =   873
      font            =   "frmCierreDia.frx":0B7A
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioM200 
      Height          =   495
      Left            =   7665
      TabIndex        =   9
      Top             =   3690
      Width           =   1425
      _extentx        =   2514
      _extenty        =   873
      font            =   "frmCierreDia.frx":0BA6
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel1 
      Height          =   420
      Left            =   7680
      Top             =   30
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   741
      Caption         =   "Monedas"
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
   Begin PCGestion.bsGradientLabel bsGradientLabel2 
      Height          =   420
      Left            =   9930
      Top             =   30
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   741
      Caption         =   "Billetes"
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
   Begin PCGestion.bsGradientLabel lblTotalB 
      Height          =   435
      Left            =   9360
      Top             =   3945
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   767
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
      CaptionAlignment=   2
   End
   Begin PCGestion.bsGradientLabel lblTotalM 
      Height          =   420
      Left            =   7215
      Top             =   4350
      Width           =   1995
      _ExtentX        =   3519
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
      CaptionAlignment=   2
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   2340
      Left            =   7200
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   4815
      Width           =   4305
      _cx             =   7594
      _cy             =   4128
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
      FormatString    =   $"frmCierreDia.frx":0BD2
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
   Begin PCGestion.chameleonButton cmCierreCaja 
      Height          =   795
      Left            =   5235
      TabIndex        =   36
      Top             =   6360
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "C&ERRAR CAJA"
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
      MICON           =   "frmCierreDia.frx":0C77
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblSumaMB 
      Height          =   405
      Left            =   90
      Top             =   6690
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   714
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   11250603
      Colour2         =   15640462
      Colour4         =   16777215
      CaptionAlignment=   2
   End
   Begin PCGestion.bsGradientLabel lblMovimientosCaja 
      Height          =   375
      Left            =   3855
      Top             =   3075
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   661
      Caption         =   ""
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
      Colour1         =   11513775
      Colour2         =   16558731
      CaptionAlignment=   2
   End
   Begin PCGestion.bsGradientLabel lblDescuadre 
      Height          =   420
      Left            =   2280
      Top             =   6675
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   741
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   11250603
      Colour2         =   15640462
      Colour4         =   16777215
      CaptionAlignment=   2
   End
   Begin PCGestion.bsGradientLabel lblCajaA 
      Height          =   360
      Left            =   3855
      Top             =   5430
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   635
      Caption         =   ""
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
      Colour1         =   11513775
      Colour2         =   16558731
      CaptionAlignment=   2
   End
   Begin PCGestion.bsGradientLabel lblCajaB 
      Height          =   375
      Left            =   3855
      Top             =   5805
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   661
      Caption         =   ""
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
      Colour1         =   11513775
      Colour2         =   16558731
      CaptionAlignment=   2
   End
   Begin PCGestion.bsGradientLabel lblTotalCobrosRealizados 
      Height          =   465
      Left            =   45
      Top             =   4470
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   820
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   12632256
      Colour2         =   14457707
      CaptionAlignment=   2
      ShadowDKColour  =   16711680
      TextShadowColour=   16776960
      TextShadow      =   -1  'True
      TextShadowYOffset=   1
      TextShadowXOffset=   1
   End
   Begin PCGestion.bsGradientLabel lblVentasNetas 
      Height          =   420
      Left            =   60
      Top             =   3525
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   741
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   12632256
      Colour2         =   14457707
      CaptionAlignment=   2
      ShadowDKColour  =   16711680
      TextShadowColour=   16776960
      TextShadow      =   -1  'True
      TextShadowYOffset=   1
      TextShadowXOffset=   1
   End
   Begin PCGestion.bsGradientLabel lblSaldoCajaEfectivo 
      Height          =   405
      Left            =   3840
      Top             =   2175
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   714
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   12632256
      Colour2         =   14457707
      CaptionAlignment=   2
      ShadowDKColour  =   16711680
      TextShadowColour=   16776960
      TextShadow      =   -1  'True
      TextShadowYOffset=   1
      TextShadowXOffset=   1
   End
   Begin PCGestion.bsGradientLabel lblValesEmitidos 
      Height          =   390
      Left            =   60
      Top             =   4050
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   688
      Caption         =   ""
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
      Colour1         =   11513775
      Colour2         =   16558731
      CaptionAlignment=   2
   End
   Begin PCGestion.bsGradientLabel lblDeudCobradas 
      Height          =   345
      Left            =   60
      Top             =   2760
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   609
      Caption         =   ""
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
      Colour1         =   11513775
      Colour2         =   16558731
      CaptionAlignment=   2
   End
   Begin PCGestion.bsGradientLabel lblVentasBrutas 
      Height          =   450
      Left            =   60
      Top             =   1140
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   794
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   12632256
      Colour2         =   14457707
      CaptionAlignment=   2
      ShadowDKColour  =   16711680
      TextShadowColour=   16776960
      TextShadow      =   -1  'True
      TextShadowYOffset=   1
      TextShadowXOffset=   1
   End
   Begin PCGestion.bsGradientLabel lblTotalDcto 
      Height          =   360
      Left            =   60
      Top             =   1605
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   635
      Caption         =   ""
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
      Colour1         =   11513775
      Colour2         =   16558731
      CaptionAlignment=   2
   End
   Begin PCGestion.miCombo cbCAJAS 
      Height          =   495
      Left            =   2970
      TabIndex        =   39
      Top             =   15
      Visible         =   0   'False
      Width           =   4110
      _extentx        =   7223
      _extenty        =   873
      font            =   "frmCierreDia.frx":0C93
   End
   Begin PCGestion.miText ioFFINAL 
      Height          =   495
      Left            =   810
      TabIndex        =   1
      Top             =   540
      Visible         =   0   'False
      Width           =   1425
      _extentx        =   2514
      _extenty        =   873
      font            =   "frmCierreDia.frx":0CBF
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.bsGradientLabel lblDiferenciasPVP 
      Height          =   360
      Left            =   60
      Top             =   3135
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   635
      Caption         =   ""
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
      Colour1         =   11513775
      Colour2         =   16558731
      CaptionAlignment=   2
   End
   Begin PCGestion.bsGradientLabel lblValesDctoAceptados 
      Height          =   345
      Left            =   3855
      Top             =   3465
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   609
      Caption         =   ""
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
      Colour1         =   11513775
      Colour2         =   16558731
      CaptionAlignment=   2
   End
   Begin PCGestion.bsGradientLabel lblValesDctoEmitidos 
      Height          =   375
      Left            =   3855
      Top             =   3840
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   661
      Caption         =   ""
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
      Colour1         =   11513775
      Colour2         =   16558731
      CaptionAlignment=   2
   End
   Begin PCGestion.bsGradientLabel lblArreCon 
      Height          =   360
      Left            =   3855
      Top             =   4245
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   635
      Caption         =   ""
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
      Colour1         =   11513775
      Colour2         =   16558731
      CaptionAlignment=   2
   End
   Begin PCGestion.bsGradientLabel lblArreTar 
      Height          =   360
      Left            =   3855
      Top             =   4620
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   635
      Caption         =   ""
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
      Colour1         =   11513775
      Colour2         =   16558731
      CaptionAlignment=   2
   End
   Begin PCGestion.bsGradientLabel lblComTar 
      Height          =   330
      Left            =   3855
      Top             =   4995
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   582
      Caption         =   ""
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
      Colour1         =   11513775
      Colour2         =   16558731
      CaptionAlignment=   2
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H00FF8080&
      Height          =   2760
      Left            =   3795
      Top             =   2625
      Width           =   3300
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "FINAL"
      Height          =   285
      Left            =   60
      TabIndex        =   41
      Top             =   615
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CAJA"
      Height          =   330
      Left            =   2325
      TabIndex        =   40
      Top             =   105
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H00FF8080&
      Height          =   840
      Left            =   3795
      Top             =   5385
      Width           =   3300
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H00FF8080&
      Height          =   2085
      Left            =   3780
      Top             =   540
      Width           =   3315
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H00FF8080&
      Height          =   1005
      Left            =   15
      Top             =   3990
      Width           =   3780
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H00FF8080&
      Height          =   2925
      Left            =   0
      Top             =   1080
      Width           =   3795
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descuadre"
      Height          =   285
      Left            =   2490
      TabIndex        =   38
      Top             =   6330
      Width           =   1785
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H00FF8080&
      Height          =   840
      Left            =   2250
      Top             =   6300
      Width           =   2145
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      FillColor       =   &H00FF8080&
      Height          =   840
      Left            =   45
      Top             =   6300
      Width           =   2145
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   270
      Left            =   210
      TabIndex        =   37
      Top             =   6360
      Width           =   1860
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H00FF8080&
      Height          =   4005
      Left            =   7140
      Top             =   300
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H00FF8080&
      Height          =   3600
      Left            =   9360
      Top             =   300
      Width           =   2115
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "2 €"
      Height          =   285
      Left            =   7245
      TabIndex        =   34
      Top             =   3810
      Width           =   405
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1 €"
      Height          =   285
      Left            =   7245
      TabIndex        =   33
      Top             =   3375
      Width           =   405
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      Height          =   285
      Left            =   7245
      TabIndex        =   32
      Top             =   2910
      Width           =   405
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      Height          =   285
      Left            =   7200
      TabIndex        =   31
      Top             =   2430
      Width           =   450
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      Height          =   285
      Left            =   7365
      TabIndex        =   30
      Top             =   1995
      Width           =   285
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Height          =   285
      Left            =   7365
      TabIndex        =   29
      Top             =   1500
      Width           =   285
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   285
      Left            =   7365
      TabIndex        =   28
      Top             =   1065
      Width           =   285
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   285
      Left            =   7365
      TabIndex        =   27
      Top             =   600
      Width           =   285
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "500"
      Height          =   285
      Left            =   9465
      TabIndex        =   26
      Top             =   3450
      Width           =   405
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      Height          =   285
      Left            =   9465
      TabIndex        =   25
      Top             =   2970
      Width           =   405
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   285
      Left            =   9420
      TabIndex        =   24
      Top             =   2490
      Width           =   450
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      Height          =   285
      Left            =   9570
      TabIndex        =   23
      Top             =   2010
      Width           =   285
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      Height          =   285
      Left            =   9570
      TabIndex        =   22
      Top             =   1530
      Width           =   285
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      Height          =   285
      Left            =   9570
      TabIndex        =   21
      Top             =   1050
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Height          =   285
      Left            =   9570
      TabIndex        =   20
      Top             =   570
      Width           =   285
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA"
      Height          =   285
      Left            =   60
      TabIndex        =   18
      Top             =   105
      Width           =   690
   End
End
Attribute VB_Name = "frmCierreDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Módulo       : frmCierreDia
' Fecha/Hora : 12/02/2004 11:56
' Autor         : JCastillo
' Propósito   : Cálculo de los datos del cierre
'---------------------------------------------------------------------------------------

Option Explicit

Dim CierreC As Datos_Cierre

Dim T_Billetes As Double 'total importe en billetes
Dim T_Monedas As Double  'total importe en monedas

Dim Descuadre As Double  'importe del descuadre de caja

Dim terminar As Boolean

'para consultar o cerrar la caja
Public consultar As Boolean

'/////-////////////////////////
'/////-////////////////////////
'/////-////////////////////////
Private Sub cbCancelar_Click()
   ' If MsgBox("¿Desea salir del Cierre de Caja?", vbQuestion + vbYesNo, titulo) = vbYes Then
        terminar = False
        Unload Me
   ' End If
End Sub


Private Sub cmCalcular_Click()
Dim rc_tickets As New ADODB.Recordset
Dim tmpcaja As Byte
Dim tffin As String

 On Error GoTo cmCalcular_Click_Error
   
 With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
  End With

  
 If consultar Then
 
    If cbCAJAS.Text <> "" Then
        tmpcaja = cbCAJAS.Text
    'Else
    '    MsgBox "No se permite CAJA en blanco", vbInformation, titulo
    '    cbCAJAS.SetFocus
   '     Exit Sub
    End If
  
 Else
 
    tmpcaja = CajaActual
 
 End If
  
 'distribuye ventas entre las cajas a y b
 If Not consultar Then
    Call Distribuye_Caja_AB(ioFECHA.Text, CajaActual, locCnn)
 End If
    
 If ioFFINAL.Visible And ioFFINAL.Text <> "" Then
    tffin = ioFFINAL.Text
 End If
 
 
 
 'si no es la fecha actual, consultar datos ya guardados en la base de datos
 If (ioFECHA.Text <> Format(Date, "dd/mm/yyyy")) And consultar Then
       
    CierreC = ver_cierre_caja(ioFECHA.Text, tffin, tmpcaja, locCnn)
    'si es la fecha de hoy, calcular el cierre
    
 Else
 
    'calcula los datos del cierre
    CierreC = calcula_cierre_caja(Format(ioFECHA.Text, "yyyymmdd"), tmpcaja, locCnn)
    
 End If
 
 DoEvents
 
 With CierreC
 
''''''''''''''''''' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If .t_caja_Teorico <= 0 Then
        If MsgBox("El total de dinero en caja, parece no ser correcto: " & Format(.t_caja_Teorico, "Currency") & ". Comprueba la fecha del cierre." & Chr(13) & "¿Desea continuar?", vbQuestion + vbYesNo, titulo) = vbNo Then Exit Sub
    End If
''''''''''''''''''' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'presentar resultados:
    lblEfectivo.Caption = "Efectivo: " & Format(.Total_Efectivo, "Currency")
    'lblEfectivo2.Caption = lblEfectivo.Caption
''''''''''''''''''' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    lblVales.Caption = "Vales Acep(" & .n_vales_acep & "): " & Format(.t_vales_acep, "Currency")
    lblValesEmitidos.Caption = "Vales Emitidos(" & .n_vales_emi & "): " & Format(.t_vales_emi, "Currency")
      
    lblValesDctoAceptados.Caption = "Vales Dcto. Ace(" & .n_valdctoa & "): " & Format(.t_valdctoa, "Currency")
    lblValesDctoEmitidos.Caption = "Vales Dcto. Emi(" & .n_valdctoe & "): " & Format(.t_valdctoe, "Currency")
 
''''''''''''''''''' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    lblTarjeta.Caption = "Tarjeta(" & .n_tarjeta & "): " & Format(.t_tarjeta, "Currency")
    lblDevol.Caption = "Devoluciones(" & .n_devol & "): -" & Format(.t_devol, "Currency")
    lblArreglos.Caption = "Arreglos(" & .n_arreglos & "): " & Format(.T_Arreglos, "Currency")
    lblIngresos.Caption = "Ingresos(" & .n_ingresos & "): " & Format(.t_ingresos, "Currency")
    lblPagos.Caption = "Pagos(" & .n_pagos & "): -" & Format(.t_pagos, "Currency")
    lblDeudCli.Caption = "Deud. Clientes(" & .n_deudc & "): -" & Format(.t_deudc, "Currency")
    lblDeudCobradas.Caption = "Deud. Cobradas(" & .n_deudc_pag & "): " & Format(.t_deudc_pag, "Currency")
    lblMovimientosCaja.Caption = "Mov. Caja(" & .n_movi & "): " & Format(.t_movi, "Currency")
    lblEnCaja.Caption = "C. Teórico: " & Format(.t_caja_Teorico, "Currency")
    lblTotalDcto.Caption = "Descuentos(" & .n_dcto & "): -" & Format(.t_dcto, "Currency")
    
    lblDiferenciasPVP.Caption = "Diferencias PVP(" & .n_difcampr & "): " & Format(.t_difcampr, "Currency")

    lblArreCon.Caption = "Arreg. Con.: " & Format(.T_ArreCon, "Currency")
    lblArreTar.Caption = "Arreg. Tar.: " & Format(.T_Arreglos - .T_ArreCon, "Currency")
    
    lblComTar.Caption = "Com.Tarjeta: (" & .n_comtar & "): " & Format(.t_comtar, "Currency")
    
''''''''''''''''''' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    lblCajaA.Caption = "Total en A: " & Format(.Total_A, "Currency")
    lblCajaB.Caption = "Total en B: " & Format(.Total_B, "Currency")
''''''''''''''''''' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    lblTotalCobrosRealizados.Caption = "Cobros Realizados: " & Format((.Cobros_Realizados), "Currency")
    lblVentasNetas.Caption = "Ventas Netas: " & Format((.Ventas_Netas), "Currency")
    lblSaldoCajaEfectivo.Caption = "S. Caja Efectivo: " & Format((.Saldo_Caja_Efectivo), "Currency")
    lblVentasBrutas.Caption = "Ventas Brutas:" & Format((.Ventas_Brutas), "Currency")
      
''''''''''''''''''' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 End With

 'llenar el grid con los tickets de pago por tarjeta de crédito.
 '3 - 2ª
 With fg
    .Clear
    .Rows = 1
    .Cols = 4
    .ColFormat(2) = "Currency"
    .TextMatrix(0, 1) = "Ticket Venta"
    .TextMatrix(0, 2) = "Importe"
    .TextMatrix(0, 3) = "Fecha/Hora"
 End With
 
 With rc_tickets
    
   If consultar Then
    .Open "SELECT CODIGO, IMP_PRIMERA, FMODI FROM CABVENTA WHERE (CODCAJA = " & cbCAJAS.Text & ") AND (FCOBRO IN (2, 6, 13)) and (IMP_PRIMERA > 0) AND (FHORA = '" & Format(ioFECHA.Text, "yyyymmdd") & "') AND (ESTADO = 1)", locCnn, adOpenDynamic, adLockReadOnly
   Else
    .Open "SELECT CODIGO, IMP_PRIMERA, FMODI FROM CABVENTA WHERE (CODCAJA = " & CajaActual & ") AND (FCOBRO IN (2, 6, 13)) and (IMP_PRIMERA > 0) AND (FHORA = '" & Format(ioFECHA.Text, "yyyymmdd") & "') AND (ESTADO = 1)", locCnn, adOpenDynamic, adLockReadOnly
   End If
 
     Do Until .EOF
     
        fg.Rows = fg.Rows + 1
        
        fg.TextMatrix(fg.Rows - 1, 1) = .fields("CODIGO") & Format(CajaActual, "000")
        fg.TextMatrix(fg.Rows - 1, 2) = .fields("IMP_PRIMERA")
        fg.TextMatrix(fg.Rows - 1, 3) = .fields("FMODI")
     
        .MoveNext
     
     Loop
           
    .Close
    
    .Open "SELECT CODIGO, IMP_SEGUNDA, FMODI FROM CABVENTA WHERE (CODCAJA = " & CajaActual & ") AND (FCOBRO IN (3, 9)) and (IMP_SEGUNDA > 0) AND (FHORA = '" & Format(ioFECHA.Text, "yyyymmdd") & "') AND (ESTADO = 1)", locCnn, adOpenDynamic, adLockReadOnly
     
    Do Until .EOF
     
        fg.Rows = fg.Rows + 1
        
        fg.TextMatrix(fg.Rows - 1, 1) = .fields("CODIGO") & Format(CajaActual, "000")
        fg.TextMatrix(fg.Rows - 1, 2) = .fields("IMP_SEGUNDA")
        fg.TextMatrix(fg.Rows - 1, 3) = .fields("FMODI")
        
        .MoveNext
     
    Loop
    
    fg.SubtotalPosition = flexSTAbove
    fg.subtotal flexSTSum, , 2, , vbBlue, vbWhite, True
     
    fg.AutoSize 1, fg.Cols - 1
  
    .Close
 End With
 
 Set rc_tickets = Nothing
 
 'recalcular descuadre ...
 Call calcula_monedas
 Call calcula_billetes
 
 DoEvents
 
 ioM1.SetFocus

   On Error GoTo 0
   Exit Sub

cmCalcular_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmCalcular_Click de Formulario frmCierreDia"

End Sub

Private Sub cmCierreCaja_Click()
Dim tmpmsg As String

'comprobar fecha
   On Error GoTo cmCierreCaja_Click_Error

If ioFECHA.Text = "" Then
    MsgBox "No se puede cerrar caja con fecha en blanco", vbExclamation, titulo
    Exit Sub
End If

'comprobar cierre
If CierreC.Se_Ha_Calculado = False Then
    MsgBox "Debe calcular datos antes de cerrar caja", vbExclamation, titulo
    Exit Sub
End If

'preguntar al usuario
tmpmsg = "¿Desea cerrar DEFINITIVAMENTE la caja del dia?"
If Descuadre <> 0 Then tmpmsg = tmpmsg & Chr(13) & "Hay un descuadre de: " & Format(Descuadre, "Currency")

If MsgBox(tmpmsg, vbQuestion + vbYesNo, titulo) = vbNo Then
    tmpmsg = ""
    Exit Sub
End If

'introducir el dinero real que hay en caja
CierreC.t_caja = T_Billetes + T_Monedas

Call cierra_caja(ioFECHA.Text, CierreC, locCnn)

MsgBox "Se ha CERRADO CAJA correctamente.", vbInformation, titulo

tmpmsg = ""

terminar = True 'para que salga directamente
Unload Me

   On Error GoTo 0
   Exit Sub

cmCierreCaja_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmCierreCaja_Click de Formulario frmCierreDia"

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpimp As String

Select Case KeyCode
Case vbKeyI

    KeyCode = 0

   tmpimp = InputBox("Introducir Nuevo Ticket", "Comprobación Tickets Tarjeta Crédito")
   If tmpimp = "" Then Exit Sub

Case vbKeyB

    KeyCode = 0


End Select

End Sub


Private Sub calcula_monedas()
Dim b_importe As Double


   On Error GoTo calcula_monedas_Error

If ioM1.Text <> "" Then b_importe = b_importe + (ioM1.Text * 0.01)
If ioM2.Text <> "" Then b_importe = b_importe + (ioM2.Text * 0.02)
If ioM5.Text <> "" Then b_importe = b_importe + (ioM5.Text * 0.05)
If ioM10.Text <> "" Then b_importe = b_importe + (ioM10.Text * 0.1)
If ioM20.Text <> "" Then b_importe = b_importe + (ioM20.Text * 0.2)
If ioM50.Text <> "" Then b_importe = b_importe + (ioM50.Text * 0.5)
If ioM100.Text <> "" Then b_importe = b_importe + (ioM100.Text)
If ioM200.Text <> "" Then b_importe = b_importe + (ioM200.Text * 2)

T_Monedas = b_importe
Descuadre = (T_Monedas + T_Billetes) - CierreC.t_caja_Teorico

lblTotalM.Caption = Format(b_importe, "Currency")
lblSumaMB.Caption = Format(T_Monedas + T_Billetes, "Currency")

If Descuadre = 0 Then
    lblDescuadre.Visible = False
    Shape4.Visible = False
    Label18.Visible = False
Else
    
    lblDescuadre.Visible = True
    
    If Descuadre > 0 Then
        Shape4.BorderColor = vbGreen
    Else
        Shape4.BorderColor = vbRed
    End If
    
    Shape4.Visible = True
    
    Label18.Visible = True
    lblDescuadre.Caption = Format(Descuadre, "Currency")
End If

   On Error GoTo 0
   Exit Sub

calcula_monedas_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento calcula_monedas de Formulario frmCierreDia"

End Sub


Private Sub calcula_billetes()
Dim b_importe As Double

   On Error GoTo calcula_billetes_Error

If ioB5.Text <> "" Then b_importe = b_importe + (ioB5.Text * 5)
If ioB10.Text <> "" Then b_importe = b_importe + (ioB10.Text * 10)
If ioB20.Text <> "" Then b_importe = b_importe + (ioB20.Text * 20)
If ioB50.Text <> "" Then b_importe = b_importe + (ioB50.Text * 50)
If ioB100.Text <> "" Then b_importe = b_importe + (ioB100.Text * 100)
If ioB200.Text <> "" Then b_importe = b_importe + (ioB200.Text * 200)
If ioB500.Text <> "" Then b_importe = b_importe + (ioB500.Text * 500)

T_Billetes = b_importe
Descuadre = (T_Monedas + T_Billetes) - CierreC.t_caja_Teorico

lblTotalB.Caption = Format(b_importe, "Currency")
lblSumaMB.Caption = Format(T_Monedas + T_Billetes, "Currency")

If Descuadre = 0 Then
    lblDescuadre.Visible = False
    Shape4.Visible = False
    Label18.Visible = False
Else
    
    lblDescuadre.Visible = True
    
    If Descuadre > 0 Then
        Shape4.BorderColor = vbGreen
    Else
        Shape4.BorderColor = vbRed
    End If
    
    Shape4.Visible = True
    
    Label18.Visible = True
    lblDescuadre.Caption = Format(Descuadre, "Currency")
End If


   On Error GoTo 0
   Exit Sub

calcula_billetes_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento calcula_billetes de Formulario frmCierreDia"

End Sub

Private Sub Form_Load()

Move (Screen.Width - Width) \ 2, Separacion_MDIForm

With ioFECHA
    .dspFormat = "dd/mm/yyyy"
    .LongMaxima = 10
    .Text = Date
End With

If consultar Then
    
    cmCierreCaja.Visible = False
    
    Me.Caption = "Consultar Datos del Cierre ..."
    
    Label19.Visible = True
    cbCAJAS.Visible = True
    
    'Cargar el micombo cajas
    With cbCAJAS
        .ConexionString = locCnn
        .SQLString = "SELECT CODIGO, DESCRIPCION FROM CAJAS WHERE MBAJA = 0 ORDER BY CODIGO"
        .LenCodigo = 1
        .CodigoWidth = 500
        
        .carga
        .Refresh
        .Text = CajaActual
        
         'si es un dependiente, permitir solo la caja actual
         If TipoPermiso = 0 Then
            .Locked = True
         End If
         
    End With
    
    With ioFFINAL
        .dspFormat = "dd/mm/yyyy"
        .LongMaxima = 10
        .Text = Date
        .Visible = True
    End With
    Label20.Visible = True
    
End If

With ioM1
    .SoloNumeros = True
    .Alineacion = 1
    If consultar Then .Locked = True
End With

With ioM2
    .SoloNumeros = True
    .Alineacion = 1
    If consultar Then .Locked = True
End With

With ioM5
    .SoloNumeros = True
    .Alineacion = 1
    If consultar Then .Locked = True
End With

With ioM10
    .SoloNumeros = True
    .Alineacion = 1
    If consultar Then .Locked = True
End With

With ioM20
    .SoloNumeros = True
    .Alineacion = 1
    If consultar Then .Locked = True
End With

With ioM50
    .SoloNumeros = True
    .Alineacion = 1
    If consultar Then .Locked = True
End With

With ioM100
    .SoloNumeros = True
    .Alineacion = 1
    If consultar Then .Locked = True
End With

With ioM200
    .SoloNumeros = True
    .Alineacion = 1
    If consultar Then .Locked = True
End With

With ioB5
    .SoloNumeros = True
    .Alineacion = 1
    If consultar Then .Locked = True
End With

With ioB10
    .SoloNumeros = True
    .Alineacion = 1
    If consultar Then .Locked = True
End With

With ioB20
    .SoloNumeros = True
    .Alineacion = 1
    If consultar Then .Locked = True
End With

With ioB50
    .SoloNumeros = True
    .Alineacion = 1
    If consultar Then .Locked = True
End With

With ioB100
    .SoloNumeros = True
    .Alineacion = 1
    If consultar Then .Locked = True
End With

With ioB200
    .SoloNumeros = True
    .Alineacion = 1
    If consultar Then .Locked = True
End With

With ioB500
    .SoloNumeros = True
    .Alineacion = 1
    If consultar Then .Locked = True
End With

End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If Not terminar Then
    
    If Not consultar Then
    
    If MsgBox("¿Desea salir del Cierre de Caja?", vbQuestion + vbYesNo, titulo) = vbNo Then
        Cancel = True
        Exit Sub
    End If
    
    End If
    
End If

Set frmCierreDia = Nothing

End Sub

Private Sub ioB5_Validate(Cancel As Boolean)
Call calcula_billetes
End Sub

Private Sub ioB10_Validate(Cancel As Boolean)
Call calcula_billetes
End Sub

Private Sub ioB20_Validate(Cancel As Boolean)
Call calcula_billetes
End Sub

Private Sub ioB50_Validate(Cancel As Boolean)
Call calcula_billetes
End Sub

Private Sub ioB100_Validate(Cancel As Boolean)
Call calcula_billetes
End Sub

Private Sub ioB200_Validate(Cancel As Boolean)
Call calcula_billetes
End Sub

Private Sub ioB500_Validate(Cancel As Boolean)
Call calcula_billetes
End Sub





Private Sub ioFECHA_Validate(Cancel As Boolean)

If Not consultar Then

    If ioFECHA.Text <> "" And ioFFINAL.Visible = False Then
        Call cmCalcular_Click
    Else
        MsgBox "No se permite cerrar caja con FECHA en blanco", vbInformation, titulo
        Cancel = True
        ioFECHA.CancelarValidacion
    End If

End If

End Sub



Private Sub ioFFINAL_Validate(Cancel As Boolean)


If Not consultar Then

    If ioFECHA.Text <> "" Then
        Call cmCalcular_Click
    Else
        MsgBox "No se permite cerrar caja con FECHA en blanco", vbInformation, titulo
        Cancel = True
        ioFECHA.CancelarValidacion
    End If

End If

End Sub

Private Sub ioM1_Validate(Cancel As Boolean)
Call calcula_monedas
End Sub

Private Sub ioM2_Validate(Cancel As Boolean)
Call calcula_monedas
End Sub

Private Sub ioM5_Validate(Cancel As Boolean)
Call calcula_monedas
End Sub

Private Sub ioM10_Validate(Cancel As Boolean)
Call calcula_monedas
End Sub

Private Sub ioM20_Validate(Cancel As Boolean)
Call calcula_monedas
End Sub

Private Sub ioM50_Validate(Cancel As Boolean)
Call calcula_monedas
End Sub

Private Sub ioM100_Validate(Cancel As Boolean)
Call calcula_monedas
End Sub

Private Sub ioM200_Validate(Cancel As Boolean)
Call calcula_monedas
End Sub
