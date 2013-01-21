VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmCabVen 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas"
   ClientHeight    =   7200
   ClientLeft      =   1305
   ClientTop       =   2295
   ClientWidth     =   11535
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
   Moveable        =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   11535
   StartUpPosition =   2  'CenterScreen
   Begin PCGestion.bsGradientLabel lblTotal 
      Height          =   840
      Left            =   5595
      Top             =   6330
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   1482
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   16558731
      Colour2         =   7177785
      CaptionAlignment=   2
   End
   Begin PCGestion.chameleonButton cmdPrevious 
      Height          =   630
      Left            =   1065
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5700
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
      MICON           =   "frmCabVen.frx":0000
      PICN            =   "frmCabVen.frx":001C
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
      Left            =   45
      Top             =   5310
      Width           =   11475
      _ExtentX        =   20241
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
      Left            =   30
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5700
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
      MICON           =   "frmCabVen.frx":0CEE
      PICN            =   "frmCabVen.frx":0D0A
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
      Left            =   9390
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5700
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
      MICON           =   "frmCabVen.frx":1A40
      PICN            =   "frmCabVen.frx":1A5C
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
      Left            =   10470
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5700
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
      MICON           =   "frmCabVen.frx":272E
      PICN            =   "frmCabVen.frx":274A
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6360
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
      MICON           =   "frmCabVen.frx":3480
      PICN            =   "frmCabVen.frx":349C
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
      Left            =   1110
      TabIndex        =   0
      Top             =   6360
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
      MICON           =   "frmCabVen.frx":4176
      PICN            =   "frmCabVen.frx":4192
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
      Left            =   9270
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6360
      Width           =   1170
      _ExtentX        =   2064
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
      MICON           =   "frmCabVen.frx":4A6C
      PICN            =   "frmCabVen.frx":4A88
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
      Left            =   10470
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6360
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
      MICON           =   "frmCabVen.frx":5362
      PICN            =   "frmCabVen.frx":537E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel2 
      Height          =   315
      Left            =   2160
      Top             =   5685
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   556
      Caption         =   "- A - Añadir Venta   - E - Editar Venta"
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
      Colour1         =   12632256
      Colour2         =   16558731
   End
   Begin PCGestion.bsGradientLabel lblDependiente 
      Height          =   375
      Left            =   990
      Top             =   465
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   609
      Caption         =   ""
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
      Colour1         =   12632256
      Colour2         =   16558731
      CaptionAlignment=   1
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel4 
      Height          =   315
      Left            =   5775
      Top             =   5685
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   556
      Caption         =   "-C- Asignar Cliente  -N- Nuevo Cliente"
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
      Colour1         =   12632256
      Colour2         =   16558731
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel1 
      Height          =   315
      Left            =   2160
      Top             =   6000
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   556
      Caption         =   "- D - Asig. Dependi   -F11-  Devoluciones"
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
      Colour1         =   12632256
      Colour2         =   16558731
   End
   Begin PCGestion.bsGradientLabel lblCliente 
      Height          =   375
      Left            =   6825
      Top             =   465
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   661
      Caption         =   ""
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
      Colour1         =   12632256
      Colour2         =   16558731
      CaptionAlignment=   1
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel3 
      Height          =   315
      Left            =   5790
      Top             =   6000
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   556
      Caption         =   "-I- Ir a Rejilla         -B- Borrar Seleccion"
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
      Colour1         =   12632256
      Colour2         =   16558731
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel5 
      Height          =   315
      Left            =   2340
      Top             =   6375
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      Caption         =   "- R - Arreglos"
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
      Colour1         =   12632256
      Colour2         =   16558731
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   4425
      Left            =   30
      TabIndex        =   15
      Top             =   885
      Width           =   11505
      _cx             =   20294
      _cy             =   7805
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16626604
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
      FormatString    =   $"frmCabVen.frx":6058
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
   Begin PCGestion.bsGradientLabel bsGradientLabel6 
      Height          =   315
      Left            =   2325
      Top             =   6735
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   556
      Caption         =   "- T - TERMINAR"
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   16761024
      Colour2         =   14737632
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel7 
      Height          =   315
      Left            =   4020
      Top             =   6375
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   556
      Caption         =   "- S - Descuento"
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
      Colour1         =   12632256
      Colour2         =   16558731
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "DEPEND."
      Height          =   300
      Left            =   105
      TabIndex        =   14
      Top             =   502
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FCAA8B&
      Caption         =   "Total"
      Height          =   300
      Left            =   4980
      TabIndex        =   13
      Top             =   6870
      Width           =   570
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENTE"
      Height          =   300
      Left            =   5865
      TabIndex        =   12
      Top             =   510
      Width           =   855
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
      Left            =   9000
      TabIndex        =   5
      Top             =   45
      Width           =   2490
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
      Left            =   990
      TabIndex        =   4
      Top             =   45
      Width           =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      Height          =   330
      Left            =   90
      TabIndex        =   3
      Top             =   60
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ultima modificación"
      Height          =   315
      Left            =   6780
      TabIndex        =   2
      Top             =   75
      Width           =   2160
   End
End
Attribute VB_Name = "frmCabVen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'---------------------------------------------------------------------------------------
' Modulo      : frmCabVen
' Fecha/Hora  : 17/01/2004 22:12
' Autor       : JCASTILLO
' Propósito   : Cabecera de ventas
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

Dim rc_ventas As New ADODB.Recordset
Dim rc_arreglos As New ADODB.Recordset
Dim rccli As New ADODB.Recordset

Dim oSQL As New clsSmartSQL
Dim dSQL As New clsSmartSQL

Public Entrar_Agregando As Boolean

'---------------------------------------------------------------------------------------
'si se llama desde el grid de ventas, simplemente presentar la venta correspondiente
'y no permitir Agregar nuevas Ventas
Public Desde_Grid As Boolean
Public D_Codigo As Long
Public D_CodCaja As Byte
'---------------------------------------------------------------------------------------


Dim conta_lineas As Long

'Public añade_nueva As Boolean

'---------------------------------------------------------------------------------------
' Procedimiento : des_enlaza_campos
' Fecha/Hora    : 11/12/2003 11:12
' Autor         : JCastillo
' Propósito     :  Desenlaza los controles para cerrar la conexión
'---------------------------------------------------------------------------------------
Private Sub des_enlaza_campos()

'des-enlazar controles
With ioCODIGO
  Set .DataSource = Nothing
        .DataField = ""
End With
      
'With cbCODCLI
'  Set .DataSource = Nothing
'        .DataField = ""
'End With
  
With ioFMODI
  Set .DataSource = Nothing
        .DataField = ""
End With

End Sub


'---------------------------------------------------------------------------------------
' Procedimiento : enlaza_campos
' Fecha/Hora    : 11/12/2003 11:13
' Autor         : JCastillo
' Propósito     :   Vuelve a enlazar los campos al recordset
'---------------------------------------------------------------------------------------
Private Sub enlaza_campos()

With ioCODIGO
  Set .DataSource = rc
        .DataField = "CODIGO"
End With
      
'With cbCODCLI
'    .DataField = "CODALMDEST"
'    Set .DataSource = rc
'End With
  
On Error Resume Next
With ioFMODI
  Set .DataSource = rc
        .DataField = "FMODI"
End With

On Error GoTo 0

End Sub



'---------------------------------------------------------------------------------------
' Subrutina   : fg_DblClick
' Fecha/Hora  : 08/02/2004 18:41
' Autor       : JCASTILLO
' Propósito   : Ir a editar artículos o el arreglo seleccionado
'---------------------------------------------------------------------------------------
Private Sub fg_dblClick()
Dim T_Arreglo As Variant
Dim tmpcodcost As Long
Dim tmpnomcost As String

Dim codop As Byte
Dim tmpid As Long

   On Error GoTo fg_dblClick_Error

If rc.RecordCount <= 0 Then Exit Sub


'ver tipo 1=articulo, 2=arreglo
Select Case fg.TextMatrix(fg.Row, 1)

'si esta en blanco no hacer nada
Case ""

  Exit Sub

'si es un articulo ...
Case "1"

    tmpid = fg.TextMatrix(fg.Row, 2)
    
    'si entro desde el grid, avisar que la venta ya se ha introducido en almacén, preguntar
    'si se desea deshacer la venta
    If Desde_Grid Then
        
            If rc.fields("ESTADO") = 1 Then
                If MsgBox("La venta actual ya ha sido aceptada, si desea modificarla, primero se debe pasar la venta a PENDIENTE. ¿Desea continuar?", vbQuestion + vbYesNo, titulo) = vbNo Then Exit Sub
                Call Deshacer_Venta(rc.fields("CODIGO"), rc.fields("CODCAJA"), locCnn)
                DoEvents
                rc.Requery
            End If
        
    End If
    
    DoEvents
    
    With frmDetVenta
        .CODIGO_VENTA = rc.fields("CODIGO")
        .CODIGO_USR = rc.fields("CODPER")
        .IR_A_ID = tmpid ' fg.TextMatrix(fg.Row, 2)  'ir al id marcado
        ' .IR_A_ID = 1
        Set .rc = rc_ventas
        Me.WindowState = vbMinimized
        .Show 1
        Me.WindowState = vbNormal
        Call carga_grid_ventas
    End With

'si es un arreglo, presentarlo para cambiar datos
Case "2"


tmpid = fg.TextMatrix(fg.Row, 2)

'obtener los datos que no tenemos en el grid ...
T_Arreglo = devuelve_matriz("SELECT CODCOST, DESCRIPCION, COSTE FROM ARREGLOS WHERE ID = " & fg.TextMatrix(fg.Row, 2) & " AND CODCAJ = " & CajaActual, locCnn)

'cargar el nombre de la costurera, si ya ha sido asignado
If Not IsNull(T_Arreglo(0)) Then
    tmpcodcost = T_Arreglo(0)
    tmpnomcost = devuelve_campo("SELECT NOMBRE FROM COSTURE WHERE CODIGO = " & tmpcodcost, locCnn)
End If

If Not IsArray(T_Arreglo) Then
    lblstatus.Caption = "Error al editar el arreglo"
    Exit Sub
End If

With frmNuArr

    .Solo_Actualizar = True
    .Sel_Caja = CajaActual
    .Sel_ID = tmpid 'fg.TextMatrix(fg.Row, 2)
    
    If tmpnomcost <> "@" And Trim(tmpnomcost) <> "" Then
        .ioNOMBRE.Text = Trim(tmpnomcost)
    End If
    
    .ioCOSTO.Text = T_Arreglo(2)
    .ioDESCRIPCION.Text = Trim(T_Arreglo(1))
    .ioPVP.Text = fg.TextMatrix(fg.Row, 8)
   
   ' Select Case fg.TextMatrix(fg.Row, 11)
            
    '        Case "PENDIENTE"
            
    .cbESTADO.Text = 1
    '        Case "SERVIDO"
            
    '                .cbESTADO.Text = 2
    '        Case "CANCELADO"
            
    '                .cbESTADO.Text = 3
                      
    'End Select
            
    '.cbESTADO.Text = fg.TextMatrix(fg.Row, 6)
    
    .Show 1

End With

Call carga_grid_ventas

End Select

   On Error GoTo 0
   Exit Sub

fg_dblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento fg_DblClick de Formulario frmCabVen"

End Sub


'---------------------------------------------------------------------------------------
' Procedimiento : Borra_Apunte_Seleccionado
' Fecha/Hora    : 21/01/2004 10:18
' Autor         : JCastillo
' Propósito     :  Borra el apunte seleccionado en la base de datos (artículo
' o arreglo
'---------------------------------------------------------------------------------------
Private Sub Borra_Apunte_Seleccionado()

   On Error GoTo Borra_Apunte_Seleccionado_Error

        
  If rc.RecordCount <= 0 Then Exit Sub

'ver tipo 1=articulo, 2=arreglo
Select Case fg.TextMatrix(fg.Row, 1)

'si esta en blanco no hacer nada
Case ""

'si es un articulo ...
Case "1"

'borrar el articulo

 If MsgBox("Desea quitar el ARTÍCULO seleccionado: " & fg.TextMatrix(fg.Row, 3), vbQuestion + vbYesNo, titulo) = vbNo Then Exit Sub
 
 locCnn.Execute "DELETE FROM DETVENTA WHERE CODVEN = " & rc.fields("CODIGO") & " AND CODCAJA = " & CajaActual & " AND LINEA = " & fg.TextMatrix(fg.Row, 2)
 DoEvents
 Call carga_grid_ventas

'si es un arreglo
Case "2"

 If MsgBox("Desea quitar el ARREGLO seleccionado " & fg.TextMatrix(fg.Row, 3), vbQuestion + vbYesNo, titulo) = vbNo Then Exit Sub
 
 'quitarlo de la venta actual
 locCnn.Execute "UPDATE ARREGLOS SET CODVEN = 0 WHERE CODVEN = " & rc.fields("CODIGO") & " AND CODCAJ = " & CajaActual & " AND ID = " & fg.TextMatrix(fg.Row, 2)
 DoEvents
 Call carga_grid_ventas


End Select
    
    DoEvents
    
    Call Form_KeyDown(vbKeyI, 0)

   On Error GoTo 0
   Exit Sub

Borra_Apunte_Seleccionado_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Borra_Apunte_Seleccionado de Formulario frmCabVen"
End Sub


Private Sub fg_KeyPress(KeyAscii As Integer)

'si se pulsa enter ...
If KeyAscii = 13 Then
    Call fg_dblClick
    KeyAscii = 0
End If

End Sub

Private Sub Form_Activate()

   With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
   End With
   
End Sub

Private Sub Form_GotFocus()

   With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
   End With

End Sub


Private Sub ioCODIGO_Change()
ioCODIGO.Caption = Format(ioCODIGO.Caption, "00000")
End Sub

Private Sub Form_Load()

   On Error GoTo Form_Load_Error

  'Move (Screen.Width - Width) \ 2, Separacion_MDIForm
  
   DoEvents
 
   frmVerFecha.Show 1
 
  
   With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
   End With
   
  Set rc = New Recordset
  oSQL.AddTable "CABVENTA"
  
  'si entra normal
  If Not Desde_Grid Then
    oSQL.AddSimpleWhereClause "ESTADO", 0
    oSQL.AddSimpleWhereClause "CODCAJA", CajaActual
    'si entra desde el grid, seleccionar la venta
  Else
    oSQL.AddSimpleWhereClause "CODIGO", D_Codigo
    oSQL.AddSimpleWhereClause "CODCAJA", D_CodCaja
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

'With cbCODCLI
'    .ConexionString = locCnn
'    .LenCodigo = 8
'    .SQLString = "SELECT CODIGO, RAZO FROM CLIENTES WHERE MBAJA = 0 ORDER BY CODIGO"
'    .CodigoWidth = 900
'    .DataField = "CODCLI"
'    .carga
'    Set .DataSource = rc
'End With

With fg
    .Cols = 11
    .Rows = 1
    .Editable = flexEDNone
End With
  
With ioFMODI
  Set .DataSource = rc
        .DataField = "FMODI"
End With

If Not prime Then

 ' Call carga_grid_ventas
  
  'If rc.RecordCount <= 0 Then
        
        'If MsgBox("No se encuentran Transferencias. ¿Crear?", vbYesNo + vbQuestion, "Transferencias") = vbNo Then
        'Unload Me
        'Else
        'Call cbAgregar_Click
        'End If
        
        'inicializa los procesos para usar ofertas
        Call Inicializa_Ofertas(locCnn)
        
        If OfertaActual > 0 Then Me.Caption = Me.Caption & "    " & OfertaDSC
                
        If Entrar_Agregando Then
            Call cbAgregar_Click
        
        Else
            Call cmdFirst_Click
        End If
  'Else
        'If añade_nueva Then
        'Call cbAgregar_Click
        'Else
   '     Call cmdFirst_Click
   '     Call cbCancelar_Click
        'End If
        
 'End If

prime = True
End If

mbDataChanged = False



   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Form_Load de Formulario frmCabVen"
      
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpdcto As String
  
   'On Error GoTo Form_KeyDown_Error

  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      
      If MsgBox("¿Desea cerrar la ventana actual?", vbQuestion + vbYesNo, "Cerrar") = vbYes Then
        cbcerrar_Click
      End If
       KeyCode = 0
 '
 '   Case vbKeyEnd
 '     cmdLast_Click
 '      KeyCode = 0
      
 '   Case vbKeyHome
 '     cmdFirst_Click
 '      KeyCode = 0
      
 '   Case vbKeyPageUp
 '     If Shift = vbCtrlMask Then
'        cmdFirst_Click
 '     Else
  '      cmdPrevious_Click
   '   End If
    '   KeyCode = 0
      
 '   Case vbKeyPageDown
 '     If Shift = vbCtrlMask Then
 '       cmdLast_Click
 '     Else
 '       cmdNext_Click
 '     End If
 '      KeyCode = 0
      
      Case vbKeyF1
            Call cbAgregar_Click
             KeyCode = 0
        
      Case vbKeyF2
            Call cbactualizar_Click
             KeyCode = 0
        
      Case vbKeyF4
       
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
         
       'para q por error no vaya a la ventana de cerrar formulario
       Case vbKeyF10
        KeyCode = 0
        
        
       'añadir artículos
       Case vbKeyA
       
        'si entro desde el grid, avisar que la venta ya se ha introducido en almacén, preguntar
        'si se desea deshacer la venta
        If Desde_Grid Then
        
            If rc.fields("ESTADO") = 1 Then
                If MsgBox("La venta actual ya ha sido aceptada, si desea modificarla, primero se debe pasar la venta a PENDIENTE. ¿Desea continuar?", vbQuestion + vbYesNo, titulo) = vbNo Then Exit Sub
                Call Deshacer_Venta(rc.fields("CODIGO"), rc.fields("CODCAJA"), locCnn)
                DoEvents
                rc.Requery
            End If
        
        End If
        
        If rc.RecordCount <= 0 Then Exit Sub
        With frmDetVenta
        
            .CODIGO_VENTA = rc.fields("CODIGO")
            .CODIGO_USR = rc.fields("CODPER")
            .IR_A_ID = 0
            Set .rc = rc_ventas
            
            Me.WindowState = vbMinimized
            
            .Show 1
            
            Me.WindowState = vbNormal
            
            DoEvents
'            Me.SetFocus
            
            Call carga_grid_ventas
            
            DoEvents
            
        End With
         KeyCode = 0
        
       'editar artículos
       Case vbKeyE
       
        If rc.RecordCount <= 0 Then Exit Sub
        Call fg_dblClick
         KeyCode = 0
       
       'IR A GRID
      Case vbKeyI
            
        If rc.RecordCount <= 0 Then Exit Sub
        fg.TabStop = True
        fg.SetFocus
         KeyCode = 0
        
      'Asignar Cliente ...
      Case vbKeyC
      
       'abre el grid de los clientes
       Call Abre_Grid_Clientes
        KeyCode = 0
      
      
      'crear nuevo cliente rapido
      Case vbKeyN
        
        If rc.RecordCount <= 0 Then Exit Sub
        With frmNuCliRap
        
            .Show 1
            
            DoEvents
            Me.SetFocus
            
            If .ID_Cliente_Creado > 0 Then
            rc.fields("CODCLI") = .ID_Cliente_Creado
            rc.fields("CAJACLI") = .Caja_Cliente
            rc.Update
            lblCliente.Caption = .RAZO_Creado
            End If
            
        
        End With
        
        Set frmNuCliRap = Nothing
         KeyCode = 0
        
        
      'seleccionar dependiente
      Case vbKeyD
        
        Call selecciona_Dependiente
        KeyCode = 0
        
        
     'borrar apunte seleccionado
      Case vbKeyB
        
        Call Borra_Apunte_Seleccionado
        KeyCode = 0
         
        
      'Arreglos, mostrar el dialogo de seleccionar una opcion
      Case vbKeyR
        
        If rc.RecordCount <= 0 Then Exit Sub
        
        frmOptArreglo.Show 1
        
        DoEvents
        Me.SetFocus
       
            Select Case frmOptArreglo.Opcion
            
            Case 0 ' cancelado
            
                Exit Sub
            
            Case 1 ' nuevo
            
                With frmNuArr

                    If (fg.TextMatrix(fg.Row, 2) <> "" And IsNumeric(fg.TextMatrix(fg.Row, 2) <> "")) Then
                        'buscar la linea correspondiente
                        rc_ventas.Find "LINEA = " & fg.TextMatrix(fg.Row, 2), , adSearchForward, 1
        
                        If Not rc_ventas.EOF Then
                            .mi_Codart = rc_ventas.fields("CODART")
                            .mi_Tempor = rc_ventas.fields("TEMPOR")
                            .mi_Color = rc_ventas.fields("CODCOL")
                            .mi_talla = rc_ventas.fields("CODTALLA")
                        End If
                        
                        .lblArticulo.Caption = fg.TextMatrix(fg.Row, 3) & " " & fg.TextMatrix(fg.Row, 4) & " " & fg.TextMatrix(fg.Row, 5)
                                    
                    Else
                    
                    .ioCODBAR.Visible = True
                    
                    End If
                               
                    .CODIGO_VENTA = rc.fields("CODIGO")
                    
                    
                    .Show 1
                    
                    DoEvents
                    
                    'cargar las ventas
                    Call carga_grid_ventas
        
                End With
                        
            Case 2 'seleccionar
            
                With frmFlexArre
                    .Venta_Actual = rc.fields("CODIGO")
                    .desde_ventas = True
                    .Show 1
                End With
                                
                Call carga_grid_ventas
                
            End Select
         
         Set frmOptArreglo = Nothing
         KeyCode = 0
       
      'T, terminar venta, y pasar a una nueva venta
      Case vbKeyT
      
      If fg.Rows <= 1 Then Exit Sub
            
      If fg.TextMatrix(1, 11) <> "" Then
      
        With frmFcobro
        
            'pasarle el importe
            .ID_Dependiente = rc.fields("CODPER")
            .Importe = fg.TextMatrix(1, 11)
            
            If IsNull(rc.fields("CODCLI")) Then
            .ID_Cliente = 0
            Else
            .ID_Cliente = rc.fields("CODCLI")
            End If
            
            If IsNull(rc.fields("CAJACLI")) Then
            .Caja_Cliente = 0
            Else
            .Caja_Cliente = rc.fields("CAJACLI")
            End If
      
            .Show 1
            DoEvents
                    
            If .Cancelado = False Then
                'actualizar los datos ...
        
                rc.fields("FCOBRO") = .F_Cobro
                rc.fields("IMP_PRIMERA") = .I_Primera
                rc.fields("IMP_SEGUNDA") = .I_Segunda
                rc.fields("COMISION") = .I_Comision
                rc.Update
                
                'aceptar definitivamente la venta
                Call Aceptar_Venta
                
                DoEvents
                
                If Not Desde_Grid Then Call cbAgregar_Click
                
            End If
                
            Set frmFcobro = Nothing
                
        End With
      
        
      
      End If
            
      KeyCode = 0
      
      'vales
      ' Case vbKeyV
      '
       ' Me.Hide
       ' DoEvents
       ' frmMntVal.Show 1
       ' Me.Show
      
      'para que me deje moverme
      Case vbKeyUp, vbKeyDown, 13
      
      Case vbKeyF11
                        
            Set rc.ActiveConnection = Nothing
            Me.WindowState = vbMinimized
            
            With frmMntDev
                .desde_ventas = True
                .Show 1
            End With
            
            Set rc.ActiveConnection = locCnn
            Me.WindowState = vbNormal
                                    
            DoEvents
            
      Exit Sub
      
      'introducir descuento
      Case vbKeyS
        
            If rc_ventas.RecordCount <= 0 Then Exit Sub
            tmpdcto = InputBox("Introduzca un descuento para todos los artículos de la venta", "Introducir descuento", "0")
        
            If Trim(tmpdcto) = "" Or Not IsNumeric(Trim(tmpdcto)) Then Exit Sub
        
            If CLng(tmpdcto) > 100 Then
                lblstatus.Caption = "Descuento mayor del 100% no válido"
                KeyCode = 0
                Exit Sub
            End If
        
            locCnn.Execute "UPDATE DETVENTA SET DCTO = " & tmpdcto & " WHERE (CODVEN = " & rc.fields("CODIGO") & ") AND (CODCAJA = " & CajaActual & ")"
            DoEvents
            Call carga_grid_ventas
        
            KeyCode = 0
            lblstatus.Caption = ""
        
  End Select
  
   On Error GoTo 0
   Exit Sub

Form_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Form_KeyDown de Formulario frmCabVen"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error Resume Next

If rc.EditMode <> adEditNone Then rc.CancelUpdate

If rc_ventas.State = 1 Then rc_ventas.Close
Set rc_ventas = Nothing

If rccli.State = 1 Then rccli.Close
Set rccli = Nothing

If rc.State = 1 Then rc.Close
Set rc = Nothing

   With locCnn
   ' If .State <> 0 Then .Close
   End With
   
Set oSQL = Nothing

'If Me.MDIChild = True Then frmMenuTactil.Show
Set frmCabVen = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub



'=======================================================================================================================================================================
Private Sub rc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  On Error GoTo rc_MoveComplete_Error
  
  '''''''''''''''''''''''''''''''''''''''''''''' ''''  '''' '''''''''
  'Esto mostrará la posición de registro actual (para) este RecorDseT
  '''''''''''''''''''''''''''''''''''''''''''''' ''''  '''' '''''''''
  
  If rc.AbsolutePosition > 0 Then
  
    Call carga_grid_ventas
    
    If rc.fields("CODPER") > 0 Then
        lblDependiente.Caption = devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & rc.fields("CODPER"), locCnn)
    Else
        lblDependiente.Caption = ""
    End If
    
    If ((rc.fields("CODCLI") > 0 And Not IsNull(rc.fields("CODCLI"))) And (rc.fields("CAJACLI") > 0) And Not IsNull(rc.fields("CAJACLI"))) Then
        lblCliente.Caption = devuelve_campo("SELECT RAZO FROM CLIENTES WHERE CODIGO = " & rc.fields("CODCLI") & " AND CODCAJA = " & rc.fields("CAJACLI"), locCnn)
    Else
        lblCliente.Caption = ""
    End If
  
  Else
    lblDependiente.Caption = ""
    lblCliente.Caption = ""
  End If
  

   On Error GoTo 0
   Exit Sub

rc_MoveComplete_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento rc_MoveComplete de Formulario frmPtrans"
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
' Subrutina   : selecciona_Dependiente
' Fecha/Hora  : 20/01/2004 21:21
' Autor       : JCASTILLO
' Propósito   :
'---------------------------------------------------------------------------------------
Private Sub selecciona_Dependiente()
   On Error GoTo selecciona_Dependiente_Error

  'seleccionar dependiente
  With frmSelDep
  .Show 1
    
  lblDependiente.Caption = .Nombre_Dep
  

  
  End With

   On Error GoTo 0
   Exit Sub

selecciona_Dependiente_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento selecciona_Dependiente de Formulario frmCabVen"
End Sub

Private Sub cbAgregar_Click()
  Dim tmpcodigo As Variant
  
   On Error GoTo cbAgregar_Click_Error


 ' cbCODPER.Enabled = False
  
  With rc
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
   
    Call selecciona_Dependiente
    
    If frmSelDep.S_Cancelado Then
        Set frmSelDep = Nothing
        Exit Sub
    End If
    
    
    .AddNew
    
    tmpcodigo = devuelve_campo("select max(CODIGO) + 1 from CABVENTA where CODCAJA = " & CajaActual)
    
    'Si devuelve @ esque ha habido un error
    If tmpcodigo = "@" Then tmpcodigo = 1
    
    .fields("CODPER") = frmSelDep.ID_Dependiente
    .fields("CODIGO") = tmpcodigo
    .fields("CODCAJA") = CajaActual
        
    'End If

    Set frmSelDep = Nothing

    Set tmpcodigo = Nothing
    
    lblstatus.Caption = "Agregar registro"
    mbAddNewFlag = True
    SetButtons False
    cbActualizar.Visible = True
    Set frmSelDep = Nothing
    
    'guardar el registro ...
    Call cbactualizar_Click
    
    DoEvents
    'añadir ventas ...
    
    Call Form_KeyDown(vbKeyA, 0)
    'SendKeys "{A}"
      
'ioDescripcion.SetFocus
  End With

DoEvents

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'no funciona cbcodalorig.setfocus ¿?
'cbCODCLI.SetFocus
'SendKeys "{UP}"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


   On Error GoTo 0
   Exit Sub

cbAgregar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbAgregar_Click de Formulario frmPtrans"

End Sub

''Private Sub cbeliminar_Click()
'    On Error GoTo DeleteErr
' With rc
'    '.Delete
'    '.MoveNext
'    .Fields("mbaja") = True
'   .Fields("FBAJA") = Date
'   If .EOF Then .MoveLast
' End With
'
' Call cbactualizar_Click
'
'
''Exit Sub
'DeleteErr:
'  MsgBox Err.Description
'End Sub

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
  
   On Error GoTo cbactualizar_Click_Error

  'validaciones
  
  'rc.Fields("CODPER") = UsuarioActual
  rc.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    rc.Requery
    Call cmdLast_Click
    'rc.MoveLast              'va al nuevo registro
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  
  lblstatus.Caption = ""
  
  
  
  'If ioULTIMAS.Text <> "" Then carga_grid (CLng(ioULTIMAS.Text))
   
  DoEvents

  
  'Call cbAgregar_Click

   On Error GoTo 0
   Exit Sub

cbactualizar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbactualizar_Click de Formulario frmPtrans"

End Sub

Private Sub cbcerrar_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  If Not rc.BOF Then rc.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
If Err.Number = -2147217842 Then Exit Sub
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  If Not rc.EOF Then rc.MoveLast
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
  cbActualizar.Visible = Not bVal
  cbCerrar.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
'  cbLista.Visible = bVal
'  dtAgregar.Visible = bVal
    
'  cbCODPER.Locked = bVal
'  cbCODCLI.Locked = bVal
    
End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : carga_grid_ventas
' Fecha/Hora  : 17/01/2004 20:47
' Autor       : JCASTILLO
' Propósito   : Carga la parte del grid correspondiente a los artículos que forman la
'               venta
'---------------------------------------------------------------------------------------
Private Sub carga_grid_ventas()
Dim tmpcodcolor As Long
Dim linea As Long
Dim miArt As Variant
Dim calcdcto As Single
Dim var As Long

'variables para guardar los totales
Dim T_Unidades As Double
Dim T_SubTotal As Double
Dim T_Total As Double
Dim t_iva As Double
Dim T_RE As Double
Dim T_Arreglos As Double

   On Error GoTo carga_grid_ventas_Error
 
If rc.fields("CODIGO") = 0 Then Exit Sub
'abrir articulos de la venta actual
If rc_ventas.State = 1 Then rc_ventas.Close
rc_ventas.Open "SELECT * FROM DETVENTA WHERE CODVEN = " & rc.fields("CODIGO") & " AND CODCAJA = " & CajaActual & " ORDER BY LINEA", locCnn, adOpenDynamic, adLockOptimistic

'Poner titulos antes de nada
With fg
        .Clear
        .Rows = 1
        .Cols = 12
        .ColHidden(1) = True
        .ColHidden(2) = True
        
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(9) = flexAlignRightCenter
        .ColAlignment(10) = flexAlignRightCenter
        
        .TextMatrix(0, 3) = "MODELO"
        .TextMatrix(0, 4) = "TALLA"
        .TextMatrix(0, 5) = "COLOR"
        .TextMatrix(0, 6) = "UNIDADES"
        .TextMatrix(0, 7) = "PVP"
        .TextMatrix(0, 8) = "SUBT."
        .TextMatrix(0, 9) = "DCTO"
        .TextMatrix(0, 10) = "IVA"
        .TextMatrix(0, 11) = "TOTAL"
End With

conta_lineas = 0

'si no hay articulos, salir
If rc_ventas.RecordCount > 0 Then
    
    lblTotal.Caption = ""
        'Exit Sub

    'CODIGO DE OPERACION:
    ' -> 1  la linea hace referencia a un articulo
    ' -> 2  la linea hace referencia a un arreglo
    

    With fg
        
        .Redraw = flexRDNone
        
        Do
       
        conta_lineas = conta_lineas + 1
        .Rows = .Rows + 1
        
        .TextMatrix(.Rows - 1, 1) = 1  'codigo de operacion: ARTICULO
        .TextMatrix(.Rows - 1, 2) = rc_ventas.fields("LINEA")
        .TextMatrix(.Rows - 1, 3) = Format(rc_ventas.fields("CODART"), "00000") & " " & devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & rc_ventas.fields("CODART") & " AND TEMPOR = " & rc_ventas.fields("TEMPOR"))
        .TextMatrix(.Rows - 1, 4) = Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & rc_ventas.fields("CODTALLA")))
        
       'obtener el texto del color y su codigo de color (para colorear
       'la celda del grid)
       If rc_ventas.fields("CODCOL") > 0 Then
      
            tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & rc_ventas.fields("CODCOL"))
            .TextMatrix(.Rows - 1, 5) = devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & rc_ventas.fields("CODCOL"))
            .Col = 5
            .Row = .Rows - 1
            .CellBackColor = tmpcodcolor
            .Col = 2
        
       End If
                
        'Unidades
        T_Unidades = T_Unidades + rc_ventas.fields("UNIDADES")
        .TextMatrix(.Rows - 1, 6) = rc_ventas.fields("UNIDADES")
        'PVP
        .TextMatrix(.Rows - 1, 7) = rc_ventas.fields("PREVEN")
        
        'importe (subtotal)
        
        .TextMatrix(.Rows - 1, 8) = rc_ventas.fields("UNIDADES") * rc_ventas.fields("PREVEN")
        T_SubTotal = T_SubTotal + .TextMatrix(.Rows - 1, 8)
        
        'DCTO
        .TextMatrix(.Rows - 1, 9) = rc_ventas.fields("DCTO") & " %"
        
        'iva
        .TextMatrix(.Rows - 1, 10) = rc_ventas.fields("IVA") & " %"
        t_iva = t_iva + rc_ventas.fields("IVA")
        
        'RE
        T_RE = T_RE + rc_ventas.fields("RE")
        
        'calcular importe descuento
        calcdcto = (.TextMatrix(.Rows - 1, 8) * rc_ventas.fields("DCTO")) / 100
        
        'total  = ((importe-dcto) + IVA)
        .TextMatrix(.Rows - 1, 11) = .TextMatrix(.Rows - 1, 8) - calcdcto + ((.TextMatrix(.Rows - 1, 8) * rc_ventas.fields("IVA")) / 100)
        T_Total = T_Total + .TextMatrix(.Rows - 1, 11)
        
        
        '  .TextMatrix(.Rows - 1, 5) = devuelve_campo("SELECT AÑO + ' ' + TEMPORADA FROM TEMPOR WHERE IDTEM = " & rc.Fields("TEMPOR").Value)
        '  .TextMatrix(.Rows - 1, 6) = rc.Fields("FMODI").Value
    

    
        If Not rc_ventas.EOF Then rc_ventas.MoveNext
        Loop Until rc_ventas.EOF
    
   
 
    

    'guardar totales en el registro ...
    
    rc.fields("SUBTOT") = T_SubTotal
    rc.fields("IVATOT") = t_iva
    rc.fields("RETOT") = T_RE
    rc.Update
    
    
    '=================================================
    'INSERTAR LINEA DE SUBTOTAL
    '=================================================
    
    .Rows = .Rows + 1
    .TextMatrix(.Rows - 1, 3) = "T. Artículos"
    
    .TextMatrix(.Rows - 1, 6) = T_Unidades
    .TextMatrix(.Rows - 1, 8) = T_SubTotal
    .TextMatrix(.Rows - 1, 10) = t_iva
    .TextMatrix(.Rows - 1, 11) = T_Total
     
     
    .Row = .Rows - 1
    For var = 1 To .Cols - 1
    .Col = var
    .CellBackColor = vbWhite
    Next var

    End With
 
    End If
    
    
    With fg
    
   ' T_Unidades = 0
   ' T_SubTotal = 0
   ' T_IVA = 0
   ' T_RE = 0
    
    '=================================================
    'AHORA CARGAR LOS ARREGLOS PARA ESTA VENTA
    '=================================================
     If rc_arreglos.State = 1 Then rc_arreglos.Close
     rc_arreglos.Open "SELECT ID, CODART, TEMPOR, CODTALLA, CODCOL, PVP, ESTADO, FMODI FROM ARREGLOS WHERE ESTADO = 1 AND CODVEN = " & rc.fields("CODIGO") & " AND CODCAJ = " & CajaActual, locCnn, adOpenStatic, adLockOptimistic
          
     If rc_arreglos.RecordCount > 0 Then
        
        'añadir linea separadora
        'conta_lineas = conta_lineas + 1
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 3) = "ARREGLOS:"
        .Row = .Rows - 1
        .Col = 3
        .CellBackColor = vbWhite
    
        Do Until rc_arreglos.EOF
        
            conta_lineas = conta_lineas + 1
            .Rows = .Rows + 1
        
            .TextMatrix(.Rows - 1, 1) = 2  'codigo de operacion: ARREGLO
            .TextMatrix(.Rows - 1, 2) = rc_arreglos.fields("ID")
            
            If Not IsNull(rc_arreglos.fields("CODART")) Then
            .TextMatrix(.Rows - 1, 3) = Format(rc_arreglos.fields("CODART"), "00000") & " " & devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & rc_arreglos.fields("CODART") & " AND TEMPOR = " & rc_arreglos.fields("TEMPOR"))
            Else
            .TextMatrix(.Rows - 1, 3) = "Arreglos Varios"
            End If
            
            'talla
            If Not IsNull(rc_arreglos.fields("CODTALLA")) Then _
            .TextMatrix(.Rows - 1, 4) = Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & rc_arreglos.fields("CODTALLA")))
            
             'obtener el texto del color y su codigo de color (para colorear
             'la celda del grid)
             If rc_arreglos.fields("CODCOL") > 0 Then
      
                tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & rc_arreglos.fields("CODCOL"))
                .TextMatrix(.Rows - 1, 5) = devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & rc_arreglos.fields("CODCOL"))
                .Col = 5
                .Row = .Rows - 1
                .CellBackColor = tmpcodcolor
                .Col = 2
        
             End If
            
        

            'PVP
            .TextMatrix(.Rows - 1, 7) = rc_arreglos.fields("PVP")
            'SUBTOTAL
            .TextMatrix(.Rows - 1, 8) = rc_arreglos.fields("PVP")
            
            'IVA
            '.TextMatrix(.Rows - 1, 9) = 0
            
            'TOTAL
            .TextMatrix(.Rows - 1, 11) = rc_arreglos.fields("PVP")
            T_Arreglos = T_Arreglos + .TextMatrix(.Rows - 1, 11)
            
            If Not rc_arreglos.EOF Then rc_arreglos.MoveNext
        
        Loop
    
    End If
    
    '=================================================
    'INSERTAR LINEA DE SUBTOTAL
    '=================================================
    
    .Rows = .Rows + 1
    .TextMatrix(.Rows - 1, 3) = "T. Arreglos"
    .TextMatrix(.Rows - 1, 11) = T_Arreglos
    
    .Row = .Rows - 1
    
    For var = 1 To .Cols - 1
    .Col = var
    .CellBackColor = vbWhite
    Next
    
    '=================================================

        .ColFormat(7) = "Currency"
        .ColFormat(8) = "Currency"
        .ColFormat(11) = "Currency"
    
        '.SubtotalPosition = flexSTAbove
        '.subtotal flexSTSum, , 6, , vbBlue, vbWhite
        '.subtotal flexSTSum, , 8, , vbBlue, vbWhite
        '.subtotal flexSTSum, , 8, , vbBlue, vbWhite
        '.subtotal flexSTSum, , 11, , vbBlue, vbWhite
        '.TextMatrix(1, 4) = "Totales:"
        '.TextMatrix(1, 3) = "Total (" & conta_lineas & ") Art."
        
        .AddItem "", 1
        .TextMatrix(1, 3) = "Total (" & conta_lineas & ") Art."
        .TextMatrix(1, 6) = T_Unidades
        .TextMatrix(1, 8) = T_SubTotal
        .TextMatrix(1, 11) = T_Total + T_Arreglos
        
        'actualizar la etiqueta de totales
        If .TextMatrix(1, 11) <> "" Then
            lblTotal.Caption = Format(.TextMatrix(1, 11), "Currency")
        Else
            lblTotal.Caption = ""
        End If
        
     .Row = 1
    For var = 1 To .Cols - 1
    .Col = var
    .CellForeColor = vbWhite
    .CellFontBold = True
    .CellBackColor = vbBlue
    Next
                
    .Col = 2
    .AutoSize 1, .Cols - 1
    .Redraw = True
    '.Enabled = True
    
    DoEvents
    
    .HighLight = flexHighlightWithFocus
    .Editable = flexEDKbd
    .FocusRect = flexFocusHeavy
    .AllowBigSelection = False
    .AllowSelection = True
    
    DoEvents
    
    End With
    
   On Error GoTo 0
   Exit Sub

carga_grid_ventas_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_grid_ventas de Formulario frmCabVen"

End Sub


'Asigna el cliente seleccionado en el flexgrid, para llamar desde el flexclientes
Public Sub Asignar_cliente_flex(Codigo_Cliente As Long, codcaja As Byte)

With frmFlexCli
    
    If .seleccionado Then
    
        'asignar valores ...
        rc.fields("CODCLI") = Codigo_Cliente 'rccli.Fields("CODIGO")
        rc.fields("CAJACLI") = codcaja 'rccli.Fields("CODCAJA")
        lblCliente.Caption = devuelve_campo("SELECT RAZO FROM CLIENTES WHERE CODIGO = " & rc.fields("CODCLI") & " AND CODCAJA = " & rc.fields("CAJACLI"), locCnn)

    
    'dejar como estaba
    Else
    
      '  rc.Fields("CODCLI") = Null
      '  rc.Fields("CAJACLI") = Null
      '  lblCliente.Caption = ""
        
    End If
    
End With
    
        rc.Update
        

End Sub




'---------------------------------------------------------------------------------------
' Subrutina   : Abre_Grid_Clientes
' Fecha/Hora  : 18/01/2004 14:48
' Autor       : JCASTILLO
' Propósito   : Abre el grid de clientes, y obtiene un cliente para la venta
'---------------------------------------------------------------------------------------
Private Sub Abre_Grid_Clientes()
Dim cliSql As New clsSmartSQL
Dim rccli As New ADODB.Recordset


   On Error GoTo Abre_Grid_Clientes_Error

cliSql.AddTable "CLIENTES"
cliSql.AddOrderClause "CODCAJA"
cliSql.AddOrderClause "CODIGO"

rccli.Open cliSql.SQL, locCnn, adOpenDynamic, adLockReadOnly

With frmFlexCli

    .Caption = "Clientes ..."
    Set .miosql = cliSql
            
    .desde_ventas = True
    Set .miRc = rccli
       
    DoEvents
  
    Me.Visible = False
  
    '.MDIChild = True
    .Show
        
    

    'Set frmFlexCli = Nothing
    
    DoEvents
    
End With

   Set cliSql = Nothing
   
   On Error GoTo 0
   Exit Sub

Abre_Grid_Clientes_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Abre_Grid_Clientes de Formulario frmCabVen"

End Sub

'---------------------------------------------------------------------------------------
' Procedimiento: Aceptar_Venta
' Fecha/Hora    : 28/01/2004 13:47
' Autor            : JCastillo
' Propósito      :Acepta la venta actual (descontar unidades de stock y cambiar
'                       el estado de la venta
'---------------------------------------------------------------------------------------
Private Sub Aceptar_Venta()
Dim tmpconn As ADODB.Connection

    On Error GoTo Aceptar_Venta_Error
        
    If MsgBox("¿Desea Aceptar la Venta actual?", vbQuestion + vbYesNo, titulo) = vbNo Then
        Set tmpconn = Nothing
        Exit Sub
    End If
    
    'imprimir el ticket actual
    Call Imprime_Ticket(frmFcobro.I_Cambio, rc, rc_ventas, rc_arreglos)
    
    DoEvents
    
    If rc.fields("ESTADO") = 1 Then Exit Sub
        
    Set tmpconn = New ADODB.Connection
    
    With tmpconn
        .CursorLocation = adUseServer
        .Open strLocCnn
        .BeginTrans
    End With
        
    If rc_ventas.RecordCount > 0 Then
    
    With rc_ventas
        
        If Not .BOF Then .MoveFirst
        
        Do Until .EOF
            'quitar unidades de stock
            stock .fields("CODART"), .fields("TEMPOR"), .fields("CODTALLA"), .fields("CODCOL"), AlmacenActual, .fields("UNIDADES"), False, tmpconn
            DoEvents
            
            .MoveNext
        Loop
    
    End With
    
    
    End If
    
    
    'marcar los arreglos como servidos
    With rc_arreglos
    
    If .RecordCount > 0 Then
    
        If Not .BOF Then .MoveFirst
        
        Do Until .EOF
            .fields("ESTADO") = 2
            'fmodi como fecha actual ...
            .fields("FMODI") = Date
            .Update
            .MoveNext
        Loop
            
    End If
    
    End With
        
    'marcar los vales como aceptados en la base de datos
    Call frmFcobro.Marcar_Vales_Como_Aceptados(tmpconn)
    
    'añadir una deuda si es pago pendiente
    If frmFcobro.impor_deuda > 0 Then
         
        Call inserta_deuda_cliente(rc.fields("CODCLI"), rc.fields("CAJACLI"), rc.fields("CODIGO"), frmFcobro.impor_deuda, "", tmpconn)
    
    End If
        
    With tmpconn
        .CommitTrans
        .Close
    End With
        
    Set tmpconn = Nothing
    
    DoEvents
    
    'poner como aceptada
    rc.fields("ESTADO") = 1
    rc.fields("FHORA") = Date
    rc.Update

    DoEvents
    
    rc.Requery

    'si no hay mas registros, borrar todo
    If rc.RecordCount <= 0 Then
            
        fg.Rows = 1
        fg.Clear
        lblTotal.Caption = ""
        lblDependiente.Caption = ""
    
    End If

   On Error GoTo 0
   Exit Sub

   With tmpconn
   
   If .State = 1 Then
    .RollbackTrans
    .Close
   End If
   
   Set tmpconn = Nothing
   
   End With
   
Aceptar_Venta_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Aceptar_Venta de Formulario frmCabVen"
End Sub


