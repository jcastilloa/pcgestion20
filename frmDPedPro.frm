VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmDPedPro 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Añadir artículos al pedido ..."
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9990
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   Begin PCGestion.ucGrdBttn cbAñadeTallaColor 
      Height          =   510
      Left            =   1335
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   3165
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   900
      Caption         =   "+"
      ForeColor       =   65280
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "frmDPedPro.frx":0000
      GradientColor1  =   16558731
      GradientColor2  =   11250603
   End
   Begin VSFlex8LCtl.VSFlexGrid vsTallasColores 
      Height          =   2745
      Left            =   2265
      TabIndex        =   32
      Top             =   2520
      Width           =   7680
      _cx             =   13547
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
      Rows            =   1
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDPedPro.frx":001C
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
   Begin PCGestion.chameleonButton cmdPrevious 
      Height          =   630
      Left            =   1050
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5730
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
      MICON           =   "frmDPedPro.frx":00B3
      PICN            =   "frmDPedPro.frx":00CF
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
      Top             =   5310
      Width           =   9915
      _ExtentX        =   16510
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
      Left            =   15
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5730
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
      MICON           =   "frmDPedPro.frx":0DA1
      PICN            =   "frmDPedPro.frx":0DBD
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
      Left            =   7830
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5730
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
      MICON           =   "frmDPedPro.frx":1AF3
      PICN            =   "frmDPedPro.frx":1B0F
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
      Left            =   8895
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5730
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
      MICON           =   "frmDPedPro.frx":27E1
      PICN            =   "frmDPedPro.frx":27FD
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
      Left            =   15
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6390
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
      MICON           =   "frmDPedPro.frx":3533
      PICN            =   "frmDPedPro.frx":354F
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
      TabIndex        =   7
      Top             =   6390
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
      MICON           =   "frmDPedPro.frx":4229
      PICN            =   "frmDPedPro.frx":4245
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6390
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
      MICON           =   "frmDPedPro.frx":4B1F
      PICN            =   "frmDPedPro.frx":4B3B
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
      Left            =   5115
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6390
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
      MICON           =   "frmDPedPro.frx":5399
      PICN            =   "frmDPedPro.frx":53B5
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
      Left            =   7785
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6390
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
      MICON           =   "frmDPedPro.frx":5C8F
      PICN            =   "frmDPedPro.frx":5CAB
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
      Left            =   8895
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6390
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
      MICON           =   "frmDPedPro.frx":687D
      PICN            =   "frmDPedPro.frx":6899
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.miText ioUNIDADES 
      Height          =   525
      Left            =   3330
      TabIndex        =   1
      Top             =   930
      Width           =   825
      _ExtentX        =   1455
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
   Begin PCGestion.miText ioDCTO 
      Height          =   525
      Left            =   4785
      TabIndex        =   2
      Top             =   930
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
   Begin PCGestion.chameleonButton cbSeleccionaArticulo 
      Height          =   525
      Left            =   5070
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5820
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   926
      BTYPE           =   9
      TX              =   "Buscar ..."
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
      MICON           =   "frmDPedPro.frx":7573
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblARTICULO 
      Height          =   360
      Left            =   1005
      Top             =   525
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   635
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   15640462
      Colour2         =   11311500
      CaptionAlignment=   1
   End
   Begin PCGestion.miCombo cbTALLA 
      Height          =   540
      Left            =   960
      TabIndex        =   5
      Top             =   1980
      Width           =   4005
      _ExtentX        =   7064
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
      Left            =   2685
      TabIndex        =   4
      Top             =   1485
      Width           =   7245
      _ExtentX        =   12779
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
      Left            =   6405
      TabIndex        =   6
      Top             =   1980
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
   Begin PCGestion.ucGrdBttn cbBorraTallaColor 
      Height          =   510
      Left            =   1335
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   3735
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   900
      Caption         =   "-"
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "frmDPedPro.frx":758F
      GradientColor1  =   16558731
      GradientColor2  =   11250603
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel1 
      Height          =   315
      Left            =   180
      Top             =   2535
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      GradientType    =   2
      Caption         =   "Tallas y Colores"
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
   Begin PCGestion.miText ioPRECOM 
      Height          =   480
      Left            =   975
      TabIndex        =   0
      Top             =   930
      Width           =   1305
      _ExtentX        =   2302
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
      dspFormat       =   ""
      Enabled         =   -1  'True
      EsPassword      =   -1  'True
   End
   Begin PCGestion.miText cbTIPOIVA 
      Height          =   525
      Left            =   6435
      TabIndex        =   3
      Top             =   930
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
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   210
      TabIndex        =   39
      Top             =   3765
      Width           =   270
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   210
      TabIndex        =   38
      Top             =   3225
      Width           =   270
   End
   Begin VB.Label lblLINEA 
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
      Left            =   2040
      TabIndex        =   37
      Top             =   60
      Width           =   3330
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "AÑADIR"
      Height          =   315
      Left            =   480
      TabIndex        =   35
      Top             =   3285
      Width           =   795
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "QUITAR"
      Height          =   300
      Left            =   480
      TabIndex        =   36
      Top             =   3840
      Width           =   795
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "COLOR"
      Height          =   330
      Left            =   5655
      TabIndex        =   31
      Top             =   2070
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TALLA"
      Height          =   330
      Left            =   255
      TabIndex        =   30
      Top             =   2055
      Width           =   675
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORIA TALLA"
      Height          =   345
      Left            =   630
      TabIndex        =   29
      Top             =   1575
      Width           =   1935
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
      TabIndex        =   27
      Top             =   60
      Width           =   960
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ARTICULO"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   60
      TabIndex        =   26
      Top             =   555
      Width           =   900
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PRECIO C. FINAL"
      Height          =   600
      Left            =   7755
      TabIndex        =   25
      Top             =   885
      Width           =   975
   End
   Begin VB.Label lblPrecomFin 
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
      Height          =   345
      Left            =   8715
      TabIndex        =   24
      Top             =   1005
      Width           =   1215
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "IVA"
      Height          =   330
      Left            =   5955
      TabIndex        =   23
      Top             =   1005
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DCTO"
      Height          =   285
      Left            =   4170
      TabIndex        =   22
      Top             =   1005
      Width           =   570
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PRECIO COMPRA"
      Height          =   630
      Left            =   30
      TabIndex        =   21
      Top             =   870
      Width           =   915
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UNIDADES"
      Height          =   300
      Left            =   2310
      TabIndex        =   20
      Top             =   1020
      Width           =   990
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
      Left            =   7485
      TabIndex        =   10
      Top             =   60
      Width           =   2445
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      Height          =   315
      Left            =   90
      TabIndex        =   9
      Top             =   90
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ultima modificación"
      Height          =   315
      Left            =   5370
      TabIndex        =   8
      Top             =   90
      Width           =   2070
   End
End
Attribute VB_Name = "frmDPedPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'El recordset para almacenar el detalle del formulario desde donde se llama
Public WithEvents rc As ADODB.Recordset
Attribute rc.VB_VarHelpID = -1

Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim prime As Boolean

Dim oSQL As New clsSmartSQL

Public NumeroPedido As Long

Dim tmp_precom As Double
Dim tmp_cmprecom As Boolean
Dim tmp_iva As Single
Dim tmp_cmiva As Boolean

Dim Iva_Descontado As Boolean

'---------------------------------------------------------------------------------------
' Procedure : cbAñadeTallaColor_Click
' DateTime  : 09/11/2003 19:10
' Author    : Administrador
' Purpose   : Añadir otro registro para el mismo artículo, para permitir
'             otra combinación talla/color
'---------------------------------------------------------------------------------------
'
Private Sub cbAñadeTallaColor_Click()
Dim tmpcodigo As Variant
Dim tmpcodart As Long
Dim tmpunidades As Single
Dim tmpdcto1 As Single
Dim tmpdcto2 As Single
Dim tmpalmorig As Byte
Dim tmpmodelo As String
Dim tmpprecom As Single
Dim tmpcolor As Integer
Dim tmpcattall As Integer
Dim tmpiva As Single
Dim tmplinea As Long

'si no desea, salir
   On Error GoTo cbAñadeTallaColor_Click_Error

If mbEditFlag Or mbAddNewFlag Then
    lblstatus.Caption = "Debe GUARDAR los cambios antes de añadir TALLA/COLOR"
    Exit Sub
End If

If cbAñadeTallaColor.Enabled = False Then Exit Sub
cbAñadeTallaColor.Enabled = False

tmplinea = rc.fields("LINEA")
If MsgBox("¿Desea añadir otra talla/color para el artículo actual? " & Chr(13) & ioCODIGO.Caption & "   " & lblArticulo.Caption & "(" & tmplinea & ")", vbQuestion + vbYesNo) = vbNo Then Exit Sub

  With rc
  
  '  If Not IsNull(.fields("CODCOL")) Then
  '      lblstatus.Caption = "No se permite Color en blanco"
  '      Exit Sub
  '  End If
    
  '  If Not IsNull(.fields("CODTALLA")) Then
  '      lblstatus.Caption = "No se permite Talla en blanco"
  '      Exit Sub
  '  End If
           
    DoEvents
    
    .Find "LINEA =" & tmplinea, , adSearchForward, 1
    
    DoEvents
    
    'obtener datos del registro actual para hacer la copia
    If Not IsNull(.fields("CODART")) Then tmpcodart = .fields("CODART").Value
    If cbTIPOIVA.Text <> "" Then tmpiva = CSng(cbTIPOIVA.Valor)
    If Not IsNull(.fields("UNIDADES")) Then tmpunidades = .fields("UNIDADES").Value
    If Not IsNull(.fields("DCTO")) Then tmpdcto1 = .fields("DCTO").Value
    If Not IsNull(.fields("PRECOM")) Then tmpprecom = .fields("PRECOM").Value
    If Not IsNull(.fields("ALMORIG")) Then tmpalmorig = .fields("ALMORIG").Value
    If Not IsNull(.fields("CODCOL")) Then tmpcolor = .fields("CODCOL")
    
    If cbCATTALL.Text <> "" Then tmpcattall = CInt(cbCATTALL.Text)
  
    tmpmodelo = lblArticulo.Caption
    
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
      If .EditMode <> adEditNone Then .CancelUpdate
    End If
    
    .AddNew
    
    tmpcodigo = devuelve_campo("select max(LINEA) + 1 from DETPEDPRO where NUMERO = " & NumeroPedido, locCnnSP)
        
    'meter los datos del registro
    
    If tmpcodigo = "@" Then tmpcodigo = 1
    
        .fields("LINEA") = tmpcodigo
        .fields("ALMORIG") = tmpalmorig
        .fields("TEMPOR") = frmPedProv.TemporadaWrk  ' TemporadaActual
        .fields("CODART") = tmpcodart
        .fields("UNIDADES") = tmpunidades
        .fields("DCTO") = tmpdcto1
        cbTIPOIVA.Text = tmpiva
        .fields("IVA") = tmpiva
        .fields("PRECOM") = tmpprecom
        .fields("CODCOL") = tmpcolor
        .fields("NUMERO") = NumeroPedido
        If tmpcattall > 0 Then cbCATTALL.Text = tmpcattall
        lblArticulo.Caption = tmpmodelo
    
    
    .Update

    Set tmpcodigo = Nothing
   
    lblstatus.Caption = "Agregar registro"
    mbAddNewFlag = True
    SetButtons False
    cbActualizar.Visible = True
    
    ioUNIDADES.SetFocus
  '
  End With


tmpmodelo = ""


  On Error GoTo 0
   Exit Sub

cbAñadeTallaColor_Click_Error:

  
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cbAñadeTallaColor_Click of Formulario frmDPedPro"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cbBorraTallaColor_Click
' DateTime  : 10/11/2003 17:06
' Author    : Administrador
' Purpose   : Borra la combinación TALLA/COLOR del registro actual. (Borra
'                el registro actual). Además reorganiza los numeros LINEA para que sea
'                correlativo.
'---------------------------------------------------------------------------------------
Private Sub cbBorraTallaColor_Click()
Dim tmpdesctalla As String
Dim tmpdesccolor As String
'Dim num_regs As Single
Dim tmpcodigo As Long

On Error GoTo cbBorraTallaColor_Click_Error


If rc.RecordCount = 0 Then Exit Sub

If mbEditFlag Or mbAddNewFlag Then
    lblstatus.Caption = "Debe GUARDAR los cambios antes de añadir TALLA/COLOR"
    Exit Sub
End If

'que cancele cualquier operación pendiente
'If mbAddNewFlag Or mbEditFlag Then Call cbCancelar_Click

If cbTALLA.Text <> "" Then tmpdesctalla = devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & cbTALLA.Text, locCnnSP)
If cbCOLOR.Text <> "" Then tmpdesccolor = devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & cbCOLOR.Text, locCnnSP)

'si no quiere salir
If MsgBox("¿Desea elimnar la combinación talla/color actual?" & Chr(13) & "Talla: [" & tmpdesctalla & "] - Color: [" & tmpdesccolor & "]" & Chr(13) & ioCODIGO.Caption & "  " & lblArticulo.Caption, vbQuestion + vbYesNo) = vbNo Then
        tmpdesctalla = ""
        tmpdesccolor = ""
    Exit Sub
End If

        'comprobar si ya esta metido en almacén
        If rc.fields("METIDO") = True Then
            
            If MsgBox("El artículo ya se encuentra en almacén. ¿Desea quitar de todas formas?", vbQuestion + vbYesNo, titulo) = vbNo Then
                'salir
                Exit Sub
            End If
            
            'quitar unidades del almacén ...
            Call stock(rc.fields("CODART"), rc.fields("TEMPOR"), rc.fields("CODTALLA"), rc.fields("CODCOL"), AlmacenActual, rc.fields("UNIDADES"), False, locCnnSP)
             
        End If
        
        'borrar el registro actual
        tmpcodigo = rc.fields("CODART")
        rc.Delete
        rc.UpdateBatch adAffectAllChapters
        
        DoEvents
       
        'si no hay mas registros, actualizar el grid de pedidos
        'y cerrar el formulario actual
        If rc.RecordCount = 0 Then
            Call frmPedProv.refresca_grid_externo(True)
            
            
            
            Unload Me
            Exit Sub
        End If
        
        Call cmdFirst_Click
        
        If Not rc.EOF Then
            rc.Find "CODART = " & tmpcodigo
            
            If rc.EOF Then Call cmdFirst_Click
        End If
        
        Call frmPedProv.refresca_grid_externo(True)
        
        
With cbTALLA
    Set .DataSource = Nothing
    .DataField = ""
    .DataField = "CODTALLA"
    Set .DataSource = rc
End With

With ioUNIDADES
    Set .DataSource = Nothing
    .DataField = ""
    .DataField = "UNIDADES"
    Set .DataSource = rc
End With

With cbCOLOR
    Set .DataSource = Nothing
    .DataField = ""
    .DataField = "CODCOL"
    Set .DataSource = rc
End With
        
        
        
       ' rc.UpdateBatch adAffectAllChapters
        
      '  rc.Requery
                
      '  locCnnSP.Execute "DELETE FROM DETPEDPRO WHERE LINEA = " & rc.Fields("LINEA") & " AND NUMERO  = " & rc.Fields("NUMERO")
        
      '  rc.CancelBatch
        
        'rc.Requery
 
        'reorganizar numeros de linea
        'If Not rc.EOF Then rc.MoveLast
        
        'total registros
        'num_regs = rc.RecordCount
        
        'Do Until rc.BOF
           
        'rc.Fields("LINEA") = num_regs
        'rc.Update
        'rc.MovePrevious
        'num_regs = num_regs - 1
        
        'Loop
                
        'If Not rc.BOF Then rc.MoveFirst

        'rc.UpdateBatch adAffectAllChapters
        
        tmpdesctalla = ""
        tmpdesccolor = ""
        
                
        lblstatus.Caption = "El registro se ha borrado satisfactoriamente"
                       
                       
        If Not rc.RecordCount <= 0 Then
            Call carga_grid(rc.fields("CODART"))
        Else
            vsTallasColores.Clear
            DoEvents
            Call cbcerrar_Click
            DoEvents
        End If

   On Error GoTo 0
   Exit Sub

cbBorraTallaColor_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cbBorraTallaColor_Click of Formulario frmDPedPro"
End Sub



Private Sub cbCATTALL_Validate(Cancel As Boolean)

If (Not mbAddNewFlag) And (Not mbEditFlag) Then Exit Sub

If cbCATTALL.Text = "" Then Exit Sub

With cbTALLA
    .ConexionString = locCnnSP
    .LenCodigo = 3
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM TALLAS WHERE (MBAJA = 0) AND (CODIGO > 0) AND CATTALL = " & cbCATTALL.Text & " ORDER BY CODIGO"
    .DataField = "CODTALLA"
    .carga
    .CodigoWidth = 500
    Set .DataSource = rc
End With

End Sub


Private Sub cbSeleccionaArticulo_Click()
'Dim tmpbook As Variant
Dim artsql As New clsSmartSQL
Dim mrc As New ADODB.Recordset

On Error GoTo cbSeleccionaArticulo_Click_Error

With artsql
    .AddTable "MAARTIC"
    .AddOrderClause "CODIGO"
    .AddSimpleWhereClause "MBAJA", "0"
    .AddSimpleWhereClause "HIST", "0"
End With

mrc.Open artsql.SQL, locCnnSP, adOpenStatic, adLockReadOnly

    With frmFlexArt
    
    Set .miosql = artsql
    Set .miRc = mrc
    'el recordset de este formulario
    
    Set .rc_detalle = rc
    .NumeroPedido = NumeroPedido
    .add_en_detalle = True
    
    .tmpcodprov = frmPedProv.cbCODPROV.Text
    .Temporada_Defecto = frmPedProv.TemporadaWrk
    
    .Show 1
    DoEvents
    
    End With
    
    mrc.Close
    Set mrc = Nothing
'tmpbook = rc.Bookmark
     
'rc.Bookmark = tmpbook
'Set tmpbook = Nothing
Set artsql = Nothing

Call carga_grid(rc.fields("CODART"))


   On Error GoTo 0
   Exit Sub

cbSeleccionaArticulo_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbSeleccionaArticulo_Click de Formulario frmDPedPro"
    
End Sub

Private Sub cbTIPOIVA_GotFocus()
    If cbTIPOIVA.Text <> "" Then tmp_iva = cbTIPOIVA.Text
End Sub

Private Sub cbTIPOIVA_Validate(Cancel As Boolean)

   On Error GoTo cbTIPOIVA_Validate_Error
      
If (Not mbAddNewFlag) And (Not mbEditFlag) Then Exit Sub

tmp_cmiva = True
      
Call calcula_descuentos

   On Error GoTo 0
   Exit Sub

cbTIPOIVA_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbTIPOIVA_Validate de Formulario frmDPedPro"

End Sub

Private Sub Form_Activate()
Dim tmpbok As Variant

   On Error GoTo Form_Activate_Error

If Not prime Then

  If rc.RecordCount = 0 Then
        
        If MsgBox("No se encuentran Artículos para este pedido. ¿Añadir?", vbYesNo + vbQuestion, "Artículos") = vbNo Then
        Unload Me
        Else
        Call cbAgregar_Click
        End If
        
  Else
        
        tmpbok = rc.Bookmark
        Call cmdFirst_Click
        Call cbCancelar_Click
        
        rc.Bookmark = tmpbok
        
        Call carga_grid(rc.fields("CODART"))
        
        Set tmpbok = Nothing
        
        Call cbedicion_Click
         
            
  End If
  
  
prime = True
End If

   On Error GoTo 0
   Exit Sub

Form_Activate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Activate of Formulario frmDPedPro"
    
End Sub

Private Sub Form_Load()
  
   On Error GoTo Form_Load_Error
   
   Move (Screen.Width - Width) \ 2, Separacion_MDIForm

  With locCnnSP
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnnSP
    End If
   End With
   
  'Set rc = New Recordset
  'oSQL.AddTable "RDETPEDPRO"
  'oSQL.AddOrderClause "LINEA"
  
  'oSQL.AddSimpleWhereClause "NUMERO", NumeroPedido
  
  'rc.Open oSQL.SQL, locCnn, adOpenStatic, adLockOptimistic
    

''''''''
'////////////////////////////////////////////////////////
' Enlazar controles:
'////////////////////////////////////////////////////////

  With ioCODIGO
  Set .DataSource = rc
        .DataField = "CODART"
End With
 
With cbTALLA
    .ConexionString = locCnnSP
    .LenCodigo = 3
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM TALLAS WHERE (MBAJA = 0) AND (CODIGO > 0) ORDER BY CODIGO"
    .DataField = "CODTALLA"
    .carga
    .CodigoWidth = 500
    Set .DataSource = rc
End With



With cbCOLOR
    .ConexionString = locCnnSP
    .LenCodigo = 3
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM COLORES WHERE (MBAJA = 0) AND (CODIGO > 0) ORDER BY CODIGO"
    .DataField = "CODCOL"
    .carga
    .CodigoWidth = 500
    Set .DataSource = rc
End With


With ioPRECOM
 'Set .DataSource = rc
.Alineacion = 1
'.dspFormat = "Currency"
.SoloNumeros = True
.PermitirBlanco = False
   'Set .DataSource = rc
        ' .DataField = "PRECOM"
   ' .displayformat = "00000.00 €"
   .dspFormat = "Currency"
   '.Mask = "#####.##"
    
       ' .Format = "####0.##"
End With

'With ioPRECOM
'    Set .DataSource = rc
'        .DataField = "PRECOM"
'        .Alignment = 1
'End With

'With ioIVA
'Set .DataSource = rc
'    .DataField = "IVA"
'End With

With cbTIPOIVA
'    .ConexionString = locCnn
'    .LenCodigo = 3
'    .SQLString = "SELECT IVA, cast(IVA AS char(" & .LenCodigo & ")) + ' %'  FROM IVA WHERE MBAJA = 0 ORDER BY CODIGO"
    '.DataField = "IVA"
'    .carga
'    DoEvents
'    .CodigoWidth = 500
    .PermitirBlanco = False
    .SoloNumeros = True
    'Set .DataSource = rc
End With

    
    With ioUNIDADES
        .DataField = "UNIDADES"
        Set .DataSource = rc
        .LongMaxima = 5
        .PermitirBlanco = False
        .SoloNumeros = True
        .dspFormat = "00.00"
        .Alineacion = 1
        .PermitirBlanco = False
    End With
        
    With ioDCTO
        .DataField = "DCTO"
        Set .DataSource = rc
        .LongMaxima = 5
        .PermitirBlanco = False
        .SoloNumeros = True
        .dspFormat = "00.00"
        .Alineacion = 1
    End With

  With ioFMODI
  Set .DataSource = rc
        .DataField = "FMODI"
  End With
  
   With cbCATTALL
      .LenCodigo = 4
      .CodigoWidth = 670
      .ConexionString = locCnnSP
      .SQLString = "SELECT CODIGO, DESCRIPCION FROM CATTALL ORDER BY CODIGO"
      .carga
  End With
      
  With vsTallasColores
   .AllowSelection = True
   .HighLight = flexHighlightNever
  End With
  
  mbDataChanged = False

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmDPedPro"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : carga_grid
' DateTime  : 09/11/2003 13:37
' Author    : Administrador
' Purpose   : Llena el grid con las unidades de cada talla y color
'             del acticulo actual
'---------------------------------------------------------------------------------------
'
Private Sub carga_grid(fgCodart As Long)
Dim tmprc As New ADODB.Recordset
Dim tmplinea As Long
Dim tmpcodcolor As Long

   On Error GoTo carga_grid_Error

    tmprc.Open "SELECT LINEA, UNIDADES, CODTALLA, CODCOL FROM DETPEDPRO WHERE CODART = " & fgCodart & " AND NUMERO = " & rc.fields("NUMERO") & " ORDER BY LINEA", locCnnSP, adOpenDynamic, adLockReadOnly
    
    With vsTallasColores
        .Clear
        .Cols = 5
        
        .TextMatrix(0, 2) = "UNIDADES"
        .TextMatrix(0, 3) = "TALLA"
        .TextMatrix(0, 4) = "COLOR"
        
        .Rows = 1
    
    Do Until tmprc.EOF

        .Rows = .Rows + 1
        
        'numero de linea
        .TextMatrix(.Rows - 1, 1) = tmprc.fields("LINEA").Value
        
        If Not IsNull(tmprc.fields("UNIDADES")) Then
        .TextMatrix(.Rows - 1, 2) = tmprc.fields("UNIDADES").Value
        End If
        
        If Not IsNull(tmprc.fields("CODTALLA")) And tmprc.fields("CODTALLA") <> 0 Then
        .TextMatrix(.Rows - 1, 3) = devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & tmprc.fields("CODTALLA").Value, locCnnSP)
        End If
        
        'obtener el texto del color y su codigo de color (para colorear
        'la celda del grid)
        If Not IsNull(tmprc.fields("CODCOL")) And tmprc.fields("CODCOL") <> 0 Then
        
            tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & tmprc.fields("CODCOL"), locCnnSP)
            .TextMatrix(.Rows - 1, 4) = devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & tmprc.fields("CODCOL"), locCnnSP)
            .Col = 4
            .Row = .Rows - 1
            .CellBackColor = tmpcodcolor
            .Col = 2
        
        End If
            
    tmprc.MoveNext
    
    Loop
    
    '.SubtotalPosition = flexSTBelow
    .subtotal flexSTSum, , 2, , vbBlue, vbWhite
    .TextMatrix(1, 4) = "Total"
    
    End With
    
    tmprc.Close
    Set tmprc = Nothing
    
    

   Exit Sub

carga_grid_Error:
    
    
    Set tmprc = Nothing
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure carga_grid of Formulario frmDPedPro"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If (mbEditFlag Or mbAddNewFlag) And KeyCode <> vbKeyF10 Then Exit Sub

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
      
      Case vbKeyF10
            
            KeyCode = 0
            
            If Not Iva_Descontado Then
            
                If (rc.EditMode = adEditNone) Or Screen.ActiveControl.Name <> "ioPRECOM" Then Exit Sub
            
                'descontar el IVA del PRECIO
                If ioPRECOM.Text <> "" And cbTIPOIVA.Text <> "" Then
            
                    ioPRECOM.Text = Format(ioPRECOM.Text - ((ioPRECOM.Text * cbTIPOIVA.Text) / 100), "Currency")
                    lblstatus.Caption = "Se ha descontado el " & cbTIPOIVA.Text & " % de IVA a el precio de compra"
                    Iva_Descontado = True
                    Exit Sub
                                   
                Else
            
                    lblstatus.Caption = "Debe establecer un precio y un IVA para realizar la operación"
                    Exit Sub
                        
               End If
            End If
            
      'Añadir talla/color
      Case vbKeyA
      
        Call cbAñadeTallaColor_Click
        KeyCode = 0
      
      'Quitar talla/color
      Case vbKeyQ
      
        Call cbBorraTallaColor_Click
        KeyCode = 0
      
                  
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

On Error Resume Next

If rc.EditMode <> adEditNone Then rc.CancelUpdate

'rc.Requery
DoEvents

'rc.Close
'Set rc = Nothing

'Set oSQL = Nothing

''If Me.MDIChild = True Then frmMenuTactil.Show
Set frmDPedPro = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub ioCODIGO_Change()
ioCODIGO.Caption = Format(ioCODIGO.Caption, "00000")
End Sub




Private Sub ioDCTO_Validate(Cancel As Boolean)

   On Error GoTo ioDCTO_Validate_Error


If (Not mbAddNewFlag) And (Not mbEditFlag) Then Exit Sub

With ioDCTO
If Trim(.Text) <> "" Then
    If CDbl(Replace(.Text, ".", ",")) >= 100 Then
    lblstatus.Caption = "No se permite un Descuento mayor o igual al 100%"
    .Text = 0
    .CancelarValidacion
    Cancel = True
    Else
        lblstatus.Caption = ""
    End If
End If

End With

Call calcula_descuentos

   On Error GoTo 0
   Exit Sub

ioDCTO_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ioDCTO_Validate of Formulario FrmDPedPro"

End Sub


'Private Sub ioIVA_Change()
 '   If InStr(1, cbTIPOIVA.Text, "%") = 0 Then ioIVA.Caption = cbTIPOIVA.Text & " %"
'End Sub

Private Sub ioprecom_Change()
    'ioPRECOM.Caption = Format(ioPRECOM.Text, "currency")
    Call calcula_descuentos
End Sub








Private Sub ioPRECOM_GotFocus()

    If (Not mbAddNewFlag) And (Not mbEditFlag) Then Exit Sub
    
    If ioPRECOM.Text = "" Then ioPRECOM.Text = 0
    tmp_precom = CDbl(ioPRECOM.Text)
    lblstatus.Caption = "Pulse F10 para descontar el IVA (para precios IVA incluido)"

End Sub


Private Sub ioPRECOM_Validate(Cancel As Boolean)
        
If rc.RecordCount = 0 Then Exit Sub

If (Not mbAddNewFlag) And (Not mbEditFlag) Then Exit Sub
    
tmp_cmprecom = True
With ioPRECOM

    If .Text = "" Then
        lblstatus.Caption = "Debe establecer un precio de COMPRA para el artículo"
        .SetFocus
        .CancelarValidacion
        Cancel = True
    ElseIf CDbl(.Valor) = 0 Then
        lblstatus.Caption = "Debe establecer un precio de COMPRA para el artículo"
        .SetFocus
        .CancelarValidacion
        Cancel = True
        Exit Sub
    End If
 
End With

   Call calcula_descuentos
   
End Sub

Private Sub ioUNIDADES_Validate(Cancel As Boolean)

If (Not mbAddNewFlag) And (Not mbEditFlag) Then Exit Sub

With ioUNIDADES

If Trim(.Valor) = "" Then Exit Sub

If CDbl(.Valor) = 0 Then

    lblstatus.Caption = "Unidades no puede ser 0"
    '.SetFocus
    '.CancelarValidacion
    'Cancel = True
    Exit Sub

Else

    Call calcula_descuentos

End If

End With

End Sub



Private Sub lblprecomfin_Change()
    lblPrecomFin.Caption = Format(lblPrecomFin.Caption, "currency")
End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : calcula_descuentos
' Fecha/Hora  : 11/01/2004 16:38
' Autor       : JCASTILLO
' Propósito   : Calcular precio final de compra del artículo
'---------------------------------------------------------------------------------------
Private Sub calcula_descuentos()
Dim tmpprecom As Single
Dim tmpiva As Single
Dim tmpdcto As Single

'si estan los descuentos vacios poner el precio directamente
 On Error GoTo calcula_descuentos_Error
    
    If IsNull(rc.fields("PRECOM").Value) Then Exit Sub
        
        'obtener precio de compra
        tmpprecom = rc.fields("PRECOM").Value
           
        'quitar descuento
        If ioDCTO.Text <> "" Then
            tmpdcto = ioDCTO.Text
        Else
            tmpdcto = 0
        End If
        
        If tmpdcto > 0 Then tmpprecom = tmpprecom - ((tmpprecom * CSng(tmpdcto)) / 100)
        
        'añadir iva
        If cbTIPOIVA.Text <> "" Then
            tmpiva = cbTIPOIVA.Text
        Else
            tmpiva = 0
        End If
        
        If tmpiva > 0 Then tmpprecom = tmpprecom + ((tmpprecom * CSng(tmpiva)) / 100)
        
        'mostrar resultado
        lblPrecomFin.Caption = Format(tmpprecom, "currency")
    

   On Error GoTo 0
   Exit Sub

calcula_descuentos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcula_descuentos of Formulario FrmMntArt"

End Sub

Private Sub rc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  
  'On Error Resume Next
  
  Dim tmpart As Variant
  
  With rc
  
  If .AbsolutePosition > 0 Then
  lblstatus.Caption = "Registro: " & CStr(.AbsolutePosition)
  
  lblLINEA.Caption = "Linea : " & Format(.fields("LINEA"), "00000")

  
  'cargar la categoria de la talla
  If Not IsNull(.fields("CODTALLA")) Then
  cbCATTALL.Text = devuelve_campo("SELECT CATTALL FROM TALLAS WHERE CODIGO = " & .fields("CODTALLA"), locCnnSP)
  Else
  cbCATTALL.Text = ""
  End If
  
  If .fields("CODART").Value <> "" Then
  
  tmpart = devuelve_matriz("SELECT MODELO, REF FROM MAARTIC WHERE CODIGO = " & .fields("CODART").Value & " AND TEMPOR = " & rc.fields("TEMPOR"), locCnnSP)
  
  If Not IsArray(tmpart) Then
    lblArticulo.Caption = "Error al leer el artículo"
    Exit Sub
  End If
  
  lblArticulo.Caption = Trim(tmpart(1)) & " - " & Trim(tmpart(0)) '  devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & .fields("CODART").Value & " AND TEMPOR = " & rc.fields("TEMPOR"), locCnnSP)
  
  Set tmpart = Nothing
  
  End If
      
  If Not IsNull(.fields("PRECOM").Value) Then ioPRECOM.Text = .fields("PRECOM").Value
  If Not IsNull(.fields("IVA").Value) Then cbTIPOIVA.Text = .fields("IVA").Value
  
  
  End If
  
  
  

  
  
  End With

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

   On Error GoTo cbAgregar_Click_Error

With FrmMntArt
    .NumeroPedido = NumeroPedido
    .add_en_detalle = True
    Set .rc_detalle = rc
    .TemporadaWrk = frmPedProv.TemporadaWrk
    .Show
    DoEvents
End With

 Call carga_grid(rc.fields("CODART"))
  
  'Dim tmpcodigo As Variant
  
 'On error GoTo AddErr
  
  'With rc
    
   ' If Not (.BOF And .EOF) Then
      'mvBookMark = .Bookmark
     ' If .EditMode <> adEditNone Then .CancelUpdate
    'End If
    
       
    
    '.AddNew
    
   ' tmpcodigo = devuelve_campo("select max(LINEA) + 1 from DETPEDPRO where NUMERO = " & NumeroPedido)
    
   ' If tmpcodigo <> "@" Then
  '      .Fields("LINEA") = tmpcodigo
    '    .Fields("TEMPOR") = TemporadaActual
 '   End If
'
    'Set tmpcodigo = Nothing
   '
  '  lblstatus.Caption = "Agregar registro"
 '   mbAddNewFlag = True
    'SetButtons False
    'cbActualizar.Visible = True
    '
   ' ioDCTO1.SetFocus
  '
  'End With

   On Error GoTo 0
   Exit Sub

cbAgregar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cbAgregar_Click of Formulario frmDPedPro"
 
End Sub

Private Sub cbEliminar_Click()
    On Error GoTo DeleteErr
  With rc
    .fields("mbaja") = True
    .fields("FBAJA") = Date
    If .EOF Then .MoveLast
  End With
 
  Call cbactualizar_Click
  
Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub


Private Sub cbedicion_Click()
  On Error GoTo EditErr

          'si ya se ha introducido, no dejar modificar hasta que se haga un deshacer artículo.
   If rc.fields("METIDO") = True Then
    lblstatus.Caption = "El artículo ya se ha introducido en STOCK. Haga Deshacer Artículo para editar."
    Exit Sub
   End If
   
  mvBookMark = rc.Bookmark
   
  lblstatus.Caption = "Modificar registro"
  mbEditFlag = True
  SetButtons False
  cbActualizar.Visible = True
  
  ioUNIDADES.SetFocus
  
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
  cbAñadeTallaColor.Enabled = True
  
End Sub

Private Sub cbactualizar_Click()
  On Error GoTo UpdateErr
 
  'rc.Fields("FMODI").Value = Now 'poner fmodi para que no casque
  
 ' rc.Update 'Batch adAffectAll

'VALIDACIONES
If ioUNIDADES.Text = "" Then ioUNIDADES.Text = "0"
If cbTIPOIVA.Text = "" Then cbTIPOIVA.Text = "0"

With ioPRECOM

If .Text = "" Then
    lblstatus.Caption = "No se permite Precio de compra en blanco"
    .CancelarValidacion
    .SetFocus
    Exit Sub
ElseIf CDbl(.Valor) = 0 Then
    lblstatus.Caption = "No se permite Precio de compra a 0"
    .CancelarValidacion
    .SetFocus
    Exit Sub
End If

 If CDbl(.Valor) = 0 Then
        .SetFocus
        lblstatus.Caption = "Debe establecer un precio de COMPRA para el artículo"
        .CancelarValidacion
        .SetFocus
 End If
 
End With


With ioUNIDADES
If .Text = "0" Then
    lblstatus.Caption = "No se permite CERO unidades"
    .CancelarValidacion
    .SetFocus
    Exit Sub
End If
End With

With cbTALLA
If .Text = "" Then
    lblstatus.Caption = "No se permite TALLA en blanco"
     .SetFocus
    Exit Sub
End If
End With

With cbCOLOR
If .Text = "" Then
    lblstatus.Caption = "No se permite COLOR en blanco"
      .SetFocus
    Exit Sub
End If
End With

lblstatus.Caption = ""

rc.fields("IVA").Value = Replace(cbTIPOIVA.Text, ",", ".")
rc.fields("PRECOM").Value = CDbl(ioPRECOM.Text)

rc.UpdateBatch adAffectAll
'---------------------------------------------------------------------------------------
' Propósito   : Validación para precio de compra. Si cambia el precio de
'               compra en relación del anterior al entrar al campo, cambiar
'               los precios de compra de ese articulo en ese pedido
'---------------------------------------------------------------------------------------

  With ioPRECOM
   
        'si ha cambiado el precio de compra ...
        If tmp_cmprecom And (ioPRECOM.Text <> tmp_precom) Then
            
            If MsgBox("Se ha cambiado el Precio de Compra para este artículo. ¿Desea cambiar el precio para todas las tallas y colores de este artículo?", vbQuestion + vbYesNo) = vbYes Then
        
                
                locCnnSP.Execute "UPDATE DETPEDPRO SET PRECOM = " & Replace(ioPRECOM.Valor, ",", ".") & _
                " WHERE NUMERO = " & rc.fields("NUMERO") & " AND CODART = " & rc.fields("CODART") & " AND TEMPOR =" & rc.fields("TEMPOR")
                
                DoEvents
            
            End If
            tmp_precom = 0
            tmp_cmprecom = False
            
        End If
        
End With

'---------------------------------------------------------------------------------------
' Propósito   : Validación con funcionamiento similar a la del precio de compra
'---------------------------------------------------------------------------------------

  'si ha cambiado el iva ...
        If tmp_cmiva And (cbTIPOIVA.Text <> tmp_iva) And (cbTIPOIVA.Text <> "") Then
            
            If MsgBox("Se ha cambiado el IVA para este artículo. ¿Desea cambiar el IVA para todas las tallas y colores de este artículo?", vbQuestion + vbYesNo) = vbYes Then
        
                locCnnSP.Execute "UPDATE DETPEDPRO SET IVA = " & Replace(cbTIPOIVA.Text, ",", ".") & _
                " WHERE NUMERO = " & rc.fields("NUMERO") & " AND CODART = " & rc.fields("CODART") & " AND TEMPOR =" & rc.fields("TEMPOR")
                
                DoEvents
                           
            End If
            tmp_iva = 0
            tmp_cmiva = False
            
        End If



  DoEvents
  If mbAddNewFlag Then
    rc.MoveLast              'va al nuevo registro
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  
  lblstatus.Caption = ""
  
  Iva_Descontado = False
  
  Call carga_grid(rc.fields("CODART"))
   
  Call frmPedProv.refresca_grid_externo(True)
  

  
With cbTALLA
    Set .DataSource = Nothing
    .DataField = ""
    .DataField = "CODTALLA"
    Set .DataSource = rc
End With

With cbCOLOR
    Set .DataSource = Nothing
    .DataField = ""
    .DataField = "CODCOL"
    Set .DataSource = rc
End With

With ioUNIDADES
    Set .DataSource = Nothing
    .DataField = ""
    .DataField = "UNIDADES"
    Set .DataSource = rc
End With

  
  Call cmdLast_Click
  DoEvents
  Call carga_grid(rc.fields("CODART"))
  
  cbAñadeTallaColor.Enabled = True
  
  Exit Sub
UpdateErr:

  cbAñadeTallaColor.Enabled = True
  MsgBox Err.Description, vbInformation, "Atención"
End Sub

Private Sub cbcerrar_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  rc.MoveFirst
  mbDataChanged = False
  
   Call carga_grid(rc.fields("CODART"))

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

  Call carga_grid(rc.fields("CODART"))

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
  
   Call carga_grid(rc.fields("CODART"))

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
  
   Call carga_grid(rc.fields("CODART"))

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

  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal

  cbCATTALL.Locked = bVal
  cbTALLA.Locked = bVal
  cbCOLOR.Locked = bVal
    
End Sub

Private Sub vsTallasColores_dblClick()

    With vsTallasColores

    If .TextMatrix(.Row, 1) = "" Or Not IsNumeric(.TextMatrix(.Row, 1)) Then Exit Sub

        'ir al registro especificado
        rc.MoveFirst
        rc.Find "LINEA = " & .TextMatrix(.Row, 1), , adSearchForward
        
    End With
    
    DoEvents
    
    'editar el articulo seleccionado
    Call cbedicion_Click

End Sub
