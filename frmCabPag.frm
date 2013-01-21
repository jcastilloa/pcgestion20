VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmCabPag 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Deudas de Clientes ..."
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11655
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
   ScaleHeight     =   7050
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ckVerVentas 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ver todas las Ventas"
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
      Left            =   9435
      TabIndex        =   14
      Top             =   435
      Width           =   2100
   End
   Begin VB.CheckBox ckVerPagos 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ver todos los Pagos"
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
      Left            =   9435
      TabIndex        =   13
      Top             =   75
      Width           =   2100
   End
   Begin PCGestion.bsGradientLabel lblStatus 
      Height          =   345
      Left            =   60
      Top             =   5865
      Width           =   11580
      _ExtentX        =   20029
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
   Begin VSFlex8Ctl.VSFlexGrid fgArt 
      Height          =   4035
      Left            =   3705
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1800
      Width           =   4620
      _cx             =   8149
      _cy             =   7117
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
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCabPag.frx":0000
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
   Begin PCGestion.chameleonButton cbCancelar 
      Height          =   795
      Left            =   6915
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6240
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
      MICON           =   "frmCabPag.frx":00DE
      PICN            =   "frmCabPag.frx":00FA
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
      Height          =   795
      Left            =   5077
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6240
      Width           =   840
      _ExtentX        =   1482
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
      MICON           =   "frmCabPag.frx":09D4
      PICN            =   "frmCabPag.frx":09F0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VSFlex8Ctl.VSFlexGrid fgPag 
      Height          =   4035
      Left            =   8340
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1800
      Width           =   3315
      _cx             =   5847
      _cy             =   7117
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
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCabPag.frx":16CA
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
   Begin PCGestion.bsGradientLabel lblCliente 
      Height          =   345
      Left            =   1080
      Top             =   75
      Width           =   6420
      _ExtentX        =   11324
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
   Begin PCGestion.chameleonButton cbAceptar 
      Height          =   795
      Left            =   5962
      TabIndex        =   5
      Top             =   6240
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "&Aceptar"
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
      MICON           =   "frmCabPag.frx":17A8
      PICN            =   "frmCabPag.frx":17C4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel4 
      Height          =   315
      Left            =   7575
      Top             =   105
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      Caption         =   "-C- Asignar Cliente"
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
      Colour2         =   16761024
      CaptionAlignment=   1
   End
   Begin PCGestion.bsGradientLabel lblTotal 
      Height          =   375
      Left            =   1065
      Top             =   825
      Width           =   1965
      _ExtentX        =   3466
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
      Colour1         =   16558731
      Colour2         =   12632256
      CaptionAlignment=   1
   End
   Begin PCGestion.bsGradientLabel lblPagado 
      Height          =   375
      Left            =   4125
      Top             =   825
      Width           =   1965
      _ExtentX        =   3466
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
      Colour1         =   16558731
      Colour2         =   12632256
      CaptionAlignment=   1
   End
   Begin PCGestion.bsGradientLabel lblPendiente 
      Height          =   375
      Left            =   7395
      Top             =   825
      Width           =   1935
      _ExtentX        =   3413
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
      Colour1         =   16558731
      Colour2         =   12632256
      CaptionAlignment=   1
   End
   Begin VSFlex8Ctl.VSFlexGrid fgVen 
      Height          =   4035
      Left            =   30
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1800
      Width           =   3660
      _cx             =   6456
      _cy             =   7117
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
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCabPag.frx":249E
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
   Begin PCGestion.bsGradientLabel lblDependiente 
      Height          =   345
      Left            =   1065
      Top             =   435
      Width           =   6435
      _ExtentX        =   11351
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
   Begin PCGestion.bsGradientLabel bsGradientLabel2 
      Height          =   315
      Left            =   7575
      Top             =   450
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      Caption         =   "-D- Asignar Depend"
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
      Colour2         =   16761024
      CaptionAlignment=   1
   End
   Begin PCGestion.chameleonButton cbAñadirPago 
      Height          =   795
      Left            =   3772
      TabIndex        =   16
      Top             =   6240
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "-A-  &Añadir Pago"
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
      MICON           =   "frmCabPag.frx":257C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DEPEND."
      Height          =   270
      Left            =   30
      TabIndex        =   15
      Top             =   450
      Width           =   915
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pagos realizados"
      Height          =   315
      Left            =   8370
      TabIndex        =   12
      Top             =   1485
      Width           =   3255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Artículos Venta"
      Height          =   270
      Left            =   3750
      TabIndex        =   11
      Top             =   1485
      Width           =   4530
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ventas Pendientes de Cobro"
      Height          =   270
      Left            =   0
      TabIndex        =   10
      Top             =   1485
      Width           =   3690
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PENDIENTE"
      Height          =   270
      Left            =   6210
      TabIndex        =   8
      Top             =   870
      Width           =   1125
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PAGADO"
      Height          =   270
      Left            =   3060
      TabIndex        =   7
      Top             =   870
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      Height          =   270
      Left            =   180
      TabIndex        =   6
      Top             =   870
      Width           =   765
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENTE"
      Height          =   270
      Left            =   30
      TabIndex        =   4
      Top             =   90
      Width           =   915
   End
End
Attribute VB_Name = "frmCabPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Módulo       : frmDeudCli
' Fecha/Hora : 14/05/2004 11:28
' Autor         : JCastillo
' Propósito   :  Cobrar deudas de clientes
'---------------------------------------------------------------------------------------
Option Explicit

Public Codigo_Cliente As Long
Public Caja_Cliente As Long
'dependiente que acepta el cobro
Public ID_Usuario As Long

Dim Codigo_Pago As Long
Dim Caja_Pago As Long

Dim prime As Boolean

Dim Total_Deudas As Currency
Dim Total_Pagado As Currency

Dim rcpag As New ADODB.Recordset

Private Sub cbAñadirPago_Click()

   On Error GoTo cbAñadirPago_Click_Error

If (ID_Usuario = 0) Then
    lblstatus.Caption = "Debe seleccionar un dependiente"
    Exit Sub
End If

If (Caja_Cliente = 0) Or (Codigo_Cliente = 0) Then
    lblstatus.Caption = "Debe seleccionar un cliente"
    Exit Sub
End If

If (rcpag.State = 0) Then
    lblstatus.Caption = "Debe seleccionar una venta"
    Exit Sub
End If

lblstatus.Caption = ""

With frmDetPag
    Load frmDetPag
    DoEvents
    Set .rc = rcpag
    
    .Codigo_Pago = Codigo_Pago
    .codigo_caja = Caja_Pago
    .ID_Usuario = ID_Usuario
    
    '************************** AQUI *********************
    'ASIGNAR CODIGO DE PAGO Y CAJA DE PAGO MEDIANTE VARIABLES
    '_________________________________________________________
    
    .Show 1
End With

DoEvents

Call carga_totales_ventas_pendientes

   On Error GoTo 0
   Exit Sub

cbAñadirPago_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbAñadirPago_Click de Formulario frmDeudCli"

End Sub


Private Sub cbCancelar_Click()

Call Form_KeyDown(vbKeyEscape, 0)

End Sub



Private Sub fgPag_dblClick()

   On Error GoTo fgPag_Click_Error

If fgPag.Rows <= 1 Then Exit Sub
If rcpag.State = 0 Then Exit Sub

If Trim(fgPag.TextMatrix(fgPag.Row, 6)) = "" Then Exit Sub
If Not IsNumeric(fgPag.TextMatrix(fgPag.Row, 6)) Then Exit Sub

With frmDetPag
    Load frmDetPag
    DoEvents
    Set .rc = rcpag
    
    .Codigo_Pago = Codigo_Pago
    .codigo_caja = Caja_Pago
    .ID_Usuario = ID_Usuario
    .Linea_Pago = fgPag.TextMatrix(fgPag.Row, 6)
    
    '************************** AQUI *********************
    'ASIGNAR CODIGO DE PAGO Y CAJA DE PAGO MEDIANTE VARIABLES
    '_________________________________________________________
    
    .Show 1
End With

DoEvents

rcpag.Requery

Call carga_totales_ventas_pendientes

   On Error GoTo 0
   Exit Sub

fgPag_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento fgPag_Click de Formulario frmCabPag"

End Sub

Private Sub fgVen_Click()
Dim tmpcod As Long
Dim tmpcaja As Byte
    
   On Error GoTo fgVen_Click_Error

    With fgVen
                  
        If .Rows <= 1 Then Exit Sub
        
        'obtener el numero de ticket y descomponer para sacar caja
        'y venta
        If .TextMatrix(.Row, 4) <> "" And IsNumeric(.TextMatrix(.Row, 4)) Then
        
            tmpcaja = CByte(Right(.TextMatrix(.Row, 4), 3))
            tmpcod = CLng(Left(.TextMatrix(.Row, 4), Len(.TextMatrix(.Row, 4)) - 3))
            Codigo_Pago = .TextMatrix(.Row, 5)
            Caja_Pago = tmpcaja
            
            Call carga_grid_ventas(tmpcod, tmpcaja)
            DoEvents
            Call carga_grid_arreglos(tmpcod, tmpcaja)
            DoEvents
            
            Call carga_grid_pagos(Codigo_Pago, tmpcaja, 0, 0, True)
                            
        End If
    
    End With

   On Error GoTo 0
   Exit Sub

fgVen_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento fgVen_Click de Formulario frmCabPag"

End Sub


'---------------------------------------------------------------------------------------
' Subrutina   : carga_grid_ventas
' Fecha/Hora  : 16/05/2004 10:16
' Autor       : JCASTILLO
' Propósito   : Carga el grid con los registros de la venta seleccionada
'---------------------------------------------------------------------------------------
Private Sub carga_grid_ventas(CODIGO_VENTA As Long, codigo_caja As Byte)
Dim tmpart  As Variant
Dim tmpcodcolor  As Variant
Dim rcdet As New ADODB.Recordset
Dim totimpor As Currency
Dim tmpiva As Currency
Dim tmpdcto As Currency

   On Error GoTo carga_grid_ventas_Error

   With fgArt
        .Clear
        .Cols = 14
        .Rows = 2
        .ColHidden(1) = True
        '.ColHidden(13) = True
        .ColFormat(9) = "Currency"
        .ColFormat(11) = "Currency"
        .ColAlignment(3) = flexAlignCenterCenter
        .TextMatrix(0, 2) = "Prov."
        .TextMatrix(0, 3) = "Ref."
        .TextMatrix(0, 4) = "Modelo"
        .TextMatrix(0, 5) = "Temp."
        .TextMatrix(0, 6) = "Talla"
        .TextMatrix(0, 7) = "Color"
        .TextMatrix(0, 8) = "Uds."
        .TextMatrix(0, 9) = "PVP"
        .TextMatrix(0, 10) = "Dcto"
        .TextMatrix(0, 11) = "Total"
        .TextMatrix(0, 12) = "CBarras"
    
      'abrir detalle
     If rcdet.State = 1 Then rcdet.Close
     rcdet.Open "SELECT CODART, TEMPOR, CODTALLA, CODCOL, UNIDADES, PREVEN, DCTO, IVA FROM DETVENTA WHERE CODVEN = " & CODIGO_VENTA & " AND CODCAJA = " & codigo_caja & " ORDER BY LINEA", locCnn, adOpenStatic, adLockOptimistic

    Do Until rcdet.EOF

        'articulo no existe en la transferencia.
        tmpart = devuelve_matriz("SELECT CODPROV, PREVEN, REF FROM MAARTIC WHERE CODIGO = " & rcdet.fields("CODART") & " AND TEMPOR = " & rcdet.fields("TEMPOR"), locCnn)
      
        'articulo OK, existe.
        .AddItem "", 1
        .TextMatrix(1, 1) = 0
        .TextMatrix(1, 2) = Trim(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & tmpart(0), locCnn))
        .TextMatrix(1, 3) = Trim(tmpart(2))
        
        .TextMatrix(1, 4) = rcdet.fields("CODART") & " " & Trim(devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & rcdet.fields("CODART") & " AND TEMPOR = " & rcdet.fields("TEMPOR"), locCnn))
        .TextMatrix(1, 5) = Trim(devuelve_campo("SELECT ABREVIA FROM TEMPOR WHERE IDTEM = " & rcdet.fields("TEMPOR"), locCnn))
        
        If rcdet.fields("CODTALLA") > 0 Then
            .TextMatrix(1, 6) = Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & rcdet.fields("CODTALLA"), locCnn))
        End If
        
        If rcdet.fields("CODCOL") > 0 Then
            .TextMatrix(1, 7) = Trim(devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & rcdet.fields("CODCOL"), locCnn))
            tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & rcdet.fields("CODCOL"), locCnn)
            .Row = 1
            .Col = 7
            .CellBackColor = tmpcodcolor
            .Col = 3
        End If
        
        'sacar subtotal
        totimpor = (rcdet.fields("UNIDADES") * rcdet.fields("PREVEN"))
        'sacar el importe del dcto
        tmpdcto = (totimpor * rcdet.fields("DCTO")) / 100
        'sacar el importe del iva
        tmpiva = ((totimpor - tmpdcto) * rcdet.fields("IVA")) / 100
        
    
        .TextMatrix(1, 8) = rcdet.fields("UNIDADES")
        .TextMatrix(1, 9) = rcdet.fields("PREVEN")
        .TextMatrix(1, 10) = rcdet.fields("DCTO")
        .TextMatrix(1, 11) = totimpor - tmpdcto + tmpiva
        
        .TextMatrix(1, 12) = Conforma_CB(rcdet.fields("CODART"), rcdet.fields("TEMPOR"), rcdet.fields("CODTALLA"), rcdet.fields("CODCOL"))
        
       ' .TextMatrix(1, 12) = rcdet.fields("CODIGO")
        
        rcdet.MoveNext
        
        totimpor = 0
        tmpdcto = 0
        tmpiva = 0
    
    Loop
            
        .subtotal flexSTSum, , 8, , vbBlue, vbWhite
        .subtotal flexSTSum, , 11, , vbBlue, vbWhite
        .TextMatrix(1, 7) = "Uds: "
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 10) = "Total:"
        .AutoSize 1, .Cols - 1
        
    End With

    rcdet.Close
    Set rcdet = Nothing


   On Error GoTo 0
   Exit Sub

carga_grid_ventas_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_grid_ventas de Formulario frmDeudCli"
End Sub


'---------------------------------------------------------------------------------------
' Subrutina   : carga_grid_arreglos
' Fecha/Hora  : 16/05/2004 13:15
' Autor       : JCASTILLO
' Propósito   : Carga los arreglos correspondientes a la venta
'---------------------------------------------------------------------------------------
Private Sub carga_grid_arreglos(CODIGO_VENTA As Long, codigo_caja As Byte)
Dim tmpart  As Variant
Dim tmpcodcolor  As Variant
Dim rcdet As New ADODB.Recordset
  
   On Error GoTo carga_grid_arreglos_Error

   With fgArt
   
'        .Clear
'        .Cols = 13
'        .Rows = 2
'        .ColHidden(1) = True
'        .ColHidden(12) = True
'        .ColFormat(9) = "Currency"
'        .ColFormat(11) = "Currency"
'        .ColAlignment(3) = flexAlignCenterCenter
'        .TextMatrix(0, 2) = "Prov."
'        .TextMatrix(0, 3) = "Ref."
'        .TextMatrix(0, 4) = "Modelo"
'        .TextMatrix(0, 5) = "Temp."
'        .TextMatrix(0, 6) = "Talla"
'        .TextMatrix(0, 7) = "Color"
'        .TextMatrix(0, 8) = "Uds."
'        .TextMatrix(0, 9) = "PVP"
'        .TextMatrix(0, 10) = "Dcto"
'        .TextMatrix(0, 11) = "Total"
'
    
      'abrir detalle
     If rcdet.State = 1 Then rcdet.Close
     rcdet.Open "SELECT ID, CODART, TEMPOR, CODTALLA, CODCOL, PVP FROM ARREGLOS WHERE CODVEN = " & CODIGO_VENTA & " AND CODCAJ = " & codigo_caja & " ORDER BY ID", locCnn, adOpenStatic, adLockOptimistic

    Do Until rcdet.EOF
        
        .AddItem "", 1
        .TextMatrix(1, 1) = 0
            
        'si se ha asignado un artículo al arreglo
        If rcdet.fields("CODART") > 0 And rcdet.fields("TEMPOR") > 0 Then
            tmpart = devuelve_matriz("SELECT CODPROV, PREVEN, REF FROM MAARTIC WHERE CODIGO = " & rcdet.fields("CODART") & " AND TEMPOR = " & rcdet.fields("TEMPOR"), locCnn)
              
            'articulo OK, existe.

            .TextMatrix(1, 2) = Trim(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & tmpart(0), locCnn))
            .TextMatrix(1, 3) = Trim(tmpart(2))
        
            .TextMatrix(1, 4) = rcdet.fields("CODART") & " " & Trim(devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & rcdet.fields("CODART") & " AND TEMPOR = " & rcdet.fields("TEMPOR"), locCnn))
            .TextMatrix(1, 5) = Trim(devuelve_campo("SELECT ABREVIA FROM TEMPOR WHERE IDTEM = " & rcdet.fields("TEMPOR"), locCnn))
        
            If rcdet.fields("CODTALLA") > 0 Then
                .TextMatrix(1, 6) = Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & rcdet.fields("CODTALLA"), locCnn))
            End If
        
            If rcdet.fields("CODCOL") > 0 Then
                .TextMatrix(1, 7) = Trim(devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & rcdet.fields("CODCOL"), locCnn))
                tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & rcdet.fields("CODCOL"), locCnn)
                .Row = 1
                .Col = 7
                .CellBackColor = tmpcodcolor
                .Col = 3
            End If
        
        Else
        
        .TextMatrix(1, 2) = "Arreglo Varios"
        .TextMatrix(1, 4) = "Arreglo Varios"
        
        End If
        
        .TextMatrix(1, 9) = rcdet.fields("PVP")
        .TextMatrix(1, 11) = rcdet.fields("PVP")
        
       ' .TextMatrix(1, 12) = rcdet.fields("CODIGO")
        
        rcdet.MoveNext
           
    Loop
            
        .subtotal flexSTSum, , 8, , vbBlue, vbWhite
        .subtotal flexSTSum, , 11, , vbBlue, vbWhite
        .TextMatrix(1, 7) = "Uds: "
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 10) = "Total:"
        .AutoSize 1, .Cols - 1
        
    End With

    rcdet.Close
    Set rcdet = Nothing


   On Error GoTo 0
   Exit Sub

carga_grid_arreglos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_grid_arreglos de Formulario frmDeudCli"

 End Sub

Private Sub Asigna_Dependiente()
            
   On Error GoTo Asigna_Dependiente_Error

            With frmSelDep
                .Show 1
                ID_Usuario = .ID_Dependiente
            End With
            DoEvents
            Set frmSelDep = Nothing
            DoEvents
            
            With locCnn
                If .State = 0 Then
                .CursorLocation = adUseClient
                .Open strLocCnn
                End If
            End With
            
            lblDependiente.Caption = Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & ID_Usuario, locCnn))

   On Error GoTo 0
   Exit Sub

Asigna_Dependiente_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Asigna_Dependiente de Formulario frmDeudCli"
            
End Sub


Private Sub Form_Activate()

  If Caja_Cliente = 0 And Codigo_Cliente = 0 Then prime = False
  
  If prime Then Exit Sub

  'si esta a 0 mostrar la pantalla de seleccion de dependiente
  Do
    If ID_Usuario = 0 Then
  
        Call Asigna_Dependiente
      
    Else
        
        lblDependiente.Caption = Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & ID_Usuario, locCnn))
    
    End If
  Loop Until ID_Usuario > 0

  'si estan a 0 mostrar el grid de clientes
  If Caja_Cliente = 0 And Codigo_Cliente = 0 Then
          
        Call Abre_Grid_Clientes
      
  'si vienen con datos de algun otro formulario, mostrar el nombre directamente
  ElseIf Caja_Cliente > 0 And Codigo_Cliente > 0 Then
   lblCliente.Caption = devuelve_campo("SELECT RAZO FROM CLIENTES WHERE CODIGO = " & Codigo_Cliente & " AND CODCAJA = " & Caja_Cliente, locCnn)
    'cargar ventas para el cliente
    Call carga_totales_ventas_pendientes
  End If
  
  prime = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
      
   Select Case KeyCode
   
      'añadir / editar pagos ...
      Case vbKeyA
   
        Call cbAñadirPago_Click
        KeyCode = 0
   
      'Asignar Cliente ...
      Case vbKeyC

       'abre el grid de los clientes
       Call Abre_Grid_Clientes
       KeyCode = 0
            
       'selecciona dependiente
      Case vbKeyD
        
       Call Asigna_Dependiente
       KeyCode = 0
      
      Case vbKeyEscape
      
        If MsgBox("¿Desea cerrar la ventana actual?", vbQuestion + vbYesNo, titulo) = vbYes Then Unload Me
    
   End Select
   
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
            
    .desde_pagos = True
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

'Asigna el cliente seleccionado en el flexgrid, para llamar desde el flexclientes
Public Sub Asignar_cliente_flex(CodigoCliente As Long, codcaja As Byte)

With frmFlexCli
    
    If .seleccionado Then
    
        'asignar valores ...
        Codigo_Cliente = CodigoCliente 'rccli.Fields("CODIGO")
        Caja_Cliente = codcaja 'rccli.Fields("CODCAJA")
        lblCliente.Caption = devuelve_campo("SELECT RAZO FROM CLIENTES WHERE CODIGO = " & Codigo_Cliente & " AND CODCAJA = " & Caja_Cliente, locCnn)
        
        'cargar ventas para el cliente
        Call carga_totales_ventas_pendientes
    
    'dejar como estaba
    'Else
    
      '  rc.Fields("CODCLI") = Null
      '  rc.Fields("CAJACLI") = Null
      '  lblCliente.Caption = ""
        
    End If
    
End With
    
     '   rc.Update
        
End Sub



'---------------------------------------------------------------------------------------
' Procedimiento : carga_totales_ventas_pendientes
' Fecha/Hora     : 14/05/2004 11:05
' Autor             : JCastillo
' Propósito       : Carga una lista-resumen de las ventas pendientes en el listbox
'---------------------------------------------------------------------------------------
Private Sub carga_totales_ventas_pendientes()
Dim rc As New ADODB.Recordset

   On Error GoTo carga_totales_ventas_pendientes_Error

    With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
    End With
  
    Total_Deudas = 0
    Total_Pagado = 0
  
    rc.Open "SELECT FMODI, CODIGO, CODVEN, CODCAJA, IMPORTE, PAGADO FROM CABDEUDCLI WHERE CODCLI = " & Codigo_Cliente & " AND CAJACLI = " & Caja_Cliente & " AND ESTADO IN (0, 1) ORDER BY FMODI DESC", locCnn, adOpenStatic, adLockReadOnly
    
    'Poner titulos antes de nada
    With fgVen
        .Clear
        .Rows = 1
        .Cols = 6
        
        .ColHidden(5) = True
        .ColFormat(2) = "Currency"
        .ColFormat(3) = "Currency"
              
        .TextMatrix(0, 1) = "FECHA"
        .TextMatrix(0, 2) = "IMPORTE"
        .TextMatrix(0, 3) = "PAGADO"
        .TextMatrix(0, 4) = "TICKET"
    
    If rc.RecordCount < 0 Then Exit Sub
    
    Do Until rc.EOF
    
        fgVen.Rows = fgVen.Rows + 1
        'fecha
        .TextMatrix(.Rows - 1, 1) = Format(rc.fields("FMODI"), "dd/mm/yyyy")
        'importe total de la deuda
        .TextMatrix(.Rows - 1, 2) = rc.fields("IMPORTE")
        'importe que ya se ha pagado
        
        .TextMatrix(.Rows - 1, 3) = rc.fields("PAGADO")
        'nº de ticket
        .TextMatrix(.Rows - 1, 4) = CStr(rc.fields("CODVEN")) & Format(rc.fields("CODCAJA"), "000")
        'codigo del pago
        .TextMatrix(.Rows - 1, 5) = rc.fields("CODIGO")
        
        Total_Deudas = Total_Deudas + rc.fields("IMPORTE")
        Total_Pagado = Total_Pagado + rc.fields("PAGADO")
        
        rc.MoveNext
        
        'lstVentasPen.AddItem Format(rc.fields("FMODI"), "dd/mm/yyyy") & " - " & ". Ticket: " & CStr(rc.fields("CODVEN")) & Format(Caja_Cliente, "000") & ". " & Format(rc.fields("IMPORTE"), "Currency")
    Loop
                
        If .Rows > 1 Then
            .SubtotalPosition = flexSTAbove
            .subtotal flexSTCount, , 4, , vbBlue, vbWhite
            .subtotal flexSTSum, , 3, , vbBlue, vbWhite
            .subtotal flexSTSum, , 2, , vbBlue, vbWhite
            .TextMatrix(1, 4) = "Nº (" & .TextMatrix(1, 4) & ")"
            .TextMatrix(1, 1) = ""
        End If
        
        'si ha devuelto algun resultado:
        If .Rows > 2 Then
        
            .Select 2, 1, 2, .Cols - 1
            Call fgVen_Click
        
        End If
        
    End With
        
    lblTotal.Caption = Format(Total_Deudas, "Currency")
    lblPagado.Caption = Format(Total_Pagado, "Currency")
    lblPendiente.Caption = Format(Total_Deudas - Total_Pagado, "Currency")
        
    rc.Close
    Set rc = Nothing

   On Error GoTo 0
   Exit Sub

carga_totales_ventas_pendientes_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_totales_ventas_pendientes de Formulario frmDeudCli"

End Sub


'---------------------------------------------------------------------------------------
' Subrutina   : carga_grid_pagos
' Fecha/Hora  : 16/05/2004 13:31
' Autor       : JCASTILLO
' Propósito   : Cargar los pagos.
'               - Si mostrar_todos = false, se muestran solo los pagos de una determinada
'               venta, para lo q se necesita codigo_deuda y caja_deuda.
'               - Si mostrar_todos = true, se muestran todos los pagos para un determinado
'               ciente, para lo que se necesita codigo_cliente y caja_cliente.
'               - Los campos que no se necesiten en cada momento deben ir a 0.
'---------------------------------------------------------------------------------------
Private Sub carga_grid_pagos(codigo_deuda As Long, caja_deuda As Byte, Codigo_Cliente As Long, Caja_Cliente As Byte, mostrar_todos As Boolean)
   
   On Error GoTo carga_grid_pagos_Error

    With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
    End With

   'Poner titulos antes de nada
    With fgPag
        .Clear
        .Rows = 1
        .Cols = 7
        '.ColHidden(5) = True
        .ColHidden(6) = True
        .ColFormat(2) = "Currency"
        .TextMatrix(0, 1) = "FECHA"
        .TextMatrix(0, 2) = "IMPORTE"
        .TextMatrix(0, 3) = "COMEN."
        .TextMatrix(0, 4) = "FACTURA"
        .TextMatrix(0, 5) = "USUARIO"
       
         
     
    If rcpag.State = 1 Then rcpag.Close
    
    If mostrar_todos = False Then
        rcpag.Open "SELECT CODIGO, CODCAJA, LINEA, CODPER, IMPORTE, IMPORTE_CON, FACTURA, FMODI, DESCRIPCION, MBAJA FROM DETDEUDCLI WHERE MBAJA = 0 AND CODIGO = " & codigo_deuda & " AND CODCAJA = " & caja_deuda, locCnn, adOpenDynamic, adLockOptimistic
    Else
        rcpag.Open "SELECT CODIGO, CODCAJA, LINEA, CODPER, IMPORTE, IMPORTE_CON, FACTURA, FMODI, DESCRIPCION, MBAJA FROM DETDEUDCLI WHERE MBAJA = 0 AND CODCAJA = " & caja_deuda, locCnn, adOpenDynamic, adLockOptimistic
    End If
    
    If rcpag.RecordCount < 0 Then Exit Sub
    
    Do Until rcpag.EOF
    
        .Rows = .Rows + 1
        'fecha
        .TextMatrix(.Rows - 1, 1) = Format(rcpag.fields("FMODI"), "dd/mm/yyyy")
        'importe
        .TextMatrix(.Rows - 1, 2) = rcpag.fields("IMPORTE")
        'descripcion del pago
        If Not IsNull(rcpag.fields("DESCRIPCION")) Then .TextMatrix(.Rows - 1, 3) = Trim(rcpag.fields("DESCRIPCION"))
        'Factura
        .TextMatrix(.Rows - 1, 4) = rcpag.fields("FACTURA")
        'Usuario
        .TextMatrix(.Rows - 1, 5) = rcpag.fields("CODPER")
        
        .TextMatrix(.Rows - 1, 6) = rcpag.fields("LINEA")
        
        rcpag.MoveNext
     
     Loop
                
        If .Rows > 1 Then
            .SubtotalPosition = flexSTAbove
            .subtotal flexSTCount, , 4, , vbBlue, vbWhite
            .subtotal flexSTSum, , 2, , vbBlue, vbWhite
            .TextMatrix(1, 4) = "Nº (" & .TextMatrix(1, 4) & ")"
            .TextMatrix(1, 1) = ""
        End If
        
    End With
    
   On Error GoTo 0
   Exit Sub

carga_grid_pagos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_grid_pagos de Formulario frmDeudCli"
End Sub


Private Sub Form_Load()

  With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
  End With
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   On Error GoTo Form_QueryUnload_Error

    If rcpag.State = 1 Then rcpag.Close
    Set rcpag = Nothing
    Set frmCabPag = Nothing

   On Error GoTo 0
   Exit Sub

Form_QueryUnload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Form_QueryUnload de Formulario frmDeudCli"

End Sub


