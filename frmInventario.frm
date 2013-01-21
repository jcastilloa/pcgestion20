VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmInventario 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventario ..."
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11670
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   11670
   Begin PCGestion.chameleonButton cbAceptarINV 
      Height          =   810
      Left            =   9225
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6435
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   1429
      BTYPE           =   3
      TX              =   "&ACEPTAR INVENTARIO"
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
      MICON           =   "frmInventario.frx":0000
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
      Height          =   810
      Left            =   10710
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6435
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1429
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
      MICON           =   "frmInventario.frx":001C
      PICN            =   "frmInventario.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblTotal_B 
      Height          =   465
      Left            =   6585
      Top             =   480
      Width           =   1785
      _ExtentX        =   3149
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
      CaptionColour   =   16711680
      Colour1         =   14737632
      Colour2         =   12632256
      CaptionAlignment=   2
   End
   Begin PCGestion.bsGradientLabel lblTotal_A 
      Height          =   465
      Left            =   6585
      Top             =   0
      Width           =   1785
      _ExtentX        =   3149
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
      CaptionColour   =   16711680
      Colour1         =   14737632
      Colour2         =   12632256
      CaptionAlignment=   2
   End
   Begin PCGestion.miText ioCODBAR 
      Height          =   525
      Left            =   1035
      TabIndex        =   3
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
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
   Begin PCGestion.bsGradientLabel lblTotal_General 
      Height          =   465
      Left            =   9855
      Top             =   480
      Width           =   1785
      _ExtentX        =   3149
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
      CaptionColour   =   16711680
      Colour1         =   14737632
      Colour2         =   12632256
      CaptionAlignment=   2
   End
   Begin PCGestion.miText ioPERCHERO 
      Height          =   525
      Left            =   1035
      TabIndex        =   0
      Top             =   15
      Width           =   990
      _ExtentX        =   1746
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
   Begin PCGestion.miText ioESTANTE 
      Height          =   525
      Left            =   2805
      TabIndex        =   1
      Top             =   15
      Width           =   1005
      _ExtentX        =   1773
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
   Begin PCGestion.miText ioCASILLA 
      Height          =   525
      Left            =   4575
      TabIndex        =   2
      Top             =   0
      Width           =   1005
      _ExtentX        =   1773
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
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   5055
      Left            =   15
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   11640
      _cx             =   20532
      _cy             =   8916
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
      FormatString    =   $"frmInventario.frx":0912
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
   Begin PCGestion.miText iovPerchero 
      Height          =   525
      Left            =   810
      TabIndex        =   14
      Top             =   6390
      Width           =   990
      _ExtentX        =   1746
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
   Begin PCGestion.miText iovEstante 
      Height          =   525
      Left            =   2400
      TabIndex        =   15
      Top             =   6840
      Width           =   1005
      _ExtentX        =   1773
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
   Begin PCGestion.miText iovCasilla 
      Height          =   525
      Left            =   2400
      TabIndex        =   16
      Top             =   6390
      Width           =   1005
      _ExtentX        =   1773
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
   Begin PCGestion.chameleonButton cmBorrarSel 
      Height          =   345
      Left            =   3450
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   609
      BTYPE           =   9
      TX              =   "Con&sultar"
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
      MICON           =   "frmInventario.frx":09F0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton chameleonButton1 
      Height          =   345
      Left            =   5190
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   609
      BTYPE           =   9
      TX              =   "&Borrar Selección"
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
      MICON           =   "frmInventario.frx":0A0C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cmBorrarUltimo 
      Height          =   345
      Left            =   5190
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   609
      BTYPE           =   9
      TX              =   "Borrar &Último"
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
      MICON           =   "frmInventario.frx":0A28
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblstatus 
      Height          =   360
      Left            =   30
      Top             =   6030
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   635
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
   Begin PCGestion.chameleonButton cmAbrirInventario 
      Height          =   345
      Left            =   3450
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   609
      BTYPE           =   9
      TX              =   "Abrir &Inventario"
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
      MICON           =   "frmInventario.frx":0A44
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.OptionButton optReemplazar 
      Height          =   375
      Left            =   7020
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6855
      Width           =   2160
      BackColor       =   12632256
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "3810;661"
      Value           =   "0"
      Caption         =   "Reemplazar Inventario"
      FontName        =   "Trebuchet MS"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Perchero"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   -30
      TabIndex        =   19
      Top             =   6465
      Width           =   840
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Estante"
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
      Left            =   1590
      TabIndex        =   18
      Top             =   6900
      Width           =   780
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Casilla"
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
      Left            =   1695
      TabIndex        =   17
      Top             =   6465
      Width           =   660
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Casilla"
      Height          =   315
      Left            =   3780
      TabIndex        =   11
      Top             =   90
      Width           =   780
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Estante"
      Height          =   315
      Left            =   2010
      TabIndex        =   10
      Top             =   75
      Width           =   780
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Perchero"
      Height          =   300
      Left            =   0
      TabIndex        =   9
      Top             =   60
      Width           =   990
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total General"
      Height          =   315
      Left            =   8400
      TabIndex        =   8
      Top             =   555
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total B"
      Height          =   315
      Left            =   5730
      TabIndex        =   7
      Top             =   585
      Width           =   765
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "C. Barras"
      Height          =   300
      Left            =   60
      TabIndex        =   6
      Top             =   570
      Width           =   945
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total A"
      Height          =   345
      Left            =   5760
      TabIndex        =   5
      Top             =   60
      Width           =   780
   End
   Begin MSForms.OptionButton optAñadir 
      Height          =   375
      Left            =   7005
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6435
      Width           =   1890
      BackColor       =   12632256
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "3334;661"
      Value           =   "1"
      Caption         =   "Añadir a Inventario"
      FontName        =   "Trebuchet MS"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Modulo      : frmInventario
' Fecha/Hora  : 06/03/2004 21:20
' Autor       : JCASTILLO
' Propósito   : Formulario de entrada para los datos del inventario
'---------------------------------------------------------------------------------------
Dim miCod As MiCodBar
Dim InvCnn As ADODB.Connection
Dim rc As ADODB.Recordset

Dim conta_lineas As Long

Dim tmp_precom As Currency  'para alm. la suma total de precios de compra
Dim tmp_totalA As Currency  'para almacenar el total en A
Dim tmp_totalB As Currency  'para almacenar el total en B

'para controlar si se debe crear una nueva base de datos de inventarios
Dim creado_o_cargado As Boolean

Private Sub añade_linea_grid(gCodart As Long, gTempor As Long, gCodtalla As Integer, gCodCol As Integer, gPerchero As Long, gEstante As Long, gCasilla As Long, gID As Long)
Dim t_articulo As Variant
Dim t_tempor As String
Dim t_talla As String
Dim t_color As Variant
Dim var As Byte
Dim stren As String

'obtener datos ...
t_articulo = devuelve_matriz("SELECT MODELO, CODPROV, REF, PRECOM, IVACOM from MAARTIC where CODIGO = " & gCodart & " and TEMPOR = " & gTempor, locCnn)
t_tempor = Trim(devuelve_campo("SELECT ABREVIA FROM TEMPOR WHERE IDTEM = " & gTempor, locCnn))
t_talla = Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & gCodtalla, locCnn))
t_color = devuelve_matriz("SELECT DESCRIPCION, CODCOL FROM COLORES WHERE CODIGO = " & gCodCol, locCnn)

  
    'si tiene IVA, sumar a importe en A
    If t_articulo(4) > 0 Then
     tmp_totalA = tmp_totalA + t_articulo(3)
     stren = "A"
    'si no tiene IVA sumar a importe en B
    Else
    tmp_totalB = tmp_totalB + t_articulo(3)
    stren = "B"
    End If
    

With fg
    .AddItem "", 2
    .TextMatrix(2, 1) = gID
    .TextMatrix(2, 2) = Trim(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & t_articulo(1), locCnn))
    .TextMatrix(2, 3) = Trim(t_articulo(2))
    .TextMatrix(2, 4) = Format(gCodart, "00000") & " " & Trim(t_articulo(0))
    .TextMatrix(2, 5) = t_tempor
    .TextMatrix(2, 6) = t_talla
    .TextMatrix(2, 7) = Trim(t_color(0))
    .TextMatrix(2, 8) = Format(t_articulo(3), "Currency")
    .TextMatrix(2, 9) = stren
    .TextMatrix(2, 10) = gPerchero
    .TextMatrix(2, 11) = gEstante
    .TextMatrix(2, 12) = gCasilla
    .TextMatrix(2, 13) = Now
  
    tmp_precom = tmp_precom + t_articulo(3)
    conta_lineas = conta_lineas + 1
    
    lblTotal_General.Caption = Format(tmp_precom, "Currency")
    lblTotal_A.Caption = Format(tmp_totalA, "Currency")
    lblTotal_B.Caption = Format(tmp_totalB, "Currency")
    
    'poner el color
        
    .Row = 1
    
    For var = 2 To .Cols - 1
        .Col = var
        .CellBackColor = vbBlue
        .CellForeColor = vbWhite
        .CellFontBold = True
    Next var
        
    .Col = 7
    .Row = 2
    .CellBackColor = t_color(1)
    .Col = 2
    
    .TextMatrix(1, 4) = "Unidades: " & conta_lineas
    .TextMatrix(1, 8) = "Total: " & Format(tmp_precom, "Currency")
    .TextMatrix(1, 2) = "Totales"
    .AutoSize 2, .Cols - 1
            
End With

End Sub

Private Sub cbAceptarINV_Click()
Dim tmpop As String

   On Error GoTo cbAceptarINV_Click_Error

If optAñadir.Value = True Then
    tmpop = "Añadir a stock actual"
Else
    tmpop = "Reemplazar stock actual"
End If

If MsgBox("¿Desea aceptar el inventario actual?, con la opción de: " & tmpop, vbYesNo, titulo) = vbNo Then
    lblstatus.Caption = "Se ha cancelado la actualización de stock"
    Exit Sub
End If

lblstatus.Caption = "Aceptando Inventario ..."
DoEvents

  With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
  End With

If optAñadir.Value Then
  
'si es opcion reemplazar, borrar primero los datos existentes
Else

    'borrar todo para el almacén actual ...
    locCnn.Execute "DELETE FROM STOCK WHERE CODALM = " & AlmacenActual

End If

If Not rc.BOF Then rc.MoveFirst

Do Until rc.EOF

    Call stock(rc.fields("CODART"), rc.fields("TEMPOR"), rc.fields("CODTALLA"), rc.fields("CODCOL"), AlmacenActual, 1, True, locCnn)
    rc.MoveNext
    
Loop

MsgBox "El inventario actual se ha aceptado correctamente.", vbInformation, titulo

lblstatus.Caption = "El inventario actual se ha aceptado correctamente."

   On Error GoTo 0
   Exit Sub

cbAceptarINV_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbAceptarINV_Click de Formulario frmInventario"

End Sub

Private Sub cbCancelar_Click()



   On Error GoTo cbCancelar_Click_Error

Unload Me        ' Inventario ...

   On Error GoTo 0
   Exit Sub

cbCancelar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbCancelar_Click de Formulario frmInventario"

End Sub



Private Sub cmAbrirInventario_Click()

   On Error GoTo cmAbrirInventario_Click_Error

With FrmInicio.Dialogo
     .InitDir = "c:\INVENTARIOS"
     .DialogTitle = "Abrir inventario ..."
     .Filter = "Inventarios (*.inv)|*.inv|"
     .ShowOpen
     .Filter = ""
End With
    
If (FrmInicio.Dialogo.CancelError = True) Or (Trim(FrmInicio.Dialogo.filename = "")) Then Exit Sub
   
If creado_o_cargado = False Then
    Set InvCnn = New ADODB.Connection
    Set rc = New ADODB.Recordset
Else
    rc.Close
    InvCnn.Close
End If

InvCnn.Open strCnnMdb & FrmInicio.Dialogo.filename

rc.Open "SELECT * FROM INVENTARIO", InvCnn, adOpenStatic, adLockOptimistic

Call carga_grid

creado_o_cargado = True

   On Error GoTo 0
   Exit Sub

cmAbrirInventario_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmAbrirInventario_Click de Formulario frmInventario"

End Sub

Private Sub cmBorrarUltimo_Click()

   On Error GoTo cmBorrarUltimo_Click_Error

rc.Requery
If rc.RecordCount <= 0 Then Exit Sub

If Not rc.EOF Then rc.MoveLast

rc.Delete
DoEvents

'descontar a los totales:

Select Case fg.TextMatrix(2, 9)

    Case ""
    
    Case "A"
       'restar a total A
    
        tmp_precom = tmp_precom - fg.TextMatrix(2, 8)
        tmp_totalA = tmp_totalA - fg.TextMatrix(2, 8)
               
        
    Case "B"
        'restar a total B
        tmp_precom = tmp_precom - fg.TextMatrix(2, 8)
        tmp_totalB = tmp_totalB - fg.TextMatrix(2, 8)

End Select


conta_lineas = conta_lineas - 1

lblTotal_General.Caption = Format(tmp_precom, "Currency")
lblTotal_A.Caption = Format(tmp_totalA, "Currency")
lblTotal_B.Caption = Format(tmp_totalB, "Currency")

If fg.Rows > 1 Then fg.RemoveItem 2

    fg.TextMatrix(1, 4) = "Unidades: " & conta_lineas
    fg.TextMatrix(1, 8) = "Total: " & Format(tmp_precom, "Currency")
    fg.TextMatrix(1, 2) = "Totales"
    fg.AutoSize 2, fg.Cols - 1


   On Error GoTo 0
   Exit Sub

cmBorrarUltimo_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmBorrarUltimo_Click de Formulario frmInventario"

End Sub

Private Sub Form_Load()
Dim tmpruta As String

  If Dir("C:\INVENTARIOS\", vbDirectory) = "" Then
    MkDir "C:\INVENTARIOS"
  End If

  Move (Screen.Width - Width) \ 2, Separacion_MDIForm
  
  With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
  End With
  
With ioCODBAR
    '.SoloNumeros = True
    .Alineacion = 1
    .LongMaxima = LenCodBar
End With

With ioPERCHERO
    .SoloNumeros = True
    .Alineacion = 1
    .LongMaxima = 8
End With

With ioESTANTE
    .SoloNumeros = True
    .Alineacion = 1
    .LongMaxima = 8
End With

With ioCASILLA
    .SoloNumeros = True
    .Alineacion = 1
    .LongMaxima = 8
End With

With fg
    .Cols = 14
    .Rows = 2
    .ColHidden(1) = True
    .ColAlignment(3) = flexAlignCenterCenter
    .ColAlignment(6) = flexAlignCenterCenter
    .TextMatrix(0, 2) = "Proveedor"
    .TextMatrix(0, 3) = "Ref."
    .TextMatrix(0, 4) = "Modelo"
    .TextMatrix(0, 5) = "Tempor"
    .TextMatrix(0, 6) = "Talla"
    .TextMatrix(0, 7) = "Color"
    .TextMatrix(0, 8) = "Pre.Com"
    .TextMatrix(0, 9) = "En"
    .TextMatrix(0, 10) = "Perchero"
    .TextMatrix(0, 11) = "Estante"
    .TextMatrix(0, 12) = "Casilla"
    .TextMatrix(0, 13) = "Fecha/Hora"
    .AutoSize 2, .Cols - 1
    .SelectionMode = flexSelectionByRow
    .HighLight = flexHighlightWithFocus
End With

ChDir "c:\INVENTARIOS\"

conta_lineas = 0

End Sub


Private Sub avisa_error(mensaje As String)

    Beep
    Espera (1)
    Beep
    Espera (1)
    Beep
    Espera (1)
    
    DoEvents
    lblstatus.Caption = mensaje
    DoEvents
    
    Beep
    Espera (1)
    Beep
    Espera (1)
    Beep
    Espera (1)
    
End Sub


Private Sub avisa_ok(mensaje As String)

    Beep
    Beep
    Beep
    DoEvents
    lblstatus.Caption = mensaje
    DoEvents
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 
  'descargar objetos de la memoria ...
  If Not rc Is Nothing Then
  With rc
    If .State = 1 Then
        .Close
    End If
  End With
  
  Set rc = Nothing
  End If
  
  If Not InvCnn Is Nothing Then
  With InvCnn
    If .State = 1 Then
        .Close
    End If
  End With
  
  Set InvCnn = Nothing
  End If
  
  Set frmInventario = Nothing

End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : ioCODBAR_Validate
' Fecha/Hora  : 06/03/2004 22:07
' Autor       : JCASTILLO
' Propósito   : Inserta un nuevo artículo en la base de datos ...
'---------------------------------------------------------------------------------------
Private Sub ioCODBAR_Validate(Cancel As Boolean)

   On Error GoTo ioCODBAR_Validate_Error

If ioCODBAR.Text = "" Then Exit Sub

If Not creado_o_cargado Then

    If MsgBox("¿Desea crear un nuevo inventario?", vbQuestion + vbYesNo, titulo) = vbNo Then
        Cancel = True
        Exit Sub
    End If

    'crea la base de datos y devuelve el nombre del fichero creado ...
    tmpruta = inicia_inventario

    'si ha habido algun error,        cerrar el formulario ...
    If tmpruta = "@" Then
        Unload Me
        Exit Sub
    End If

    Set InvCnn = New ADODB.Connection
    InvCnn.Open strCnnMdb & tmpruta

    Set rc = New ADODB.Recordset

    'introducir los datos de cabecera en el fichero
    With rc
        .Open "SELECT * FROM CONF_INVEN", InvCnn, adOpenDynamic, adLockOptimistic
        .AddNew
        .fields("CODUSR") = UsuarioActual
        .fields("CODALM") = AlmacenActual
        .Update
        .Close
        .Open "SELECT * FROM INVENTARIO", InvCnn, adOpenStatic, adLockOptimistic
    End With
    
    creado_o_cargado = True

End If


'

If (ioPERCHERO.Text = "") Or Not IsNumeric(ioPERCHERO.Text) Then
    ioPERCHERO.Text = "1"
End If

'
If (ioESTANTE.Text = "") Or Not IsNumeric(ioESTANTE.Text) Then
    ioESTANTE.Text = "1"
End If

'
If (ioCASILLA.Text = "") Or Not IsNumeric(ioCASILLA.Text) Then
    ioCASILLA.Text = "1"
End If

If (Len(ioCODBAR.Text) <> LenCodBar) Or (Not IsNumeric(ioCODBAR.Text)) Then
    DoEvents
    ioCODBAR.Text = ""
    Call avisa_error("Código no Válido")
    ioCODBAR.SetFocus
    Cancel = True
    ioCODBAR.CancelarValidacion
    Exit Sub
'+-+-+
End If
'+-+-+
  miCod = Descompone_CBAR(ioCODBAR.Text)
  
  With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
  End With
  
  
  
'si no existe ...
If devuelve_campo("SELECT CODIGO FROM MAARTIC WHERE CODIGO = " & miCod.CODIGO_ART & " AND TEMPOR = " & miCod.TEMPORADA_ART, locCnn) = "@" Then

    ioCODBAR.Text = ""
    Call avisa_error("El artículo no existe en la base de datos")
    ioCODBAR.SetFocus
    Cancel = True
    ioCODBAR.CancelarValidacion
    Exit Sub
    
End If

'añadir el artículo ...
DoEvents

With rc
    .AddNew
    .fields("CODART") = miCod.CODIGO_ART
    .fields("TEMPOR") = miCod.TEMPORADA_ART
    .fields("CODTALLA") = miCod.TALLA_ART
    .fields("CODCOL") = miCod.COLOR_ART
    .fields("PERCHERO") = ioPERCHERO.Text
    .fields("ESTANTE") = ioESTANTE.Text
    .fields("CASILLA") = ioCASILLA.Text
    .Update

    Call añade_linea_grid(.fields("CODART"), .fields("TEMPOR"), .fields("CODTALLA"), .fields("CODCOL"), .fields("PERCHERO"), .fields("ESTANTE"), .fields("CASILLA"), .fields("ID"))

End With
'+---+
'-+-+-
'--+--
'-+-+-
'+---+

ioCODBAR.Text = ""
ioCODBAR.SetFocus
lblstatus.Caption = ""

Cancel = True

Call avisa_ok("Codigo aceptado correctamente")


   On Error GoTo 0
   Exit Sub

ioCODBAR_Validate_Error:

    Call avisa_error("Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioCODBAR_Validate de Formulario frmInventario")
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioCODBAR_Validate de Formulario frmInventario"

End Sub
'-----------------------------------------------------------------------


Private Sub carga_grid()

Dim t_articulo As Variant
Dim t_tempor As String
Dim t_talla As String
Dim t_color As Variant
Dim var As Byte
Dim stren As String

'limpiar variables ...
   On Error GoTo carga_grid_Error

tmp_totalB = 0
tmp_totalA = 0
tmp_precom = 0
conta_lineas = 0

fg.Clear
fg.Rows = 1



With fg
    .Cols = 14
    .Rows = 2
    .ColHidden(1) = True
    .ColAlignment(3) = flexAlignCenterCenter
    .ColAlignment(6) = flexAlignCenterCenter
    .TextMatrix(0, 2) = "Proveedor"
    .TextMatrix(0, 3) = "Ref."
    .TextMatrix(0, 4) = "Modelo"
    .TextMatrix(0, 5) = "Tempor"
    .TextMatrix(0, 6) = "Talla"
    .TextMatrix(0, 7) = "Color"
    .TextMatrix(0, 8) = "Pre.Com"
    .TextMatrix(0, 9) = "En"
    .TextMatrix(0, 10) = "Perchero"
    .TextMatrix(0, 11) = "Estante"
    .TextMatrix(0, 12) = "Casilla"
    .TextMatrix(0, 13) = "Fecha/Hora"
    .AutoSize 2, .Cols - 1
    .SelectionMode = flexSelectionByRow
    .HighLight = flexHighlightWithFocus
End With


If rc.EOF And Not rc.BOF Then rc.MoveFirst

Do Until rc.EOF

'obtener datos ...
t_articulo = devuelve_matriz("SELECT MODELO, CODPROV, REF, PRECOM, IVACOM from MAARTIC where CODIGO = " & rc.fields("CODART") & " and TEMPOR = " & rc.fields("TEMPOR"), locCnn)
t_tempor = Trim(devuelve_campo("SELECT ABREVIA FROM TEMPOR WHERE IDTEM = " & rc.fields("TEMPOR"), locCnn))
t_talla = Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & rc.fields("CODTALLA"), locCnn))
t_color = devuelve_matriz("SELECT DESCRIPCION, CODCOL FROM COLORES WHERE CODIGO = " & rc.fields("CODCOL"), locCnn)

  
    'si tiene IVA, sumar a importe en A
    If t_articulo(4) > 0 Then
     tmp_totalA = tmp_totalA + t_articulo(3)
     stren = "A"
    'si no tiene IVA sumar a importe en B
    Else
    tmp_totalB = tmp_totalB + t_articulo(3)
    stren = "B"
    End If
    

With fg
    .AddItem "", 2
    .TextMatrix(2, 1) = rc.fields("ID")
    .TextMatrix(2, 2) = Trim(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & t_articulo(1), locCnn))
    .TextMatrix(2, 3) = Trim(t_articulo(2))
    .TextMatrix(2, 4) = Format(rc.fields("CODART"), "00000") & " " & Trim(t_articulo(0))
    .TextMatrix(2, 5) = t_tempor
    .TextMatrix(2, 6) = t_talla
    .TextMatrix(2, 7) = Trim(t_color(0))
    .TextMatrix(2, 8) = Format(t_articulo(3), "Currency")
    .TextMatrix(2, 9) = stren
    .TextMatrix(2, 10) = rc.fields("PERCHERO")
    .TextMatrix(2, 11) = rc.fields("ESTANTE")
    .TextMatrix(2, 12) = rc.fields("CASILLA")
    .TextMatrix(2, 13) = rc.fields("FMODI")
  
    tmp_precom = tmp_precom + t_articulo(3)
    conta_lineas = conta_lineas + 1
    
    lblTotal_General.Caption = Format(tmp_precom, "Currency")
    lblTotal_A.Caption = Format(tmp_totalA, "Currency")
    lblTotal_B.Caption = Format(tmp_totalB, "Currency")
    
    'poner el color
    .Col = 7
    .Row = 2
    .CellBackColor = t_color(1)
    .Col = 2
    
    
 rc.MoveNext
 
 End With
 
 Loop
 
  With fg
    
    .Row = 1
    
    For var = 2 To .Cols - 1
        .Col = var
        .CellBackColor = vbBlue
        .CellForeColor = vbWhite
        .CellFontBold = True
    Next var
        

    
    .TextMatrix(1, 4) = "Unidades: " & conta_lineas
    .TextMatrix(1, 8) = "Total: " & Format(tmp_precom, "Currency")
    .TextMatrix(1, 2) = "Totales"
    .AutoSize 2, .Cols - 1
            
End With


If conta_lineas = 0 Then MsgBox "El inventario seleccionado no tiene datos", vbExclamation, titulo

   On Error GoTo 0
   Exit Sub

carga_grid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_grid de Formulario frmInventario"

End Sub
