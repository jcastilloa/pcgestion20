VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmFlexPed 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Pedidos ..."
   ClientHeight    =   7335
   ClientLeft      =   1305
   ClientTop       =   2295
   ClientWidth     =   11625
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   11625
   Begin PCGestion.ucGrdBttn cbBorrar 
      Height          =   330
      Left            =   10890
      TabIndex        =   22
      Top             =   1095
      Width           =   750
      _extentx        =   1323
      _extenty        =   582
      caption         =   "&Borrar"
      font            =   "frmFlexPed.frx":0000
      image           =   "frmFlexPed.frx":002C
   End
   Begin PCGestion.ucGrdBttn cbLista 
      Height          =   330
      Left            =   9645
      TabIndex        =   21
      Top             =   1095
      Width           =   1260
      _extentx        =   2223
      _extenty        =   582
      caption         =   "&Consultar"
      font            =   "frmFlexPed.frx":004A
      image           =   "frmFlexPed.frx":0076
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel4 
      Height          =   285
      Left            =   2820
      Top             =   1110
      Width           =   6795
      _extentx        =   11986
      _extenty        =   503
      caption         =   "Doble Click ir a pedido seleccionado      -F4- Consultar  -F5- Ir a Rejilla  -F8- Salir"
      fount           =   "frmFlexPed.frx":0094
      captioncolour   =   0
      colour1         =   15640462
      colour2         =   7177785
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   5910
      Left            =   0
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1425
      Visible         =   0   'False
      Width           =   11625
      _cx             =   20505
      _cy             =   10425
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
      FormatString    =   $"frmFlexPed.frx":00C2
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
      Height          =   1380
      Left            =   0
      TabIndex        =   23
      Top             =   30
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   2434
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
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
      TabPicture(0)   =   "frmFlexPed.frx":01A0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label9"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label16"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cbSECCION"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ioNUMPED"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ioREF"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cbALMORIG"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ioCODART"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cbTEMPOR"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cbCODPROV"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Hoja 2"
      TabPicture(1)   =   "frmFlexPed.frx":01BC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cbCATTALL"
      Tab(1).Control(1)=   "cbFAMILIA"
      Tab(1).Control(2)=   "cbSUBFAM"
      Tab(1).Control(3)=   "Label15"
      Tab(1).Control(4)=   "Label14"
      Tab(1).Control(5)=   "Label1"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Hoja 3"
      TabPicture(2)   =   "frmFlexPed.frx":01D8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cbCODTALLA"
      Tab(2).Control(1)=   "cbCODCOL"
      Tab(2).Control(2)=   "cbESTADO"
      Tab(2).Control(3)=   "cbTIPOAB"
      Tab(2).Control(4)=   "ioFECHAINI"
      Tab(2).Control(5)=   "ioFECHAFIN"
      Tab(2).Control(6)=   "Label7"
      Tab(2).Control(7)=   "Label5"
      Tab(2).Control(8)=   "Label11"
      Tab(2).Control(9)=   "Label10"
      Tab(2).Control(10)=   "Label12"
      Tab(2).Control(11)=   "Label13"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Hoja 4"
      TabPicture(3)   =   "frmFlexPed.frx":01F4
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ioFACTURA"
      Tab(3).Control(1)=   "ioSUCODIGO"
      Tab(3).Control(2)=   "ioALBARAN"
      Tab(3).Control(3)=   "ioTRANSPORTI"
      Tab(3).Control(4)=   "chameleonButton1"
      Tab(3).Control(5)=   "Label20"
      Tab(3).Control(6)=   "Label19"
      Tab(3).Control(7)=   "Label18"
      Tab(3).Control(8)=   "Label17"
      Tab(3).ControlCount=   9
      Begin PCGestion.miCombo cbCODPROV 
         Height          =   495
         Left            =   7320
         TabIndex        =   2
         Top             =   30
         Width           =   4230
         _extentx        =   7461
         _extenty        =   873
         font            =   "frmFlexPed.frx":0210
      End
      Begin PCGestion.miCombo cbTEMPOR 
         Height          =   480
         Left            =   5640
         TabIndex        =   5
         Top             =   540
         Width           =   1785
         _extentx        =   3149
         _extenty        =   847
         font            =   "frmFlexPed.frx":023C
      End
      Begin PCGestion.miText ioCODART 
         Height          =   495
         Left            =   4020
         TabIndex        =   4
         Top             =   540
         Width           =   1080
         _extentx        =   1905
         _extenty        =   873
         font            =   "frmFlexPed.frx":0268
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miCombo cbALMORIG 
         Height          =   480
         Left            =   2955
         TabIndex        =   1
         Top             =   30
         Width           =   3765
         _extentx        =   6641
         _extenty        =   847
         font            =   "frmFlexPed.frx":0294
      End
      Begin PCGestion.miText ioREF 
         Height          =   495
         Left            =   765
         TabIndex        =   3
         Top             =   540
         Width           =   2460
         _extentx        =   4339
         _extenty        =   873
         font            =   "frmFlexPed.frx":02C0
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioNUMPED 
         Height          =   480
         Left            =   765
         TabIndex        =   0
         Top             =   30
         Width           =   1350
         _extentx        =   2381
         _extenty        =   847
         font            =   "frmFlexPed.frx":02EC
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miCombo cbCODTALLA 
         Height          =   495
         Left            =   -74220
         TabIndex        =   10
         Top             =   60
         Width           =   2625
         _extentx        =   4630
         _extenty        =   873
         font            =   "frmFlexPed.frx":0318
      End
      Begin PCGestion.miCombo cbCODCOL 
         Height          =   465
         Left            =   -70905
         TabIndex        =   11
         Top             =   60
         Width           =   3405
         _extentx        =   6006
         _extenty        =   820
         font            =   "frmFlexPed.frx":0344
      End
      Begin PCGestion.miCombo cbESTADO 
         Height          =   495
         Left            =   -66480
         TabIndex        =   12
         Top             =   60
         Width           =   3045
         _extentx        =   5821
         _extenty        =   873
         font            =   "frmFlexPed.frx":0370
      End
      Begin PCGestion.miCombo cbTIPOAB 
         Height          =   495
         Left            =   -74220
         TabIndex        =   13
         Top             =   510
         Width           =   2625
         _extentx        =   4630
         _extenty        =   873
         font            =   "frmFlexPed.frx":039C
      End
      Begin PCGestion.miText ioFECHAINI 
         Height          =   480
         Left            =   -67245
         TabIndex        =   14
         Top             =   525
         Width           =   1425
         _extentx        =   2381
         _extenty        =   847
         font            =   "frmFlexPed.frx":03C8
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioFECHAFIN 
         Height          =   480
         Left            =   -64800
         TabIndex        =   15
         Top             =   525
         Width           =   1410
         _extentx        =   2487
         _extenty        =   847
         font            =   "frmFlexPed.frx":03F4
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miCombo cbCATTALL 
         Height          =   495
         Left            =   -68910
         TabIndex        =   9
         Top             =   525
         Width           =   4155
         _extentx        =   5821
         _extenty        =   873
         font            =   "frmFlexPed.frx":0420
      End
      Begin PCGestion.miCombo cbFAMILIA 
         Height          =   480
         Left            =   -73995
         TabIndex        =   7
         Top             =   45
         Width           =   4155
         _extentx        =   7329
         _extenty        =   847
         font            =   "frmFlexPed.frx":044C
      End
      Begin PCGestion.miCombo cbSUBFAM 
         Height          =   480
         Left            =   -68910
         TabIndex        =   8
         Top             =   45
         Width           =   4140
         _extentx        =   7303
         _extenty        =   847
         font            =   "frmFlexPed.frx":0478
      End
      Begin PCGestion.miCombo cbSECCION 
         Height          =   480
         Left            =   8295
         TabIndex        =   6
         Top             =   540
         Width           =   3255
         _extentx        =   5741
         _extenty        =   847
         font            =   "frmFlexPed.frx":04A4
      End
      Begin PCGestion.miText ioFACTURA 
         Height          =   480
         Left            =   -73830
         TabIndex        =   16
         Top             =   30
         Width           =   1425
         _extentx        =   2514
         _extenty        =   847
         font            =   "frmFlexPed.frx":04D0
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioSUCODIGO 
         Height          =   480
         Left            =   -71040
         TabIndex        =   17
         Top             =   30
         Width           =   1425
         _extentx        =   2514
         _extenty        =   847
         font            =   "frmFlexPed.frx":04FC
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioALBARAN 
         Height          =   480
         Left            =   -73845
         TabIndex        =   18
         Top             =   510
         Width           =   1425
         _extentx        =   2514
         _extenty        =   847
         font            =   "frmFlexPed.frx":0528
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioTRANSPORTI 
         Height          =   480
         Left            =   -71040
         TabIndex        =   19
         Top             =   510
         Width           =   2130
         _extentx        =   3757
         _extenty        =   847
         font            =   "frmFlexPed.frx":0554
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.chameleonButton chameleonButton1 
         Height          =   555
         Left            =   -68670
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   225
         Width           =   735
         _extentx        =   1296
         _extenty        =   979
         btype           =   9
         tx              =   ""
         enab            =   -1  'True
         font            =   "frmFlexPed.frx":0580
         coltype         =   1
         focusr          =   -1  'True
         bcol            =   11513775
         bcolo           =   11513775
         fcol            =   0
         fcolo           =   0
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmFlexPed.frx":05AC
         picn            =   "frmFlexPed.frx":05CA
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   0
         ngrey           =   0   'False
         fx              =   0
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TRANSPORTE"
         Height          =   285
         Left            =   -72435
         TabIndex        =   43
         Top             =   675
         Width           =   1350
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ALBARAN"
         Height          =   285
         Left            =   -74985
         TabIndex        =   42
         Top             =   585
         Width           =   1035
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA FACTURA"
         Height          =   660
         Left            =   -72225
         TabIndex        =   41
         Top             =   0
         Width           =   1035
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FACTURA"
         Height          =   285
         Left            =   -74955
         TabIndex        =   40
         Top             =   105
         Width           =   1035
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SECCION"
         Height          =   315
         Left            =   7350
         TabIndex        =   39
         Top             =   630
         Width           =   960
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SUBFAM."
         Height          =   315
         Left            =   -69855
         TabIndex        =   38
         Top             =   120
         Width           =   960
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FAMILIA"
         Height          =   315
         Left            =   -74985
         TabIndex        =   37
         Top             =   120
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CAT. TALLA"
         Height          =   330
         Left            =   -70245
         TabIndex        =   36
         Top             =   615
         Width           =   1305
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "EN"
         Height          =   300
         Left            =   -74670
         TabIndex        =   35
         Top             =   585
         Width           =   450
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO"
         Height          =   345
         Left            =   -67290
         TabIndex        =   34
         Top             =   150
         Width           =   810
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "COLOR"
         Height          =   285
         Left            =   -71610
         TabIndex        =   33
         Top             =   135
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TALLA"
         Height          =   300
         Left            =   -74925
         TabIndex        =   32
         Top             =   150
         Width           =   690
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F.INICIAL"
         Height          =   285
         Left            =   -68295
         TabIndex        =   31
         Top             =   615
         Width           =   1035
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F.FINAL"
         Height          =   285
         Left            =   -65835
         TabIndex        =   30
         Top             =   615
         Width           =   1020
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PEDIDO"
         Height          =   285
         Left            =   15
         TabIndex        =   29
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PROV."
         Height          =   330
         Left            =   6720
         TabIndex        =   28
         Top             =   135
         Width           =   645
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TEMP."
         Height          =   285
         Left            =   5070
         TabIndex        =   27
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CODIGO"
         Height          =   360
         Left            =   3240
         TabIndex        =   26
         Top             =   585
         Width           =   795
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DESTINO"
         Height          =   315
         Left            =   2010
         TabIndex        =   25
         Top             =   120
         Width           =   960
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "REF."
         Height          =   285
         Left            =   0
         TabIndex        =   24
         Top             =   585
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmFlexPed"
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

Dim miosql As New clsSmartSQL
Dim posql As New clsSmartSQL

Dim miRc As New ADODB.Recordset
Public seleccionado As Boolean

Private Sub cbCATTALL_GotFocus()
If Tab1.Tab <> 1 Then Tab1.Tab = 1
End Sub

Private Sub cbCATTALL_Validate(Cancel As Boolean)

If cbCATTALL.Text = "" Then Exit Sub

With cbCODTALLA
    .ConexionString = locCnn
    .LenCodigo = 3
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM TALLAS WHERE CATTALL = " & CLng(cbCATTALL.Text) & " AND MBAJA = 0 ORDER BY CODIGO"
    .DataField = "CODTALLA"
    .carga
    .CodigoWidth = 500
End With

End Sub

Private Sub cbCODTALLA_GotFocus()
If Tab1.Tab <> 2 Then Tab1.Tab = 2
End Sub


Private Sub cbFamilia_GotFocus()
If Tab1.Tab <> 1 Then Tab1.Tab = 1
End Sub

Private Sub cbSECCION_GotFocus()
If Tab1.Tab <> 0 Then Tab1.Tab = 0
End Sub

Private Sub cbTEMPOR_Validate(Cancel As Boolean)

   On Error GoTo cbTEMPOR_Validate_Error

If ioCODART.Text <> "" And cbTEMPOR.Text <> "" Then
 
    If devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & CLng(ioCODART.Text) & " AND TEMPOR = " & CLng(cbTEMPOR.Text)) = "@" Then
        
                MsgBox "No existe el artículo para esa temporada!", vbInformation, titulo
                cbTEMPOR.SetFocus
                Cancel = True
                Exit Sub
    Else
    
        'lblstatus.Caption = ""
        'Call carga_almacenes_origen(cbCODALMORIG)
                
    End If
 
 End If

   On Error GoTo 0
   Exit Sub

cbTEMPOR_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbTEMPOR_Validate de Formulario frmEtiqLibre"
 
End Sub

Private Sub cbBorrar_click()

cbALMORIG.Text = AlmacenActual
cbCODPROV.Text = ""
ioREF.Text = ""
ioCODART.Text = ""
cbTEMPOR.Text = TemporadaActual
cbCATTALL.Text = ""
cbCODTALLA.Text = ""
cbCODCOL.Text = ""
cbESTADO.Text = ""
cbTIPOAB.Text = ""
ioNUMPED.Text = ""

fg.Clear
fg.Rows = 1

'Call cbLista_click

End Sub


Private Sub cbCODPROV_Validate(Cancel As Boolean)
'    If cbCODPROV.Text <> "" Then Call cbLista_click
End Sub

Private Sub cbLista_click()
Dim usa_where As Boolean
Dim usa_Art As Boolean
Dim artsql As New clsSmartSQL

   On Error GoTo cbLista_click_Error


'artsql.ClearWhereClause
'artsql.AddField "CODIGO"

artsql.AddTable "MAARTIC"
artsql.AddField "(CONVERT(char(7), CODIGO) + CONVERT(char(3), TEMPOR))"

'miosql.AddField "CONVERT(char(7), CODART) + CONVERT(char(3), TEMPOR)"

miosql.ClearWhereClause
posql.ClearWhereClause

'filtrar por fechas
If (ioFECHAINI.Text <> "" And ioFECHAFIN.Text <> "") Then
    posql.AddComplexWhereClause "FMODI >= '" & Format(Year((ioFECHAINI.Text)), "0000") & Format(Month((ioFECHAINI.Text)), "00") & Format(Day((ioFECHAINI.Text)), "00") & "' AND FMODI <= '" & Format(Year((ioFECHAFIN.Text)), "0000") & Format(Month((ioFECHAFIN.Text)), "00") & Format(Day((ioFECHAFIN.Text)), "00") & "'", LOGIC_AND
    usa_where = True
End If

If (ioREF.Text <> "") Then
    artsql.AddSimpleWhereClause "REF", ioREF.Text, , CLAUSE_LIKE, LOGIC_AND
    usa_Art = True
End If


If (ioFACTURA.Text <> "") Then
    posql.AddSimpleWhereClause "rtrim(FACTURA)", ioFACTURA.Text
    usa_Art = True
End If


If (ioALBARAN.Text <> "") Then
    posql.AddSimpleWhereClause "rtrim(ALBARAN)", ioALBARAN.Text
    usa_Art = True
End If


If (ioSUCODIGO.Text <> "" And IsDate(ioSUCODIGO.Text)) Then
    posql.AddSimpleWhereClause "rtrim(SUCODIGO)", ioSUCODIGO.Text
    usa_Art = True
End If

If (Trim(ioTRANSPORTI.Text <> "")) Then
    posql.AddSimpleWhereClause "TRNSPORTI", ioTRANSPORTI.Text, , CLAUSE_LIKE, LOGIC_AND
    usa_Art = True
End If


If (cbSECCION.Text <> "") Then
    artsql.AddSimpleWhereClause "SECCION", CLng(cbSECCION.Text), , , LOGIC_AND
    usa_Art = True
End If

If (cbFAMILIA.Text <> "") Then
    artsql.AddSimpleWhereClause "FAMILIA", CLng(cbFAMILIA.Text), , , LOGIC_AND
    usa_Art = True
End If

If (cbSUBFAM.Text <> "") Then
    artsql.AddSimpleWhereClause "SUBFAM", CLng(cbSUBFAM.Text), , , LOGIC_AND
    usa_Art = True
End If

If (cbTEMPOR.Text <> "") Then
    miosql.AddSimpleWhereClause "TEMPOR", cbTEMPOR.Text, , , LOGIC_AND
    usa_where = True
End If

If cbCODPROV.Text <> "" Then
    posql.AddSimpleWhereClause "CODPROV", CLng(cbCODPROV.Text)
    usa_where = True
End If

If cbESTADO.Text <> "" Then
    posql.AddSimpleWhereClause "ESTADO", CLng(cbESTADO.Text)
    usa_where = True
End If

If ioNUMPED.Text <> "" Then
    posql.AddSimpleWhereClause "NUMERO", CLng(ioNUMPED.Text)
    usa_where = True
End If

If cbALMORIG.Text <> "" Then
    posql.AddSimpleWhereClause "ALMORIG", CLng(cbALMORIG.Text)
    usa_where = True
'si no especifica nada, coger el almacen actual
Else
    posql.AddSimpleWhereClause "ALMORIG", AlmacenActual
    cbALMORIG.Text = AlmacenActual
    usa_where = True
End If

If ioCODART.Text <> "" Then
    miosql.AddSimpleWhereClause "CODART", CLng(ioCODART.Text)
    usa_where = True
End If

If cbTEMPOR.Text <> "" Then
    miosql.AddSimpleWhereClause "TEMPOR", CLng(cbTEMPOR.Text)
    usa_where = True
End If

If cbCODTALLA.Text <> "" Then
    miosql.AddSimpleWhereClause "CODTALLA", CLng(cbCODTALLA.Text)
    usa_where = True
End If

If cbCODCOL.Text <> "" Then
    miosql.AddSimpleWhereClause "CODCOL", cbCODCOL.Text
    usa_where = True
End If

If cbTIPOAB.Text <> "" Then
    miosql.AddSimpleWhereClause "DESTINO", cbTIPOAB.Text
    usa_where = True
End If

'si deja todo en blanco, no mostrar ningun registro
If (Not usa_where) And Not (usa_Art) Then
    fg.Clear
    Exit Sub
End If

If usa_Art Then
    miosql.AddComplexWhereClause "(CONVERT(char(7), CODART) + CONVERT(char(3), TEMPOR)) in (" & artsql.SQL & ")"
End If

miosql.AddComplexWhereClause "(CABPEDPRO.NUMERO = DETPEDPRO.NUMERO AND CABPEDPRO.ALMORIG = DETPEDPRO.ALMORIG) and DETPEDPRO.NUMERO IN (" & posql.SQL & ")", LOGIC_AND
'miosql.AddComplexWhereClause ""
  

If miRc.State = 1 Then miRc.Close
miRc.Open miosql.SQL, locCnn, adOpenStatic, adLockOptimistic

fg.Rows = 1

Call carga_grid

fg.HighLight = flexHighlightWithFocus
fg.FocusRect = flexFocusHeavy

DoEvents
Set artsql = Nothing

   On Error GoTo 0
   Exit Sub

cbLista_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbLista_click de Formulario frmFlexPed"
End Sub

Private Sub chameleonButton1_Click()

Dim linea1 As String
Dim linea2 As String
Dim tmpalm As String
         
   On Error GoTo chameleonButton1_Click_Error

    DoEvents

    If cbALMORIG.Text <> "" Then
        tmpalm = devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & cbALMORIG.Text, locCnn)
        If tmpalm = "@" Then tmpalm = ""
    End If
    
    linea1 = "Pedidos:  F.Inicial: " & ioFECHAINI.Text & ". F.Final: " & ioFECHAFIN.Text & ". Almacén: " & tmpalm
    linea2 = "Usuario: " & Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & UsuarioActual, locCnn)) & ". Fecha: " & Now
    
    Call PrintFlexGrid(fg, 1, 1, 2, linea1, linea2, 13, 2, 10)
    'fg.SaveGrid "c:\prueba.txt", flexFileCommaText
    'fg.PrintGrid "Transferencia " & rc.Fields("CODIGO"), , 2, 600, 600
    
    
    'actualizar los colores del grid, volivendolo a cargar
    'rc.Move 0

   On Error GoTo 0
   Exit Sub

chameleonButton1_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento chameleonButton1_Click de Formulario frmFlexPed"

End Sub

Private Sub fg_dblClick()
    
    If fg.Rows <= 1 Then Exit Sub
    seleccionado = True

    If Not IsNumeric(fg.TextMatrix(fg.Row, 1)) Or Not IsNumeric(fg.TextMatrix(fg.Row, 17)) Then Exit Sub
    
    If MsgBox("¿Desea abrir el pedido seleccionado (num " & fg.TextMatrix(fg.Row, 1) & ", " & fg.TextMatrix(fg.Row, 13) & ")?", vbQuestion + vbYesNo, titulo) = vbNo Then Exit Sub

    With frmPedProv
        .trabajar_con_pedido = True
        .codigo_almacen = fg.TextMatrix(fg.Row, 17)
        .NUMERO_PEDIDO = fg.TextMatrix(fg.Row, 1)
        .Show
    End With
    
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
    
        With ioCODART
            .SoloNumeros = True
            .LongMaxima = 5
            .dspFormat = "00000"
        End With
               
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

'consultar
Case vbKeyF4

    Call cbLista_click

'Ir al grid, o regresar
Case vbKeyF5
    
    If fg.Rows > 1 Then
        If fg.TabStop Then
            fg.TabStop = False
            ioNUMPED.SetFocus
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

  Move (Screen.Width - Width) \ 2, Separacion_MDIForm

  fg.Visible = False
  fg.Rows = 1
  fg.Cols = 0
    
       With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
   End With
  
With cbCODPROV
    .ConexionString = locCnn
    .LenCodigo = 5
    .SQLString = "SELECT CODIGO, NOMBRE FROM MAPROV WHERE MBAJA = 0 ORDER BY CODIGO"
    .CodigoWidth = 800
    .carga
    DoEvents
End With

With cbCODTALLA
      .ConexionString = locCnn
      .LenCodigo = 5
      .SQLString = "SELECT CODIGO, DESCRIPCION FROM TALLAs WHERE MBAJA = 0 ORDER BY CODIGO"
      .DataField = "CODTALLA"
      .carga
    '  Set .DataSource = rc
      .CodigoWidth = 800
End With
  
With cbCODCOL
      .ConexionString = locCnn
      .LenCodigo = 5
      .SQLString = "SELECT CODIGO, DESCRIPCION FROM COLORES WHERE MBAJA = 0 ORDER BY CODIGO"
      .DataField = "CODCOL"
      .carga
     ' Set .DataSource = rc
      .CodigoWidth = 800
End With
  
With cbTEMPOR
    .ConexionString = locCnn
    .SQLString = "SELECT IDTEM, ABREVIA FROM TEMPOR WHERE MBAJA = 0 ORDER BY ACTUAL DESC, IDTEM"
    .LenCodigo = 3
    .CodigoWidth = 500
    .DataField = "TEMPOR"
    .carga
   ' Set .DataSource = rc
End With

With cbCATTALL
      .LenCodigo = 4
      .CodigoWidth = 670
      .ConexionString = locCnn
      .SQLString = "SELECT CODIGO, DESCRIPCION FROM CATTALL ORDER BY CODIGO"
      .carga
End With

With cbALMORIG
      .LenCodigo = 4
      .CodigoWidth = 670
      .ConexionString = locCnn
      .SQLString = "SELECT CODIGO, DESCRIPCION FROM ALMACENES ORDER BY CODIGO"
      .carga
      DoEvents
      .Text = AlmacenActual
End With

With cbESTADO
    .añade_item "1  En creación", 1
    .añade_item "2  Parcial", 2
    .añade_item "3  En histórico", 3
    .LenCodigo = 1
    .CodigoWidth = 300
    .Text = "3"
End With

With cbTIPOAB
    .añade_item "0  A", 1
    .añade_item "1  B", 2
    .LenCodigo = 1
    .CodigoWidth = 300
End With

miosql.AddTable "DETPEDPRO"
miosql.AddTable "CABPEDPRO"


posql.AddTable "CABPEDPRO"
posql.AddField "NUMERO"

cbESTADO.Text = "3"
cbALMORIG.Text = AlmacenActual
cbTEMPOR.Text = TemporadaActual

With ioFECHAINI
    .dspFormat = "dd/mm/yyyy"
    .LongMaxima = 10
End With

With ioFECHAFIN
    .dspFormat = "dd/mm/yyyy"
    .LongMaxima = 10
End With

With ioFACTURA
    .SoloNumeros = False
    .LongMaxima = 15
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

  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    tmpstrcombo = ""
    
    Set posql = Nothing
    Set miosql = Nothing
    
   ' Set nif = Nothing
        
    'si hemos establecido un filtro que no devuelve ningun registro,
    'borrar filtro para que no de error al volver al formulario
   ' if mirc.State
    'If miRc.EOF Then Call cbBorrar_click
    
    'No descargar desde aqui, descargar desde el formulario desde donde
    'se llame
    Set frmFlexPed = Nothing
End Sub


Private Sub cbSECCION_Validate(Cancel As Boolean)
If cbSECCION.Text <> "" Then

  
With cbFAMILIA
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM FAMILIAS WHERE MBAJA = 0 AND CODSEC = " & CInt(cbSECCION.Text) & " ORDER BY CODIGO"
    .LenCodigo = 5
    .CodigoWidth = 800
    .carga
    DoEvents
End With

End If

End Sub

Private Sub cbfamilia_Validate(Cancel As Boolean)
If cbFAMILIA.Text <> "" Then

  
With cbSUBFAM
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM SUBFAM WHERE MBAJA = 0 AND CODFAM = " & CInt(cbFAMILIA.Text) & " ORDER BY CODIGO"
    .LenCodigo = 5
    .CodigoWidth = 800
    .carga
    DoEvents
End With

End If

End Sub



Private Sub ioFACTURA_GotFocus()
If Tab1.Tab <> 3 Then Tab1.Tab = 3
End Sub

Private Sub ioFECHAFIN_GotFocus()
If Tab1.Tab <> 2 Then Tab1.Tab = 2
End Sub

Private Sub ioNUMPED_GotFocus()
If Tab1.Tab <> 0 Then Tab1.Tab = 0
End Sub

Private Sub ioNUMPED_Validate(Cancel As Boolean)
'If ioNUMPED.Text <> "" Then Call cbLista_click
End Sub


'---------------------------------------------------------------------------------------
' Procedimiento : carga_grid
' Fecha/Hora    : 22/01/2004 12:59
' Autor         : JCastillo
' Propósito     : Cargar el grid con los resultados
'---------------------------------------------------------------------------------------
'
Private Sub carga_grid()
Dim tmpcodcolor As String
Dim tmprecom As Double
Dim t_articulo As Variant
Dim impdcto As Currency
Dim impiva As Currency

   On Error GoTo carga_grid_Error

   'NUMPED
   'PROVEEDOR
   'CODIGO
   'TEMPORADA
   'TALLA
   'COLOR
   'UDS
   'FECHA
   
   With fg
   
    .Clear
    .Cols = 19
    .Rows = 1
    
    .ColHidden(0) = True
    .ColHidden(18) = True
    
    .ColFormat(9) = "Currency"
    .ColFormat(10) = "Currency"
    .ColFormat(11) = "Currency"
    .ColFormat(12) = "Currency"
    
    .ColAlignment(3) = flexAlignLeftCenter
    
    .TextMatrix(0, 1) = "Pedido"
    .TextMatrix(0, 2) = "Prov."
    .TextMatrix(0, 3) = "Ref."
    .TextMatrix(0, 4) = "Modelo"
    .TextMatrix(0, 5) = "Temp."
    .TextMatrix(0, 6) = "Talla"
    .TextMatrix(0, 7) = "Color"
    .TextMatrix(0, 8) = "Uds."
    .TextMatrix(0, 9) = "P.Com."
    
    .TextMatrix(0, 10) = "Dcto."   'a partir de la 10
    .TextMatrix(0, 11) = "IVA.C."
    
    .TextMatrix(0, 12) = "Total Art."
    .TextMatrix(0, 13) = "P.Ven."
    .TextMatrix(0, 14) = "Almacen"
    .TextMatrix(0, 15) = "En"
    .TextMatrix(0, 16) = "Fecha"
    .TextMatrix(0, 17) = "CBarras"
    
    If miRc.RecordCount <= 0 Then Exit Sub
        
    Do
             .Rows = .Rows + 1
    
        If Not miRc.EOF Then
     
            .TextMatrix(.Rows - 1, 1) = miRc.fields("NUMERO")
            
             t_articulo = devuelve_matriz("SELECT PREVEN, CODPROV, REF FROM MAARTIC WHERE CODIGO = " & miRc.fields("CODART") & " AND TEMPOR = " & miRc.fields("TEMPOR"), locCnn)
            
            'proveedor
            .TextMatrix(.Rows - 1, 2) = Trim(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & t_articulo(1)))
                        
            'ref
            If Not IsNull(t_articulo(2)) Then .TextMatrix(.Rows - 1, 3) = Trim(t_articulo(2))
                                    
            .TextMatrix(.Rows - 1, 4) = Format(miRc.fields("CODART"), "00000") & "-" & Trim(devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & miRc.fields("CODART") & " AND TEMPOR = " & miRc.fields("TEMPOR")))
            If Not IsNull(miRc.fields("TEMPOR")) Then .TextMatrix(.Rows - 1, 5) = Trim(devuelve_campo("SELECT ABREVIA FROM TEMPOR WHERE IDTEM = " & miRc.fields("TEMPOR")))
            If Not IsNull(miRc.fields("CODTALLA")) Then .TextMatrix(.Rows - 1, 6) = Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & miRc.fields("CODTALLA")))
            
            'obtener el texto del color y su codigo de color (para colorear
            'la celda del grid)
            If miRc.fields("CODCOL") > 0 Then
      
                tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & miRc.fields("CODCOL"))
                .TextMatrix(.Rows - 1, 7) = Trim(devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & miRc.fields("CODCOL")))
                .Col = 7
                .Row = .Rows - 1
                .CellBackColor = tmpcodcolor
                .Col = 2
        
            End If
            
           '.TextMatrix(.Rows - 1, 6) = miRc.Fields("CODCOL")
            .TextMatrix(.Rows - 1, 8) = miRc.fields("UNIDADES")
            
            'precio de compra con iva -dcto
            tmprecom = miRc.fields("PRECOM")
            
            .TextMatrix(.Rows - 1, 9) = tmprecom + ((tmprecom * miRc.fields("IVA") / 100))
                                  
            'DCTO
            impdcto = ((miRc.fields("PRECOM") * miRc.fields("DCTO") / 100))
            .TextMatrix(.Rows - 1, 10) = impdcto
            
            'IVA
            impiva = tmprecom * miRc.fields("IVA") / 100
            .TextMatrix(.Rows - 1, 11) = impiva
         
            'Total Articulo
            .TextMatrix(.Rows - 1, 12) = (.TextMatrix(.Rows - 1, 9) * miRc.fields("UNIDADES")) - impdcto + impiva
             
            'precio de venta
            .TextMatrix(.Rows - 1, 13) = t_articulo(0)
            
            'almacen
            .TextMatrix(.Rows - 1, 14) = Trim(devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & miRc.fields("ALMORIG"), locCnn))
                        
            If miRc.fields("DESTINO") = 0 Then
            .TextMatrix(.Rows - 1, 15) = "A"
            Else
            .TextMatrix(.Rows - 1, 15) = "B"
            End If
            
            
            .TextMatrix(.Rows - 1, 16) = miRc.fields("FMODI")
            
            'código de barras
            If (Not IsNull(miRc.fields("CODART"))) And (Not IsNull(miRc.fields("TEMPOR"))) And (Not IsNull(miRc.fields("CODTALLA"))) And (Not IsNull(miRc.fields("TEMPOR"))) Then
                .TextMatrix(.Rows - 1, 17) = Conforma_CB(miRc.fields("CODART"), miRc.fields("TEMPOR"), miRc.fields("CODTALLA"), miRc.fields("CODCOL"))
            End If
 
            .TextMatrix(.Rows - 1, 18) = miRc.fields("ALMORIG")
            
            
     
        End If
    
    If Not miRc.EOF Then miRc.MoveNext
    Loop Until miRc.EOF
      
        .SubtotalPosition = flexSTAbove
        .subtotal flexSTSum, , 8, , vbBlue, vbWhite
        .subtotal flexSTSum, , 10, , vbBlue, vbWhite
        .subtotal flexSTSum, , 11, , vbBlue, vbWhite
        .subtotal flexSTSum, , 12, , vbBlue, vbWhite
        .subtotal flexSTCount, , 4, , vbBlue, vbWhite
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 7) = "Uds.:"
        .TextMatrix(1, 4) = "Articulos: " & .TextMatrix(1, 4)
        .TextMatrix(1, 12) = "Total: " & .TextMatrix(1, 12)
        
    .AutoSize 1, .Cols - 1
    .Redraw = True
 
   End With
   On Error GoTo 0
   Exit Sub

carga_grid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_grid de Formulario frmFlexPed"
End Sub





Private Sub ioREF_Validate(Cancel As Boolean)

'If ioREF.Text <> "" Then Call cbLista_click

End Sub






Private Sub ioSUCODIGO_LostFocus()

   On Error GoTo ioSUCODIGO_LostFocus_Error

If Trim(ioSUCODIGO.Text) = "" Then Exit Sub

If Not IsDate(ioSUCODIGO.Text) Then
    ioSUCODIGO.CancelarValidacion
    ioSUCODIGO.SetFocus
Else
    ioSUCODIGO.Text = Format(ioSUCODIGO.Text, "dd/mm/yyyy")
End If

   On Error GoTo 0
   Exit Sub

ioSUCODIGO_LostFocus_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioSUCODIGO_LostFocus de Formulario frmFlexPed"

End Sub

Private Sub ioTRANSPORTI_GotFocus()
    If Tab1.Tab <> 3 Then Tab1.Tab = 3
End Sub

Private Sub ioTRANSPORTI_LostFocus()
    ioNUMPED.SetFocus
End Sub
