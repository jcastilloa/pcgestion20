VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmFlexStockTallCol 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Stock por Tallas y Colores"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   11685
   Begin PCGestion.ucGrdBttn cbConsulta 
      Height          =   315
      Left            =   9667
      TabIndex        =   26
      Top             =   1095
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      Caption         =   "&Consultar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "frmFlexStockTallCol.frx":0000
   End
   Begin PCGestion.ucGrdBttn cbBorrar 
      Height          =   315
      Left            =   10807
      TabIndex        =   27
      Top             =   1095
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   556
      Caption         =   "&Borrar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "frmFlexStockTallCol.frx":001C
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel4 
      Height          =   285
      Left            =   6210
      Top             =   1110
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   503
      Caption         =   " -F4- Consultar -F5- Ir a Rejilla  -F8- Salir"
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
      Height          =   1395
      Left            =   187
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   2461
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "Hoja 1"
      TabPicture(0)   =   "frmFlexStockTallCol.frx":0038
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label11"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chameleonButton1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cbTEMPOR"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ioCODART"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cbCODCOL"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cbCODTALLA"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ioCODBAR"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Hoja 2"
      TabPicture(1)   =   "frmFlexStockTallCol.frx":0054
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cbCODALM"
      Tab(1).Control(1)=   "cbSeccion"
      Tab(1).Control(2)=   "ioREF"
      Tab(1).Control(3)=   "Label12"
      Tab(1).Control(4)=   "Label3"
      Tab(1).Control(5)=   "Label5"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Hoja 3"
      TabPicture(2)   =   "frmFlexStockTallCol.frx":0070
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cbSUBFAMILIA"
      Tab(2).Control(1)=   "cbFamilia"
      Tab(2).Control(2)=   "ioMODELO"
      Tab(2).Control(3)=   "cbCODPROV"
      Tab(2).Control(4)=   "Label6"
      Tab(2).Control(5)=   "Label4"
      Tab(2).Control(6)=   "Label2"
      Tab(2).Control(7)=   "Label1"
      Tab(2).ControlCount=   8
      Begin PCGestion.miText ioCODBAR 
         Height          =   480
         Left            =   1095
         TabIndex        =   0
         Top             =   45
         Width           =   3345
         _extentx        =   5900
         _extenty        =   847
         font            =   "frmFlexStockTallCol.frx":008C
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miCombo cbCODTALLA 
         Height          =   495
         Left            =   1080
         TabIndex        =   3
         Top             =   510
         Width           =   3300
         _extentx        =   5821
         _extenty        =   873
         font            =   "frmFlexStockTallCol.frx":00B8
      End
      Begin PCGestion.miCombo cbCODCOL 
         Height          =   465
         Left            =   5400
         TabIndex        =   4
         Top             =   510
         Width           =   3540
         _extentx        =   6244
         _extenty        =   820
         font            =   "frmFlexStockTallCol.frx":00E4
      End
      Begin PCGestion.miText ioCODART 
         Height          =   525
         Left            =   5415
         TabIndex        =   1
         Top             =   45
         Width           =   1245
         _extentx        =   2196
         _extenty        =   926
         font            =   "frmFlexStockTallCol.frx":0110
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miCombo cbTEMPOR 
         Height          =   480
         Left            =   7980
         TabIndex        =   2
         Top             =   45
         Width           =   3195
         _extentx        =   7011
         _extenty        =   847
         font            =   "frmFlexStockTallCol.frx":013C
      End
      Begin PCGestion.miCombo cbSUBFAMILIA 
         Height          =   465
         Left            =   -68115
         TabIndex        =   9
         Top             =   90
         Width           =   4245
         _extentx        =   7488
         _extenty        =   820
         font            =   "frmFlexStockTallCol.frx":0168
      End
      Begin PCGestion.miCombo cbFamilia 
         Height          =   465
         Left            =   -73965
         TabIndex        =   8
         Top             =   60
         Width           =   4920
         _extentx        =   8678
         _extenty        =   820
         font            =   "frmFlexStockTallCol.frx":0194
      End
      Begin PCGestion.miText ioMODELO 
         Height          =   480
         Left            =   -73950
         TabIndex        =   10
         Top             =   525
         Width           =   3495
         _extentx        =   6165
         _extenty        =   847
         font            =   "frmFlexStockTallCol.frx":01C0
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miCombo cbCODPROV 
         Height          =   480
         Left            =   -69795
         TabIndex        =   11
         Top             =   525
         Width           =   4050
         _extentx        =   7144
         _extenty        =   847
         font            =   "frmFlexStockTallCol.frx":01EC
      End
      Begin PCGestion.miCombo cbCODALM 
         Height          =   495
         Left            =   -69075
         TabIndex        =   6
         Top             =   15
         Width           =   4935
         _extentx        =   8705
         _extenty        =   873
         font            =   "frmFlexStockTallCol.frx":0218
      End
      Begin PCGestion.miCombo cbSeccion 
         Height          =   495
         Left            =   -73635
         TabIndex        =   7
         Top             =   510
         Width           =   4215
         _extentx        =   7435
         _extenty        =   873
         font            =   "frmFlexStockTallCol.frx":0244
      End
      Begin PCGestion.miText ioREF 
         Height          =   480
         Left            =   -73620
         TabIndex        =   5
         Top             =   30
         Width           =   3495
         _extentx        =   6165
         _extenty        =   847
         font            =   "frmFlexStockTallCol.frx":0270
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.chameleonButton chameleonButton1 
         Height          =   465
         Left            =   10245
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   540
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   820
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
         MICON           =   "frmFlexStockTallCol.frx":029C
         PICN            =   "frmFlexStockTallCol.frx":02B8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REFERENCIA"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -74880
         TabIndex        =   25
         Top             =   120
         Width           =   1230
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SECCION"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -74610
         TabIndex        =   24
         Top             =   585
         Width           =   960
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ALMACEN"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -70080
         TabIndex        =   23
         Top             =   90
         Width           =   960
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PROV."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -70410
         TabIndex        =   22
         Top             =   615
         Width           =   645
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MODELO"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -74820
         TabIndex        =   21
         Top             =   585
         Width           =   840
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FAMILIA"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74940
         TabIndex        =   20
         Top             =   165
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUBFAM"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -68985
         TabIndex        =   19
         Top             =   135
         Width           =   840
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C.BARRAS"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         TabIndex        =   18
         Top             =   120
         Width           =   990
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "COLOR"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4575
         TabIndex        =   17
         Top             =   585
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TALLA"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   345
         TabIndex        =   16
         Top             =   585
         Width           =   690
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CODIGO"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4500
         TabIndex        =   15
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TEMPORADA"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6630
         TabIndex        =   14
         Top             =   105
         Width           =   1320
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   5670
      Left            =   15
      TabIndex        =   12
      Top             =   1410
      Visible         =   0   'False
      Width           =   11700
      _cx             =   20637
      _cy             =   10001
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
      FormatString    =   $"frmFlexStockTallCol.frx":0F92
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
End
Attribute VB_Name = "frmFlexStockTallCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Modulo      : frmFlexStockTallCol
' Fecha/Hora  : 11/01/2004 13:45
' Autor       : JCASTILLO
' Propósito   : Consulta de STOCK por tallas y colores
'---------------------------------------------------------------------------------------
Option Explicit

Dim rc_s As New ADODB.Recordset
Dim prime As Boolean

Private Sub cbBorrar_click()
    
    'limpiar campos
    cbFamilia.Text = ""
    cbSUBFAMILIA.Text = ""
    cbSeccion.Text = ""
    cbCODALM.Text = ""
    cbCODTALLA.Text = ""
    cbCODCOL.Text = ""
    cbTEMPOR.Text = ""
    ioMODELO.Text = ""
    ioCODBAR.Text = ""
    cbCODPROV.Text = ""
    ioCODART.Text = ""
    ioREF.Text = ""
    
    fg.Clear
    fg.Rows = 1
      
    
'    Call cbConsulta_Click
End Sub

Private Sub cbCODCOL_GotFocus()

If Tab1.Tab <> 0 Then Tab1.Tab = 0

End Sub


'---------------------------------------------------------------------------------------
' Subrutina   : cbConsulta_Click
' Fecha/Hora  : 06/01/2004 14:08
' Autor       : JCASTILLO
' Propósito   : Consulta de stock. Modo tabla
'---------------------------------------------------------------------------------------
Private Sub cbConsulta_Click()
Dim oSQL_Maartic As New clsSmartSQL
Dim oSQL_Stock As New clsSmartSQL

   On Error GoTo cbConsulta_Click_Error

oSQL_Maartic.AddTable "MAARTIC"
oSQL_Maartic.AddSimpleWhereClause "MBAJA", 0
'seleccionar solo CODIGO para el select
oSQL_Maartic.AddField "CODIGO"

fg.Clear
fg.Rows = 1

'------------------------------------------------------------------
' PRIMERO OBTENER TODOS LOS ARTICULOS CORRESPONDIENTES AL PATRON
'------------------------------------------------------------------

'codigo articulo
If Trim(ioCODART.Text) <> "" Then
    oSQL_Maartic.AddSimpleWhereClause "CODART", CLng(ioCODART.Text)
End If

'temporada
If Trim(cbTEMPOR.Text) <> "" Then
    oSQL_Maartic.AddSimpleWhereClause "TEMPOR", CLng(cbTEMPOR.Text), , , LOGIC_AND
End If


'seccion
If Trim(cbSeccion.Text) <> "" Then
    oSQL_Maartic.AddSimpleWhereClause "SECCION", CLng(cbSeccion.Text), , , LOGIC_AND
End If

'familia
If Trim(cbFamilia.Text) <> "" Then
    oSQL_Maartic.AddSimpleWhereClause "FAMILIA", CLng(cbFamilia.Text), , , LOGIC_AND
End If

'subfamilia
If Trim(cbSUBFAMILIA.Text) <> "" Then
    oSQL_Maartic.AddSimpleWhereClause "SUBFAM", CLng(cbSUBFAMILIA.Text), , , LOGIC_AND
End If

'modelo
If Trim(ioMODELO.Text) <> "" Then
    oSQL_Maartic.AddSimpleWhereClause "MODELO", ioMODELO.Text, , CLAUSE_LIKE, LOGIC_AND
End If

'referencia
If Trim(ioREF.Text) <> "" Then
    oSQL_Maartic.AddSimpleWhereClause "REF", ioREF.Text, , CLAUSE_LIKE, LOGIC_AND
End If

'proveedor
If Trim(cbCODPROV.Text) <> "" Then
    oSQL_Maartic.AddSimpleWhereClause "CODPROV", CLng(cbCODPROV.Text), , , LOGIC_AND
End If


'------------------------------------------------------------------
' LUEGO SACA TODOS LOS REGISTROS DE STOCK PARA ESOS ARTÍCULOS, teniendo
' en cuenta los parametros TALLA, COLOR y ALMACEN
'------------------------------------------------------------------
oSQL_Stock.AddTable "STOCK"
oSQL_Stock.AddField "CODART"
oSQL_Stock.AddField "TALLA"
oSQL_Stock.AddField "COLOR"
oSQL_Stock.AddField "TEMPOR"
oSQL_Stock.AddField "CODALM"
oSQL_Stock.AddField "STOCK"
oSQL_Stock.AddField "FMODI"

'codigo articulo
If Trim(ioCODART.Text) <> "" Then
    oSQL_Stock.AddSimpleWhereClause "CODART", CLng(ioCODART.Text)
End If

'temporada
If Trim(cbTEMPOR.Text) <> "" Then
    oSQL_Stock.AddSimpleWhereClause "TEMPOR", CLng(cbTEMPOR.Text), , , LOGIC_AND
End If

'talla
If Trim(cbCODTALLA.Text) <> "" Then
    oSQL_Stock.AddSimpleWhereClause "TALLA", CLng(cbCODTALLA.Text), , , LOGIC_AND
End If

'color
If Trim(cbCODCOL.Text) <> "" Then
    oSQL_Stock.AddSimpleWhereClause "COLOR", CLng(cbCODCOL.Text), , , LOGIC_AND
End If

'almacen
If Trim(cbCODALM.Text) <> "" Then
    oSQL_Stock.AddSimpleWhereClause "CODALM", CLng(cbCODALM.Text), , , LOGIC_AND
End If

    'stock mayor q cero, para q no nos muestre los registros a cero
    oSQL_Stock.AddSimpleWhereClause "STOCK", 0, , CLAUSE_DOESNOTEQUAL
    
    oSQL_Stock.AddComplexWhereClause "CODART IN (" & oSQL_Maartic.SQL & ")", LOGIC_AND

If locCnn.State = 1 Then

    rc_s.Open oSQL_Stock.SQL, locCnn

ElseIf locCnnSP.State = 1 Then

    rc_s.Open oSQL_Stock.SQL, locCnnSP

End If

  Call carga_grid
  
  'Set fg.DataSource = rc_s


rc_s.Close

    Set oSQL_Maartic = Nothing
    Set oSQL_Stock = Nothing

   On Error GoTo 0
   Exit Sub

cbConsulta_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbConsulta_Click de Formulario frmFlexStockTallCol"

End Sub



Private Sub cbFamilia_GotFocus()
If Tab1.Tab <> 2 Then Tab1.Tab = 2
End Sub

Private Sub cbfamilia_Validate(Cancel As Boolean)

If cbFamilia.Text <> "" Then
  
With cbSUBFAMILIA
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM SUBFAM WHERE MBAJA = 0 AND CODFAM = " & CInt(cbFamilia.Text) & " ORDER BY CODIGO"
    .LenCodigo = 5
    .CodigoWidth = 800
    .carga
    DoEvents
End With

End If

End Sub


'---------------------------------------------------------------------------------------
' Subrutina   : carga_grid
' Fecha/Hora  : 11/01/2004 12:41
' Autor       : JCASTILLO
' Propósito   : Carga el grid, con las descripciones de los artículos
'---------------------------------------------------------------------------------------
Private Sub carga_grid()
Dim tmpcodcolor As Long
Dim articulo As Variant
Dim tmpsuma As Double
Dim tmpsumapreven As Currency

On Error GoTo carga_grid_Error

    With rc_s
    
    'si no tiene registros salir
    If .RecordCount = 0 Then Exit Sub
    
    If Not .BOF Then .MoveFirst
    
    fg.Clear
    fg.Rows = 1
    fg.Cols = 13
    fg.HighLight = flexHighlightWithFocus
    fg.FocusRect = flexFocusHeavy
    fg.ColHidden(0) = True
    fg.ColFormat(6) = "Currency"
    fg.ColFormat(7) = "Currency"
    fg.ColAlignment(2) = flexAlignLeftCenter
    fg.TextMatrix(0, 1) = "Prov."
    fg.TextMatrix(0, 2) = "Ref"
    fg.TextMatrix(0, 3) = "Modelo"
    fg.TextMatrix(0, 4) = "Talla"
    fg.TextMatrix(0, 5) = "Color"
    
    fg.TextMatrix(0, 6) = "P.Compra"
    fg.TextMatrix(0, 7) = "P.Venta"
    
    fg.TextMatrix(0, 8) = "Tem."
    
    fg.TextMatrix(0, 9) = "Almacen"
    fg.TextMatrix(0, 10) = "STOCK"
    fg.TextMatrix(0, 11) = "CBarras"
    fg.TextMatrix(0, 12) = "Fecha"
    
    '
    
    Do Until .EOF
    
       articulo = devuelve_matriz("SELECT MODELO, REF, CODPROV, PREVEN, PRECOM FROM MAARTIC WHERE CODIGO = " & .fields("CODART").Value & " AND TEMPOR = " & .fields("TEMPOR"), locCnn)
        
       If Not IsArray(articulo) Then
        MsgBox "¡Error al cargar el informe!", vbExclamation, titulo
        Exit Sub
       End If
        
       fg.Rows = fg.Rows + 1
       
       'proveedor
       fg.TextMatrix(fg.Rows - 1, 1) = Trim(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & articulo(2), locCnn))
       
       'referencia
       fg.TextMatrix(fg.Rows - 1, 2) = Trim(articulo(1))
       
       
       'modelo
       fg.TextMatrix(fg.Rows - 1, 3) = Format(.fields("CODART"), "00000") & " " & articulo(0)
       
       'talla
       fg.TextMatrix(fg.Rows - 1, 4) = Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & .fields("TALLA").Value, locCnn))
        
       'obtener el texto del color y su codigo de color (para colorear
       'la celda del grid)
       If Not IsNull(.fields("COLOR")) And .fields("COLOR") <> 0 Then
      
            tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & .fields("COLOR"), locCnn)
            fg.TextMatrix(fg.Rows - 1, 5) = Trim(devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & .fields("COLOR"), locCnn))
            fg.Col = 5
            fg.Row = fg.Rows - 1
            fg.CellBackColor = tmpcodcolor
            fg.Col = 2
        
       End If
       
       'precio compra
       fg.TextMatrix(fg.Rows - 1, 6) = articulo(4) 'Obtiene_Precom_Pedido(.Fields("CODART").Value, .Fields("TEMPOR").Value, .Fields("TALLA").Value, .Fields("COLOR").Value, locCnn)
       
       'precio venta
       fg.TextMatrix(fg.Rows - 1, 7) = articulo(3)
       
       'ir sumando el dinero  (precom * unidades)
       tmpsuma = tmpsuma + (articulo(4) * .fields("STOCK"))
       'lo mismo para preven
       tmpsumapreven = tmpsumapreven + (articulo(3) * .fields("STOCK"))
       
       fg.TextMatrix(fg.Rows - 1, 8) = Trim(devuelve_campo("SELECT ABREVIA FROM TEMPOR WHERE IDTEM = " & .fields("TEMPOR"), locCnn))
       fg.TextMatrix(fg.Rows - 1, 9) = Trim(devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & .fields("CODALM"), locCnn))
       fg.TextMatrix(fg.Rows - 1, 10) = .fields("STOCK")
       fg.TextMatrix(fg.Rows - 1, 11) = Conforma_CB(.fields("CODART"), .fields("TEMPOR"), .fields("TALLA"), .fields("COLOR"))
       fg.TextMatrix(fg.Rows - 1, 12) = .fields("FMODI")
      
       .MoveNext
       
    Loop
    
       fg.SubtotalPosition = flexSTAbove
       fg.subtotal flexSTSum, , 10, , vbBlue, vbWhite, True
       fg.TextMatrix(1, 0) = ""
       fg.TextMatrix(1, 5) = "Importe:"
       fg.TextMatrix(1, 6) = tmpsuma
       fg.TextMatrix(1, 7) = tmpsumapreven
       fg.TextMatrix(1, 9) = "Dif: " & Format(tmpsumapreven - tmpsuma, "Currency")
        
    End With
    
    DoEvents

    fg.ColFormat(1) = "00000"
    fg.AutoSize 1, fg.Cols - 1
    
   On Error GoTo 0
   Exit Sub

carga_grid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_grid de Formulario frmFlexStockTallCol"
End Sub

Private Sub cbSECCION_lostfocus()

If cbSeccion.Text <> "" Then

With cbFamilia
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM FAMILIAS WHERE MBAJA = 0 AND CODSEC = " & CInt(cbSeccion.Text) & " ORDER BY CODIGO"
    .LenCodigo = 5
    .CodigoWidth = 800
    .carga
    DoEvents
End With

End If

End Sub



Private Sub chameleonButton1_Click()
Dim linea1 As String
Dim linea2 As String
Dim tmpalm As String

   On Error GoTo chameleonButton1_Click_Error

    DoEvents

    If cbCODALM.Text <> "" Then
        tmpalm = devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & cbCODALM.Text, locCnn)
        If tmpalm = "@" Then tmpalm = ""
    End If
    
    linea1 = "Consulta Stock: .Almacén: " & tmpalm
    linea2 = "Usuario: " & Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & UsuarioActual, locCnn)) & ". Fecha: " & Now
    
    Call PrintFlexGrid(fg, 1, 1, 2, linea1, linea2, 13, 2, 10)
    
    'fg.SaveGrid "c:\prueba.txt", flexFileCommaText
    'fg.PrintGrid "Transferencia " & rc.Fields("CODIGO"), , 2, 600, 600
    
    
    'actualizar los colores del grid, volivendolo a cargar
    'rc.Move 0


   On Error GoTo 0
   Exit Sub

chameleonButton1_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento chameleonButton1_Click de Formulario frmFlexStockTallCol"

End Sub





'Private Sub cbLista_click()
'  Dim objparam As ADODB.Parameter
'  Dim objcmd As New ADODB.Command
'  Dim SQL As String
'  Dim cond As String
'  Dim Row As Integer
  
'  cond = "1=1 "
'  If Len(cbARTICULO.Text) > 0 Then
'    cond = cond & " AND s.ARTICULO=" & cbARTICULO.Text
'    If Len(cbCODALM.Text) > 0 Then
'        cond = cond & " AND s.CODCENTRO=" & cbCODALM.Text
'    End If
'  End If
'  If Len(ioDESC.Text) > 0 Then
'    cond = cond & " AND s.MODELO LIKE '%" & ioDESC.Text & "%' "
'  End If
'
  
'  If miRc.State > 0 Then
'    miRc.Close
'  End If
'
'  If Len(cond) = 4 Then cond = " s.ARTICULO IS NULL "
'
'  If chkTable.Value = vbUnchecked Then
'      objcmd.CommandText = "sp_crosstab"
'      objcmd.CommandType = adCmdStoredProc
'      Set objcmd.ActiveConnection = locCnn
'      objcmd.Parameters.Refresh
'
'      objcmd.Parameters("@table") = "(SELECT * FROM STOCKTALLCOL s WHERE " & cond & ") STOCK "
'
 '     objcmd.Parameters("@onrows") = "TALLAS"
    
'      objcmd.Parameters("@oncols") = "COLORES"
'
'      objcmd.Parameters("@sumcol") = "STOCK"
'
'      miRc.Open objcmd, , adOpenStatic, adLockOptimistic
'  Else
'      SQL = "SELECT CENTRO, TALLAS, COLORES, CODCOLOR, STOCK FROM STOCKTALLCOL s WHERE " & cond
'      miRc.Open SQL, locCnn, adOpenStatic, adLockOptimistic
'  End If
'
'  Set fg.DataSource = miRc
'  DoEvents
'
'  If chkTable.Value = vbChecked Then
'     For Row = 1 To fg.Rows - 1
'        fg.Cell(flexcpBackColor, Row, 4, Row, 4) = fg.ValueMatrix(Row, 4)
'        fg.Cell(flexcpForeColor, Row, 4, Row, 4) = fg.ValueMatrix(Row, 4)
'     Next Row
'  Else
'    'pendiente el coloreo en tabla pivot
'  End If
'  fg.ColFormat(1) = "0000"
'  fg.AutoSize 1, fg.Cols - 1
'End Sub

'Private Sub cbSUBFAMILIA_Click()
'    With cbARTICULO
'        .ConexionString = locCnn
'        .SQLString = "SELECT CODIGO, MODELO FROM MAARTIC WHERE FAMILIA=" & cbFamilia.Text & " AND SUBFAM=" & cbSUBFAMILIA.Text & " ORDER BY CODIGO"
'        .LenCodigo = 3
'        Set .DataSource = rc
'        .carga
'    End With
'End Sub

'Private Sub cbSUBFAMILIA_Validate(Cancel As Boolean)
'    cbSUBFAMILIA_Click
'End Sub

'Private Sub chkTable_Click()
'    cbLista_click
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode

'Ir al grid, o regresar
Case vbKeyF5
    
    If fg.Rows > 1 Then
        If fg.TabStop Then
            fg.TabStop = False
            If Tab1.Tab <> 0 Then Tab1.Tab = 0
            ioCODBAR.SetFocus
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
    Call cbConsulta_Click
    KeyCode = 0

End Select

End Sub

Private Sub Form_Activate()
    If Not prime Then
        'Set fg.DataSource = miRc
        DoEvents
        fg.Visible = True
        fg.Rows = 1
'        fg.AutoSize 1, fg.Cols - 1
        fg.AllowSelection = True
        fg.HighLight = flexHighlightNever
        prime = True
    End If
End Sub

Private Sub Form_Load()

   Move (Screen.Width - Width) \ 2, Separacion_MDIForm
   
   fg.Cols = 0
   fg.Rows = 1
   fg.TabStop = False
   
   With locCnn
        If .State = 0 Then
            .CursorLocation = adUseClient
            .Open strLocCnn
        End If
   End With

   With ioCODBAR
    .LongMaxima = LenCodBar
    .SoloNumeros = True
   End With

   With cbFamilia
        .ConexionString = locCnn
        .SQLString = "SELECT CODIGO, DESCRIPCION FROM FAMILIAS ORDER BY CODIGO"
        .LenCodigo = 5
        .CodigoWidth = 800
        'Set .DataSource = rc
        .carga
    End With
    
    With cbSUBFAMILIA
        .ConexionString = locCnn
        .SQLString = "SELECT CODIGO, DESCRIPCION FROM SUBFAM WHERE MBAJA = 0 ORDER BY CODIGO"
        .LenCodigo = 5
        .CodigoWidth = 800
        .carga
    End With

    With cbSeccion
        .ConexionString = locCnn
        .SQLString = "SELECT CODIGO, DESCRIPCION FROM SECCIONES WHERE MBAJA = 0 ORDER BY CODIGO"
        .LenCodigo = 5
        .carga
        .CodigoWidth = 800
    End With

    With cbTEMPOR
        .ConexionString = locCnn
        .SQLString = "SELECT IDTEM, AÑO + '  ' + TEMPORADA FROM TEMPOR WHERE MBAJA = 0 ORDER BY ACTUAL DESC, IDTEM"
        .LenCodigo = 3
        .CodigoWidth = 500
        .carga
    End With
    
    With cbCODALM
        .ConexionString = locCnn
        .SQLString = "SELECT CODIGO, DESCRIPCION FROM ALMACENES ORDER BY CODIGO"
        .LenCodigo = 3
        .CodigoWidth = 500
        'Set .DataSource = rc
        .carga
    End With
    
    With cbCODTALLA
        .ConexionString = locCnn
        .SQLString = "SELECT CODIGO, DESCRIPCION FROM TALLAS ORDER BY CODIGO"
        .LenCodigo = 5
        .CodigoWidth = 500
        
        .carga
    End With
    
    With cbCODPROV
        .ConexionString = locCnn
        .SQLString = "SELECT CODIGO, NOMBRE FROM MAPROV ORDER BY CODIGO"
        .LenCodigo = 5
        .CodigoWidth = 800
        
        .carga
    End With
    
    With cbCODCOL
        .ConexionString = locCnn
        .SQLString = "SELECT CODIGO, DESCRIPCION FROM COLORES ORDER BY CODIGO"
        .LenCodigo = 5
        .CodigoWidth = 800
        .carga
    End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'descargar objetos
Set rc_s = Nothing
'Set oSQL_Maartic = Nothing
'Set oSQL_Stock = Nothing
Set frmFlexStockTallCol = Nothing

End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : ioCODBAR_Validate
' Fecha/Hora  : 06/01/2004 15:00
' Autor       : JCASTILLO
' Propósito   :
'---------------------------------------------------------------------------------------
Private Sub ioCODBAR_Validate(Cancel As Boolean)

   On Error GoTo ioCODBAR_Validate_Error

With ioCODBAR

If Trim(.Text) = "" Then Exit Sub

    'si es un codigo de barras con la longitud válidad
    If Len(Trim(.Text)) = LenCodBar Then
        
        'comprobar si existe el artículo/temporada
        If devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & CLng(Left(.Text, 5)) & " AND TEMPOR = " & CLng(Mid(.Text, 6, 3))) = "@" Then
        
                MsgBox "No existe el artículo para esa temporada!, Codigo de Barras no Válido", vbInformation, titulo
                .Text = ""
                .SetFocus
                .CancelarValidacion
                Cancel = True
                Exit Sub
                
        End If
        
             'codigo de artículo
        ioCODART.Text = CLng(Left(.Text, 5))
        'temporada
        cbTEMPOR.Text = CLng(Mid(.Text, 6, 3))
                
        'talla
        cbCODTALLA.Text = CLng(Mid(.Text, 9, 2))
        'color
        cbCODCOL.Text = CLng(Mid(.Text, 11, 3))
        

    Else

        MsgBox "Código de Barras no válido", vbInformation
        .Text = ""
        .SetFocus
        .CancelarValidacion
        Cancel = True
        Exit Sub

    End If
    
End With

   On Error GoTo 0
   Exit Sub

ioCODBAR_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioCODBAR_Validate de Formulario frmFlexStockTallCol"

End Sub

Private Sub ioMODELO_Validate(Cancel As Boolean)
    If ioMODELO.Text <> "" Then Call cbConsulta_Click
End Sub

Private Sub ioREF_GotFocus()
    If Tab1.Tab <> 1 Then Tab1.Tab = 1
End Sub

Private Sub ioREF_Validate(Cancel As Boolean)
    If ioMODELO.Text <> "" Then Call cbConsulta_Click
End Sub
