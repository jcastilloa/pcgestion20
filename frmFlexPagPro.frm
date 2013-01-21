VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmFlexPagPro 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver Pagos a..."
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
      TabIndex        =   5
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
      Image           =   "frmFlexPagPro.frx":0000
   End
   Begin PCGestion.ucGrdBttn cbLista 
      Height          =   315
      Left            =   9450
      TabIndex        =   4
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
      Image           =   "frmFlexPagPro.frx":001C
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
      Left            =   30
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1545
      Visible         =   0   'False
      Width           =   11415
      _cx             =   20135
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
      FormatString    =   $"frmFlexPagPro.frx":0038
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
      Left            =   30
      TabIndex        =   6
      Top             =   0
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   2699
      _Version        =   393216
      TabOrientation  =   1
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
      TabCaption(0)   =   "Hoja 1"
      TabPicture(0)   =   "frmFlexPagPro.frx":0116
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label13"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label12"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cbTIPOPAGO"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ioFECHAFIN"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ioFECHAINI"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cbCODPROV"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ioIMPORTE"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cbCAJAS"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Hoja 2"
      TabPicture(1)   =   "frmFlexPagPro.frx":0132
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "chameleonButton1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cbESTADO"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin PCGestion.miCombo cbCAJAS 
         Height          =   525
         Left            =   630
         TabIndex        =   1
         Top             =   585
         Width           =   3780
         _extentx        =   7303
         _extenty        =   926
         font            =   "frmFlexPagPro.frx":014E
      End
      Begin PCGestion.miText ioIMPORTE 
         Height          =   480
         Left            =   10065
         TabIndex        =   0
         Top             =   615
         Width           =   1260
         _extentx        =   2223
         _extenty        =   847
         font            =   "frmFlexPagPro.frx":017A
         dspformat       =   ""
         enabled         =   -1
         espassword      =   -1
      End
      Begin PCGestion.miCombo cbCODPROV 
         Height          =   525
         Left            =   7290
         TabIndex        =   2
         Top             =   90
         Width           =   4005
         _extentx        =   7064
         _extenty        =   926
         font            =   "frmFlexPagPro.frx":01A6
      End
      Begin PCGestion.miText ioFECHAINI 
         Height          =   495
         Left            =   5520
         TabIndex        =   10
         Top             =   600
         Width           =   1380
         _extentx        =   2434
         _extenty        =   873
         font            =   "frmFlexPagPro.frx":01D2
         dspformat       =   ""
         enabled         =   -1
         espassword      =   -1
      End
      Begin PCGestion.miText ioFECHAFIN 
         Height          =   495
         Left            =   7770
         TabIndex        =   11
         Top             =   600
         Width           =   1380
         _extentx        =   2434
         _extenty        =   873
         font            =   "frmFlexPagPro.frx":01FE
         dspformat       =   ""
         enabled         =   -1
         espassword      =   -1
      End
      Begin PCGestion.miCombo cbESTADO 
         Height          =   525
         Left            =   -73800
         TabIndex        =   14
         Top             =   135
         Width           =   2565
         _extentx        =   4524
         _extenty        =   926
         font            =   "frmFlexPagPro.frx":022A
      End
      Begin PCGestion.miCombo cbTIPOPAGO 
         Height          =   525
         Left            =   630
         TabIndex        =   16
         Top             =   75
         Width           =   3780
         _extentx        =   7303
         _extenty        =   926
         font            =   "frmFlexPagPro.frx":0256
      End
      Begin PCGestion.chameleonButton chameleonButton1 
         Height          =   555
         Left            =   -71160
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   165
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
         MICON           =   "frmFlexPagPro.frx":0282
         PICN            =   "frmFlexPagPro.frx":029E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CAJA"
         Height          =   330
         Left            =   75
         TabIndex        =   17
         Top             =   675
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO"
         Height          =   330
         Left            =   -74715
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F. INICIAL"
         Height          =   330
         Left            =   4335
         TabIndex        =   13
         Top             =   660
         Width           =   1170
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F. FINAL"
         Height          =   330
         Left            =   6810
         TabIndex        =   12
         Top             =   675
         Width           =   960
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO"
         Height          =   330
         Left            =   45
         TabIndex        =   9
         Top             =   135
         Width           =   540
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IMPORTE"
         Height          =   330
         Left            =   9165
         TabIndex        =   8
         Top             =   690
         Width           =   885
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PROVEEDOR"
         Height          =   330
         Left            =   5790
         TabIndex        =   7
         Top             =   165
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmFlexPagPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Módulo     : frmFlexPagPro
' Fecha/Hora : 09/07/2004 13:10
' Autor      : JCastillo
' Propósito  :  detalle de pagos
'---------------------------------------------------------------------------------------

Option Explicit

Dim first As Boolean

'Dim tmpstrcombo As String

Dim CabPagSQL As New clsSmartSQL
Dim DetPagSQL As New clsSmartSQL

Dim miRc As New ADODB.Recordset
Dim seleccionado As Boolean

Public Pago_A_Proveedor  As Boolean



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
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Sub cbBorrar_click()

ioIMPORTE.Text = ""
ioFECHAINI.Text = Date
cbCAJAS.Text = CajaActual
cbCODPROV.Text = ""
ioFECHAINI.Text = ""
ioFECHAFIN.Text = ""
cbESTADO.Text = ""
cbTIPOPAGO.Text = ""

fg.Clear
fg.Rows = 1

End Sub

Private Sub cbtipopago_GotFocus()

If Tab1.Tab <> 0 Then Tab1.Tab = 0

End Sub

Private Sub cbestado_GotFocus()
If Tab1.Tab <> 1 Then Tab1.Tab = 1
End Sub



Private Sub cbLista_click()
Dim usa_where As Boolean
Dim nuefech As String

   On Error GoTo cbLista_click_Error

CabPagSQL.ClearWhereClause
DetPagSQL.ClearWhereClause


'comprobar si quiere solo la fecha de hoy
If (ioFECHAINI.Text <> "") And IsDate(ioFECHAINI.Text) And (ioFECHAFIN.Text = "") Then

    'CabPagSQL.AddSimpleWhereClause "FALTA", ioFECHA.Text, , CLAUSE_GREATERTHANOREQUAL
    'CabPagSQL.AddSimpleWhereClause "FALTA", CStr(DateAdd("d", 1, ioFECHA.Text)), , CLAUSE_LESSTHAN, LOGIC_AND
    'CabPagSQL.AddComplexWhereClause "Year(FALTA IN (" & DetPagSQL.SQL & ")", LOGIC_AND
    
    '>= q el dia actual
    '< que el dia siguiente
    nuefech = DateAdd("d", 1, ioFECHAINI.Text)
    CabPagSQL.AddComplexWhereClause "(FMODI >= '" & Format(Year((ioFECHAINI.Text)), "0000") & Format(Month((ioFECHAINI.Text)), "00") & Format(Day((ioFECHAINI.Text)), "00") & "' AND FMODI < '" & Format(Year((nuefech)), "0000") & Format(Month((nuefech)), "00") & Format(Day((nuefech)), "00") & "')", LOGIC_AND
    usa_where = True
         
  
   
'comprobar si quiere un rango de fechas
ElseIf (ioFECHAINI.Text <> "" And ioFECHAINI.Text <> "") And (IsDate(ioFECHAINI.Text) And IsDate(ioFECHAINI.Text)) Then

    CabPagSQL.AddComplexWhereClause "FMODI >= '" & Format(Year((ioFECHAINI.Text)), "0000") & Format(Month((ioFECHAINI.Text)), "00") & Format(Day((ioFECHAINI.Text)), "00") & "' AND FMODI <= '" & Format(Year((ioFECHAFIN.Text)), "0000") & Format(Month((ioFECHAFIN.Text)), "00") & Format(Day((ioFECHAFIN.Text)), "00") & "'", LOGIC_AND
    usa_where = True

End If

 
If ioIMPORTE.Text <> "" Then
    If CDbl(ioIMPORTE.Text) > 0 Then
        DetPagSQL.AddSimpleWhereClause "IMPORTE", CDbl(ioIMPORTE.Text), , , LOGIC_AND
        usa_where = True
    End If
End If

If cbTIPOPAGO.Text <> "" Then
    CabPagSQL.AddSimpleWhereClause "TIPOPAGO", CLng(cbTIPOPAGO.Text), , , LOGIC_AND
    usa_where = True
Else

    If Pago_A_Proveedor Then
        CabPagSQL.AddSimpleWhereClause "TIPOPAGO", 1
        usa_where = True
    Else
        CabPagSQL.AddSimpleWhereClause "TIPOPAGO", 1, , CLAUSE_DOESNOTEQUAL
        usa_where = True
    End If

End If

If cbCODPROV.Text <> "" Then
    CabPagSQL.AddSimpleWhereClause "CODPROV", cbCODPROV.Text, , , LOGIC_AND
    usa_where = True
End If

If cbCAJAS.Text <> "" Then
    CabPagSQL.AddSimpleWhereClause "CODCAJA", cbCAJAS.Text, , , LOGIC_AND
    usa_where = True
End If

If cbESTADO.Text <> "" Then
    CabPagSQL.AddSimpleWhereClause "ESTADO", cbESTADO.Text, , , LOGIC_AND
    usa_where = True
Else
    CabPagSQL.AddSimpleWhereClause "ESTADO", 3, , CLAUSE_DOESNOTEQUAL, LOGIC_AND
    usa_where = True
End If


  
'If usa_artic Then DetPagSQL.AddComplexWhereClause "(CONVERT(char(10), CODART) + CONVERT(char(3), TEMPOR)) IN (" & artsql.SQL & ")", LOGIC_AND

'If usa_artic Then DetPagSQL.AddComplexWhereClause "COD IN (" & ArtSql.SQL & ")", LOGIC_AND
If usa_where Then DetPagSQL.AddComplexWhereClause "(CONVERT(char(10), CODIGO) + CONVERT(char(3), CODCAJA)) IN (" & CabPagSQL.SQL & ")", LOGIC_AND

'If ioNOMBRE.Text <> "" Then CabPagSQL.AddComplexWhereClause "CODCOST IN (" & DetPagSQL.SQL & ")", LOGIC_AND




If miRc.State = 1 Then miRc.Close
miRc.Open DetPagSQL.SQL, locCnn, adOpenStatic, adLockOptimistic

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
 

cbTIPOPAGO.SetFocus
   On Error GoTo 0
   Exit Sub

cbLista_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbLista_click de Formulario frmFlexArre"
    cbTIPOPAGO.SetFocus


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
Dim Dsp_Prov As String
Dim Dsp_Tipo As String
Dim tmpcodigo As Long
Dim tmpcodcaja As Byte

Dim tmptotal As Currency
Dim tmppendi As Currency
Dim tmppagado As Currency

   On Error GoTo carga_grid_Error

   
   With fg
   
    .Clear
    .Cols = 15
        
    .ColFormat(5) = "Currency"
    .ColFormat(6) = "Currency"
    .ColFormat(7) = "Currency"
    .ColFormat(8) = "Currency"
    .ColFormat(9) = "Currency"
    
    .ColHidden(0) = True
    .ColHidden(1) = True
    .Rows = 1
    
    .TextMatrix(0, 2) = "Fecha"
    .TextMatrix(0, 3) = "Tipo"
    .TextMatrix(0, 4) = "Prov."
    .TextMatrix(0, 5) = "Importe"
    .TextMatrix(0, 6) = "Total"
    .TextMatrix(0, 7) = "Pagado"
    .TextMatrix(0, 8) = "Pendi."
    .TextMatrix(0, 9) = "Cuota"
    .TextMatrix(0, 10) = "Meses"
    .TextMatrix(0, 11) = "Pedido"
    .TextMatrix(0, 12) = "Factura"
    .TextMatrix(0, 13) = "Estado"
    .TextMatrix(0, 14) = "Comen."
        
    If (miRc.EOF And miRc.BOF) Then Exit Sub
        
    tmpcodigo = miRc.fields("CODIGO")
    tmpcodcaja = miRc.fields("CODCAJA")
    
    'coge datos de la cabecera del pago
    T_Cabecera = devuelve_matriz("SELECT TIPOPAGO, CODPROV, IMPORTE, PAGADO, CUOTA, MESES, NUMPED, FACTURA, ESTADO, DESCRIPCION FROM PAGOS WHERE CODIGO = " & miRc.fields("CODIGO") & " AND CODCAJA = " & miRc.fields("CODCAJA"), locCnn)
    Dsp_Prov = Trim(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & T_Cabecera(1), locCnn))
    Dsp_Tipo = Trim(devuelve_campo("SELECT DESCRIPCION FROM MAPAG WHERE CODIGO = " & T_Cabecera(0), locCnn))
    
    tmptotal = tmptotal + T_Cabecera(2)
    tmppagado = tmppagado + T_Cabecera(3)
    tmppendi = tmptotal - tmppagado
         
         
    Do
             .Rows = .Rows + 1
    
        If Not miRc.EOF Then
     
            conta_lineas = conta_lineas + 1
            
            'romper por codigo y buscar una nueva descripción de cabecera
            If (miRc.fields("CODIGO") <> tmpcodigo) And (miRc.fields("CODCAJA") <> tmpcodcaja) Then
                
                'coge datos de la cabecera del pago
                T_Cabecera = devuelve_matriz("SELECT TIPOPAGO, CODPROV, IMPORTE, PAGADO, CUOTA, MESES, PEDIDO, FACTURA, ESTADO, DESCRIPCION FROM PAGOS WHERE CODIGO = " & miRc.fields("CODIGO") & " AND CODCAJA = " & miRc.fields("CODCAJA"), locCnn)
                Dsp_Prov = Trim(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & T_Cabecera(1), locCnn))
                Dsp_Tipo = Trim(devuelve_campo("SELECT DESCRIPCION FROM MAPAG WHERE CODIGO = " & T_Cabecera(0), locCnn))
   
                tmpcodigo = miRc.fields("CODIGO")
                tmpcodcaja = miRc.fields("CODCAJA")
                
                tmptotal = tmptotal + T_Cabecera(2)
                tmppagado = tmppagado + T_Cabecera(3)
                tmppendi = tmptotal - tmppagado
            
            End If
            
            'ID
            .TextMatrix(.Rows - 1, 0) = conta_lineas
            'CAJA
            .TextMatrix(.Rows - 1, 1) = miRc.fields("CODCAJA")
            
            'FECHA (cogerla de la cabecera)
            .TextMatrix(.Rows - 1, 2) = miRc.fields("FMODI")
            
            'Tipo
            .TextMatrix(.Rows - 1, 3) = Dsp_Tipo
                                  
            'Proveedor
            .TextMatrix(.Rows - 1, 4) = Dsp_Prov
            
            'Importe (detalle)
            .TextMatrix(.Rows - 1, 5) = miRc.fields("IMPORTE")
            
            'Total
            .TextMatrix(.Rows - 1, 6) = T_Cabecera(2)
            
            'Pagado
            .TextMatrix(.Rows - 1, 7) = T_Cabecera(3)
            
            'Pendiente
            .TextMatrix(.Rows - 1, 8) = T_Cabecera(2) - T_Cabecera(3)
            
            'Cuota
            .TextMatrix(.Rows - 1, 9) = T_Cabecera(4)
            
            'Meses
            .TextMatrix(.Rows - 1, 10) = T_Cabecera(5)
            
            'numped
            .TextMatrix(.Rows - 1, 11) = T_Cabecera(6)
            
            'factura
            .TextMatrix(.Rows - 1, 12) = T_Cabecera(7)
            
            'estado
            Select Case T_Cabecera(8)
            
                Case 0
                    .TextMatrix(.Rows - 1, 13) = "Pendiente"
                
                Case 1
                    .TextMatrix(.Rows - 1, 13) = "Parcial"
                
                
                Case 2
                    .TextMatrix(.Rows - 1, 13) = "Pagado"
                
                Case 3
                    .TextMatrix(.Rows - 1, 13) = "Anulado"
            
            End Select
            
            'comentario
            .TextMatrix(.Rows - 1, 14) = T_Cabecera(9)
                       
            
        End If
    
    If Not miRc.EOF Then miRc.MoveNext
    
    'DoEvents
    
    Loop Until miRc.EOF
          
        .SubtotalPosition = flexSTAbove
        .subtotal flexSTCount, -1, 3, , vbBlue, vbWhite, True
        .subtotal flexSTSum, -1, 5, , vbBlue, vbWhite, True
        '.subtotal flexSTSum, -1, 6, , vbBlue, vbWhite, True
        '.subtotal flexSTSum, -1, 7, , vbBlue, vbWhite, True
        '.subtotal flexSTSum, -1, 8, , vbBlue, vbWhite, True
        '.subtotal flexSTSum, -1, 9, , vbBlue, vbWhite, True
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 3) = "Nº Pagos: (" & .TextMatrix(1, 3) & ")"
        .TextMatrix(1, 4) = ""
        .TextMatrix(1, 6) = tmptotal
        .TextMatrix(1, 7) = tmppagado
        .TextMatrix(1, 8) = tmppendi

    .AutoSize 1, .Cols - 1
    .Redraw = True

  End With
   
   DoEvents
   
   On Error GoTo 0
   Exit Sub

carga_grid_Error:
   
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_grid de Formulario frmFlexArre"
 
 End Sub





Private Sub chameleonButton1_Click()

Dim linea1 As String
Dim linea2 As String
Dim tmpcaja As String
         

   On Error GoTo chameleonButton1_Click_Error

    DoEvents

    If cbCAJAS.Text <> "" Then
        tmpcaja = devuelve_campo("SELECT DESCRIPCION FROM CAJAS WHERE CODIGO = " & cbCAJAS.Text, locCnn)
        If tmpcaja = "@" Then tmpcaja = ""
    End If
    
    linea1 = "Informe de Pagos. F.Inicial: " & ioFECHAINI.Text & ". F.Final: " & ioFECHAFIN.Text & ". Caja: " & tmpcaja
    linea2 = "Usuario: " & Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & UsuarioActual, locCnn)) & ". Fecha: " & Now
    
    Call PrintFlexGrid(fg, 1, 1, 2, linea1, linea2, 13, 2)
    'fg.SaveGrid "c:\prueba.txt", flexFileCommaText
    'fg.PrintGrid "Transferencia " & rc.Fields("CODIGO"), , 2, 600, 600
    
    
    'actualizar los colores del grid, volivendolo a cargar
    'rc.Move 0




   On Error GoTo 0
   Exit Sub

chameleonButton1_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento chameleonButton1_Click de Formulario frmFlexPagPro"

End Sub

Private Sub fg_dblClick()

'If miRc.State = 0 Then Exit Sub
'If miRc.RecordCount <= 0 Then Exit Sub

  '  seleccionado = True
    
'If fg.Rows <= 1 Then Exit Sub
    
      ' If IsNumeric(fg.TextMatrix(fg.Row, 0)) Then
        'posicionarse en el registro
            'miRc.Move (fg.TextMatrix(fg.Row, 0) - 1), 1
           '
          ' miRc.Close
         '
        '   DoEvents
           '
          ' Unload Me
         '
        '   Exit Sub
          
                                 
       ' End If
        
        
                
'End If
  
End Sub

Private Sub fg_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

Case 13
    
    Call fg_dblClick
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
            cbTIPOPAGO.SetFocus
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
  
With cbCODPROV
    .ConexionString = locCnn
    .LenCodigo = 5
    .SQLString = "SELECT CODIGO, NOMBRE FROM MAPROV WHERE MBAJA = 0 ORDER BY CODIGO"
    .CodigoWidth = 800
    .carga
    DoEvents
End With


 With cbTIPOPAGO

    .ConexionString = locCnn
    .LenCodigo = 5
    
    If Pago_A_Proveedor Then
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM MAPAG WHERE (MBAJA = 0) AND (CODIGO=1) ORDER BY CODIGO"
    Else
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM MAPAG WHERE (MBAJA = 0) AND (CODIGO>1) ORDER BY CODIGO"
    End If
    
    .CodigoWidth = 800
    .carga
    DoEvents
    
    If Pago_A_Proveedor Then
        .Text = 1   'fijar en pagos a proveedores
        .Enabled = False
    Else
        .Text = ""
    End If
    
End With
  

'0=PENDIENTE, 1=ACEPTADA, 2=CANCELADA
With cbESTADO
    .añade_item "0  PENDIENTE"
    .añade_item "1  PARCIAL"
    .añade_item "2  PAGADO"
    .añade_item "3  ANULADO"
    .LenCodigo = 1
    .CodigoWidth = 300
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
  
' artsql.AddTable "MAARTIC"
' artsql.AddField "(CONVERT(char(10), CODIGO) + CONVERT(char(3), TEMPOR))"
 
 DetPagSQL.AddTable "DETPAGOS"
 
 'DetPagSQL.AddField "CONVERT(char(10), CODART) + CONVERT(char(3), TEMPOR) as COD"
 'DetPagSQL.AddField "CONVERT(char(10), CODVEN) + CONVERT(char(3), CODCAJA) as CODVENTA"

 DetPagSQL.AddField "CODIGO"
 DetPagSQL.AddField "CODCAJA"
 DetPagSQL.AddField "LINEA"
 DetPagSQL.AddField "IMPORTE"
 DetPagSQL.AddField "FMODI"
 DetPagSQL.AddField "MBAJA"
 
 DetPagSQL.AddOrderClause "CODCAJA"
 DetPagSQL.AddOrderClause "CODIGO"
 DetPagSQL.AddOrderClause "LINEA"
 
 
 CabPagSQL.AddTable "PAGOS"
 CabPagSQL.AddField "CONVERT(char(10), CODIGO) + CONVERT(char(3), CODCAJA)"
 
  
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
    
    Set frmFlexPagPro = Nothing
    
        
End Sub



