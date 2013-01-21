VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmTotTransf 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Totales de Transferencias"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11595
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
   ScaleHeight     =   1605
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   Begin PCGestion.miCombo cbCODALMORIG 
      Height          =   495
      Left            =   915
      TabIndex        =   0
      Top             =   45
      Width           =   3990
      _ExtentX        =   8202
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
   Begin PCGestion.miText ioFECHAINI 
      Height          =   480
      Left            =   5910
      TabIndex        =   2
      Top             =   75
      Width           =   1380
      _ExtentX        =   2381
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
   Begin PCGestion.miText ioFECHAFIN 
      Height          =   480
      Left            =   8040
      TabIndex        =   3
      Top             =   90
      Width           =   1365
      _ExtentX        =   2487
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
   Begin PCGestion.chameleonButton cbAceptar 
      Height          =   420
      Left            =   3930
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1125
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   741
      BTYPE           =   3
      TX              =   "&Consultar"
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
      FCOL            =   16776960
      FCOLO           =   16776960
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTotFact.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cmSalir 
      Height          =   420
      Left            =   5820
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1125
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   741
      BTYPE           =   3
      TX              =   "&Salir"
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
      FCOL            =   16776960
      FCOLO           =   16776960
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTotFact.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.miCombo cbCODALMDEST 
      Height          =   495
      Left            =   915
      TabIndex        =   8
      Top             =   525
      Width           =   4005
      _ExtentX        =   8202
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
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   4710
      Left            =   45
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1875
      Visible         =   0   'False
      Width           =   11355
      _cx             =   20029
      _cy             =   8308
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
      FormatString    =   $"frmTotFact.frx":0038
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
   Begin PCGestion.miCombo cbCODPROV 
      Height          =   495
      Left            =   5895
      TabIndex        =   11
      Top             =   585
      Width           =   3615
      _ExtentX        =   6376
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
   Begin PCGestion.miCombo cbTEMPOR 
      Height          =   480
      Left            =   9975
      TabIndex        =   14
      Top             =   75
      Width           =   1620
      _ExtentX        =   2937
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
   Begin PCGestion.miCombo cbEN 
      Height          =   525
      Left            =   9960
      TabIndex        =   12
      Top             =   570
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
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "EN"
      Height          =   330
      Left            =   9525
      TabIndex        =   16
      Top             =   660
      Width           =   405
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TEMP."
      Height          =   285
      Left            =   9360
      TabIndex        =   15
      Top             =   180
      Width           =   675
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PROV."
      Height          =   300
      Left            =   5115
      TabIndex        =   13
      Top             =   660
      Width           =   825
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DESTINO"
      Height          =   330
      Left            =   -15
      TabIndex        =   9
      Top             =   615
      Width           =   915
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "F.INICIAL"
      Height          =   285
      Left            =   4935
      TabIndex        =   5
      Top             =   150
      Width           =   990
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "F.FINAL"
      Height          =   285
      Left            =   7230
      TabIndex        =   4
      Top             =   165
      Width           =   825
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ORIGEN"
      Height          =   300
      Left            =   30
      TabIndex        =   1
      Top             =   135
      Width           =   825
   End
End
Attribute VB_Name = "frmTotTransf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Módulo     : frmTotPed
' Fecha/Hora : 05/07/2004 13:02
' Autor      : JCastillo
' Propósito  :  Mostrar los totales de pedidos
'---------------------------------------------------------------------------------------
Option Explicit

Private Sub cbAceptar_Click()
Dim rc As New ADODB.Recordset
Dim oSQL As New clsSmartSQL
Dim dSQL As New clsSmartSQL
Dim usa_det As Boolean

Dim cr As New clsCrystalFormula
Dim fecha_desde As String
Dim fecha_hasta As String

   On Error GoTo cbAceptar_Click_Error


'filtro para crear la formula de crystal ...
If cbCODALMORIG.Text <> "" Then
    cr.AñadeCampo "PTRANS", "CODALMORIG", CInt(cbCODALMORIG.Text), "=", "AND"
End If

If cbCODALMDEST.Text <> "" Then
    cr.AñadeCampo "PTRANS", "CODALMDEST", CInt(cbCODALMDEST.Text), "=", "AND"
End If

If ioFECHAINI.Text <> "" And ioFECHAFIN.Text <> "" Then
    cr.AñadeCampo "PTRANS", "FMODI", ioFECHAINI.Text, ">=", "AND", True
    cr.AñadeCampo "PTRANS", "FMODI", ioFECHAFIN.Text, "<=", "AND", True
End If

If cbTEMPOR.Text <> "" Then
    cr.AñadeCampo "DETTRANS", "TEMPOR", CInt(cbTEMPOR.Text), "=", "AND"
End If

If cbCODPROV.Text <> "" Then
    cr.AñadeCampo "MAARTIC", "CODPROV", CLng(cbCODPROV.Text), "=", "AND"
End If


If cbEN.Text <> "" Then
    
    If cbEN.Text = 0 Then   'si es caja a iva compra > 0
        cr.AñadeCampo "MAARTIC", "IVACOM", 0, ">", "AND"
    Else                      'si es caja b iva compra = 0
        cr.AñadeCampo "MAARTIC", "IVACOM", 0, "=", "AND"
    End If
    
End If

Debug.Print cr.formula

Call procesa_informes(2, False, cr.formula)

Exit Sub

  With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
  End With

oSQL.AddTable "PTRANS"

oSQL.AddField "CODIGO"
oSQL.AddField "CODALMORIG"
oSQL.AddField "CODALMDEST"
oSQL.AddField "FMODI"
oSQL.AddField "TOTAL"
oSQL.AddField "GASTOS"
oSQL.AddField "DCTO"
oSQL.AddField "NUMPED"

oSQL.AddOrderClause "CODIGO"

'estados a
oSQL.AddComplexWhereClause "ESTADO IN (1,2)", LOGIC_AND

dSQL.AddTable "DETTRANS"
dSQL.AddField "CAST(CODIGO AS CHAR(10)) + CAST(CODALM AS CHAR(3))"
    
'si busca entre fechas
If ioFECHAINI.Text <> "" And ioFECHAFIN.Text <> "" Then
    oSQL.AddComplexWhereClause "(FMODI >= '" & Format(Year((ioFECHAINI.Text)), "0000") & Format(Month((ioFECHAINI.Text)), "00") & Format(Day((ioFECHAINI.Text)), "00") & "' AND FMODI <= '" & Format(Year((ioFECHAFIN.Text)), "0000") & Format(Month((ioFECHAFIN.Text)), "00") & Format(Day((ioFECHAFIN.Text)), "00") & "')", LOGIC_AND
End If

'con almacen de origen
If (cbCODALMORIG.Text <> "" And cbTEMPOR.Text = "") Then
    oSQL.AddSimpleWhereClause "CODALMORIG", CByte(cbCODALMORIG.Text), , , LOGIC_AND
End If

'con proveedor
If cbCODALMDEST.Text <> "" Then
    oSQL.AddSimpleWhereClause "CODALMDEST", CLng(cbCODALMDEST.Text), , , LOGIC_AND
End If


If (cbTEMPOR.Text <> "") Then
    
     If cbCODALMORIG.Text <> "" Then
            dSQL.AddSimpleWhereClause "CODALM", CLng(cbCODALMORIG.Text), , , LOGIC_AND
    End If
    
    dSQL.AddSimpleWhereClause "TEMPOR", CLng(cbTEMPOR.Text), , , LOGIC_AND
    
    'incluir el DISTINCTROW a guevo para que no duplique numeros de pedido
    oSQL.AddComplexWhereClause "(CODIGO IN(" & Left(dSQL.SQL, 6) & " DISTINCT " & Right(dSQL.SQL, Len(dSQL.SQL) - 6) & "))", LOGIC_AND
    
    usa_det = True

End If


If cbCODPROV.Text <> "" Then

   dSQL.AddTable "MAARTIC"
   dSQL.AddField "CODPROV", "MAARTIC"
   dSQL.SetupJoin "DETTRANS", "CAST(CODART AS CHAR(10)) + CAST(TEMPOR AS CHAR(3)))", "MAARTIC", "(CAST(CODIGO AS CHAR(10)) + CAST(TEMPOR AS CHAR(3))", CLAUSE_DOESNOTEQUAL, INNER_JOIN
    
   dSQL.AddSimpleWhereClause "CODPROV", CLng(cbCODPROV.Text), , , LOGIC_AND
   
   usa_det = True
 
End If


'si usa detalle, enlazar con el sql para el detalle
If usa_det Then
    oSQL.AddComplexWhereClause "CAST(CODIGO AS CHAR(10)) + CAST(CODALMORIG AS CHAR(3)) IN (" & dSQL.SQL & ")"
End If

Debug.Print oSQL.SQL

rc.Open oSQL.SQL, locCnn, adOpenStatic, adLockReadOnly


With fg
    .Clear
    .Rows = 1
    .Cols = 9
    
    .TextMatrix(0, 1) = "CODIGO"
    .TextMatrix(0, 2) = "ORIGEN"
    .TextMatrix(0, 3) = "DESTINO"
    .TextMatrix(0, 4) = "FECHA"
    .TextMatrix(0, 5) = "GASTOS"
    .TextMatrix(0, 6) = "DCTO."
    .TextMatrix(0, 7) = "TOTAL"
    .TextMatrix(0, 8) = "PEDIDO"
        
    Do Until rc.EOF

        .Rows = .Rows + 1
        
        .TextMatrix(.Rows - 1, 1) = rc.fields("CODIGO")
        .TextMatrix(.Rows - 1, 2) = Trim(devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & rc.fields("CODALMORIG"), locCnn))
        .TextMatrix(.Rows - 1, 3) = Trim(devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & rc.fields("CODALMDEST"), locCnn))
        .TextMatrix(.Rows - 1, 4) = rc.fields("FMODI")
        .TextMatrix(.Rows - 1, 5) = rc.fields("GASTOS")
        .TextMatrix(.Rows - 1, 6) = rc.fields("DCTO")
        .TextMatrix(.Rows - 1, 7) = rc.fields("TOTAL")
        .TextMatrix(.Rows - 1, 8) = rc.fields("NUMPED")
         
        rc.MoveNext

    Loop
    
    
    If .Rows > 1 Then
    .SubtotalPosition = flexSTAbove
    
    .subtotal flexSTSum, , 5, "Currency", vbBlue, vbWhite, True
    .subtotal flexSTSum, , 6, "Currency", vbBlue, vbWhite, True
    .subtotal flexSTSum, , 7, "Currency", vbBlue, vbWhite, True
    
    .TextMatrix(1, 1) = ""
    End If
    
    .AutoSize 1, .Cols - 1

End With

rc.Close
Set rc = Nothing
Set oSQL = Nothing
Set dSQL = Nothing

   On Error GoTo 0
   Exit Sub

cbAceptar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbAceptar_Click de Formulario frmTotPed"

End Sub




Private Sub cmSalir_Click()

    Unload Me

End Sub

Private Sub Form_Load()

Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2

  
With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
End With

fg.Clear
fg.Rows = 1

With cbCODALMORIG
      .LenCodigo = 4
      .CodigoWidth = 670
      .ConexionString = locCnn
      .SQLString = "SELECT CODIGO, DESCRIPCION FROM ALMACENES WHERE MBAJA = 0 ORDER BY CODIGO"
      .carga
      DoEvents
      .Text = AlmacenActual
End With

'CAJA DE DESTINO (0=A, 1=B)
With cbEN
    .añade_item "0  A"
    .añade_item "1  B"
    .LenCodigo = 1
    .CodigoWidth = 300
End With

With cbCODALMDEST
      .LenCodigo = 4
      .CodigoWidth = 670
      .ConexionString = locCnn
      .SQLString = "SELECT CODIGO, DESCRIPCION FROM ALMACENES WHERE MBAJA = 0 ORDER BY CODIGO"
      .carga
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

With cbCODPROV
      .LenCodigo = 6
      .CodigoWidth = 800
      .ConexionString = locCnn
      .SQLString = "SELECT CODIGO, NOMBRE FROM MAPROV WHERE MBAJA = 0 ORDER BY CODIGO"
      .carga
      DoEvents
End With
  
With cbTEMPOR
    .ConexionString = locCnn
    .SQLString = "SELECT IDTEM, ABREVIA FROM TEMPOR WHERE MBAJA = 0 ORDER BY ACTUAL DESC, IDTEM"
    .LenCodigo = 3
    .CodigoWidth = 500
    .DataField = "TEMPOR"
    .carga
    .Text = TemporadaActual
End With


End Sub

