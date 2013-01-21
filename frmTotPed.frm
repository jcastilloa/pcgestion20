VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmTotPed 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Totales de Pedidos"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11415
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
   ScaleHeight     =   2865
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   Begin PCGestion.miCombo cbCODALMORIG 
      Height          =   495
      Left            =   1245
      TabIndex        =   0
      Top             =   45
      Width           =   5130
      _ExtentX        =   9049
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
   Begin PCGestion.miCombo cbEN 
      Height          =   495
      Left            =   7170
      TabIndex        =   2
      Top             =   555
      Width           =   1620
      _ExtentX        =   2858
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
      Left            =   7395
      TabIndex        =   4
      Top             =   75
      Width           =   1425
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
      Left            =   9930
      TabIndex        =   5
      Top             =   90
      Width           =   1410
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
   Begin PCGestion.chameleonButton cmSalir 
      Height          =   420
      Left            =   5610
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2400
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
      MICON           =   "frmTotPed.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.miCombo cbCODPROV 
      Height          =   495
      Left            =   1245
      TabIndex        =   9
      Top             =   540
      Width           =   5130
      _ExtentX        =   9049
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
      Left            =   9540
      TabIndex        =   11
      Top             =   570
      Width           =   1785
      _ExtentX        =   3149
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
   Begin PCGestion.chameleonButton cbAceptar 
      Height          =   420
      Left            =   3720
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2400
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
      MICON           =   "frmTotPed.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   3720
      TabIndex        =   18
      Top             =   1635
      Width           =   3750
      Begin MSForms.OptionButton optVerDetalle 
         Height          =   405
         Left            =   2160
         TabIndex        =   20
         Top             =   240
         Width           =   1470
         BackColor       =   12632256
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "2593;714"
         Value           =   "0"
         Caption         =   "Ver Detalle"
         FontName        =   "Trebuchet MS"
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton opTotal 
         Height          =   465
         Left            =   210
         TabIndex        =   19
         Top             =   210
         Width           =   1545
         BackColor       =   12632256
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "2725;820"
         Value           =   "1"
         Caption         =   "Solo Totales"
         FontName        =   "Trebuchet MS"
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin PCGestion.miText ioFACTURA 
      Height          =   480
      Left            =   1245
      TabIndex        =   12
      Top             =   1065
      Width           =   1425
      _ExtentX        =   2514
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
   Begin PCGestion.miText ioSUCODIGO 
      Height          =   480
      Left            =   3720
      TabIndex        =   13
      Top             =   1065
      Width           =   1425
      _ExtentX        =   2514
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
   Begin PCGestion.miText ioALBARAN 
      Height          =   480
      Left            =   6225
      TabIndex        =   15
      Top             =   1065
      Width           =   1425
      _ExtentX        =   2514
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
   Begin PCGestion.miText ioTRANSPORTI 
      Height          =   480
      Left            =   9090
      TabIndex        =   14
      Top             =   1065
      Width           =   2250
      _ExtentX        =   3757
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
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "FACTURA"
      Height          =   285
      Left            =   150
      TabIndex        =   24
      Top             =   1110
      Width           =   1035
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA FACTURA"
      Height          =   660
      Left            =   2625
      TabIndex        =   23
      Top             =   990
      Width           =   1035
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ALBARAN"
      Height          =   285
      Left            =   5085
      TabIndex        =   22
      Top             =   1110
      Width           =   1035
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSPORTE"
      Height          =   285
      Left            =   7695
      TabIndex        =   21
      Top             =   1155
      Width           =   1350
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TEMP."
      Height          =   285
      Left            =   8850
      TabIndex        =   16
      Top             =   675
      Width           =   675
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PROVEEDOR"
      Height          =   330
      Left            =   30
      TabIndex        =   10
      Top             =   630
      Width           =   1200
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "F.INICIAL"
      Height          =   285
      Left            =   6360
      TabIndex        =   7
      Top             =   165
      Width           =   1035
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "F.FINAL"
      Height          =   285
      Left            =   9045
      TabIndex        =   6
      Top             =   165
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "EN"
      Height          =   270
      Left            =   6735
      TabIndex        =   3
      Top             =   660
      Width           =   330
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ALMACEN"
      Height          =   300
      Left            =   195
      TabIndex        =   1
      Top             =   165
      Width           =   990
   End
End
Attribute VB_Name = "frmTotPed"
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

'
'Private Sub cbAceptar_Vieja_Click()
'Dim rc As New ADODB.Recordset
'Dim oSQL As New clsSmartSQL
'Dim dSQL As New clsSmartSQL
'
'   On Error GoTo cbAceptar_Click_Error
'
'  With locCnn
'    If .State = 0 Then
'        .CursorLocation = adUseClient
'        .Open strLocCnn
'    End If
'  End With
'
'oSQL.AddTable "CABPEDPRO"
'
'oSQL.AddField "NUMERO"
'oSQL.AddField "CODPROV"
'oSQL.AddField "FECHA"
'oSQL.AddField "TOTALNET"
'oSQL.AddField "TOTALIVA"
'oSQL.AddField "PORTES"
'oSQL.AddField "GASTOS"
'oSQL.AddField "DESTINO"
'oSQL.AddField "ALMORIG"
'oSQL.AddField "ALBARAN"
'oSQL.AddField "FACTURA"
'
'oSQL.AddOrderClause "NUMERO"
'oSQL.AddSimpleWhereClause "ESTADO", 3, , , LOGIC_AND
'
''si busca entre fechas
'If ioFECHAINI.Text <> "" And ioFECHAFIN.Text <> "" Then
'    oSQL.AddComplexWhereClause "(FECHA >= '" & Format(Year((ioFECHAINI.Text)), "0000") & Format(Month((ioFECHAINI.Text)), "00") & Format(Day((ioFECHAINI.Text)), "00") & "' AND FECHA <= '" & Format(Year((ioFECHAFIN.Text)), "0000") & Format(Month((ioFECHAFIN.Text)), "00") & Format(Day((ioFECHAFIN.Text)), "00") & "')", LOGIC_AND
'End If
'
''con almacen de origen
'If (cbCODALMORIG.Text <> "" And cbTEMPOR.Text = "") Then
'    oSQL.AddSimpleWhereClause "CODALM", CByte(cbCODALMORIG.Text), , , LOGIC_AND
'End If
'
''con proveedor
'If cbCODPROV.Text <> "" Then
'    oSQL.AddSimpleWhereClause "CODPROV", CLng(cbCODPROV.Text), , , LOGIC_AND
'End If
'
'If cbEN.Text <> "" Then
'    oSQL.AddSimpleWhereClause "DESTINO", CLng(cbEN.Text), , , LOGIC_AND
'End If
'
'If (cbTEMPOR.Text <> "") Then
'
'    If cbCODALMORIG.Text = "" Then
'        MsgBox "Para buscar por temporada es necesario establecer ALMACEN", vbInformation, titulo
'        cbCODALMORIG.SetFocus
'
'        'descargar objetos ...
'        Set oSQL = Nothing
'        Set dSQL = Nothing
'
'        Exit Sub
'    End If
'
'    dSQL.AddTable "DETPEDPRO"
'    dSQL.AddField "NUMERO"
'
'    dSQL.AddSimpleWhereClause "ALMORIG", CLng(cbCODALMORIG.Text), , , LOGIC_AND
'    dSQL.AddSimpleWhereClause "TEMPOR", CLng(cbTEMPOR.Text), , , LOGIC_AND
'
'    'incluir el DISTINCTROW a guevo para que no duplique numeros de pedido
'    oSQL.AddComplexWhereClause "(NUMERO IN(" & Left(dSQL.SQL, 6) & " DISTINCT " & Right(dSQL.SQL, Len(dSQL.SQL) - 6) & "))", LOGIC_AND
'
'
'End If
'
'
'Debug.Print oSQL.SQL
'
'rc.Open oSQL.SQL, locCnn, adOpenStatic, adLockReadOnly
'
'
'With fg
'    .Clear
'    .Rows = 1
'    .Cols = 14
'
'    .TextMatrix(0, 1) = "PEDIDO"
'    .TextMatrix(0, 2) = "PROVEEDOR"
'    .TextMatrix(0, 3) = "FECHA"
'    .TextMatrix(0, 4) = "SUBTOTAL"
'    .TextMatrix(0, 5) = "IVA"
'    .TextMatrix(0, 6) = "PORTES"
'    .TextMatrix(0, 7) = "GASTOS"
'    .TextMatrix(0, 8) = "TOTAL"
'    .TextMatrix(0, 9) = "ALMACEN"
'    .TextMatrix(0, 10) = "DESTINO"
'    .TextMatrix(0, 11) = "ALBARAN"
'    .TextMatrix(0, 12) = "FACTURA"
'    .TextMatrix(0, 13) = "TEMP"
'
'
'    Do Until rc.EOF
'
'        .Rows = .Rows + 1
'
'        .TextMatrix(.Rows - 1, 1) = rc.fields("NUMERO")
'        .TextMatrix(.Rows - 1, 2) = Trim(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & rc.fields("CODPROV"), locCnn))
'        .TextMatrix(.Rows - 1, 3) = rc.fields("FECHA")
'        .TextMatrix(.Rows - 1, 4) = rc.fields("TOTALNET")
'        .TextMatrix(.Rows - 1, 5) = rc.fields("TOTALIVA")
'        .TextMatrix(.Rows - 1, 6) = rc.fields("PORTES")
'        .TextMatrix(.Rows - 1, 7) = rc.fields("GASTOS")
'        .TextMatrix(.Rows - 1, 8) = (rc.fields("TOTALNET") + rc.fields("TOTALIVA")) + rc.fields("PORTES") + rc.fields("GASTOS")
'        .TextMatrix(.Rows - 1, 9) = Trim(devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & rc.fields("ALMORIG"), locCnn))
'
'        If rc.fields("DESTINO") = 0 Then
'        .TextMatrix(.Rows - 1, 10) = "A"
'        Else
'        .TextMatrix(.Rows - 1, 10) = "B"
'        End If
'
'        If Not IsNull(rc.fields("ALBARAN")) Then .TextMatrix(.Rows - 1, 11) = rc.fields("ALBARAN")
'        If Not IsNull(rc.fields("FACTURA")) Then .TextMatrix(.Rows - 1, 12) = rc.fields("FACTURA")
'
'        'temporada
'        .TextMatrix(.Rows - 1, 13) = Trim(devuelve_campo("SELECT ABREVIA FROM TEMPOR WHERE IDTEM = " & CLng(cbTEMPOR.Text), locCnn))
'
'        rc.MoveNext
'
'    Loop
'
'
'    If .Rows > 1 Then
'    .SubtotalPosition = flexSTAbove
'
'    .subtotal flexSTSum, , 4, "Currency", vbBlue, vbWhite, True
'    .subtotal flexSTSum, , 5, "Currency", vbBlue, vbWhite, True
'    .subtotal flexSTSum, , 6, "Currency", vbBlue, vbWhite, True
'    .subtotal flexSTSum, , 7, "Currency", vbBlue, vbWhite, True
'    .subtotal flexSTSum, , 8, "Currency", vbBlue, vbWhite, True
'
'    .TextMatrix(1, 1) = ""
'    End If
'
'    .AutoSize 1, .Cols - 1
'
'End With
'
'rc.Close
'Set rc = Nothing
'Set oSQL = Nothing
'Set dSQL = Nothing
'
'   On Error GoTo 0
'   Exit Sub
'
'cbAceptar_Click_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbAceptar_Click de Formulario frmTotPed"
'
'End Sub


Private Sub cmSalir_Click()

    Unload Me

End Sub

Private Sub cbAceptar_Click()

Dim cr As New clsCrystalFormula
Dim fecha_desde As String
Dim fecha_hasta As String

   On Error GoTo cbAceptar_Click_Error

   
'filtro para crear la formula de crystal ...
If cbCODALMORIG.Text <> "" Then
    cr.AñadeCampo "CABPEDPRO", "ALMORIG", CInt(cbCODALMORIG.Text), "=", "AND"
End If

If Trim(ioALBARAN.Text) <> "" Then
    cr.AñadeCampo "CABPEDPRO", "ALBARAN", ioALBARAN.Text, "=", "AND", , , True
End If

If Trim(ioFACTURA.Text) <> "" Then
    cr.AñadeCampo "CABPEDPRO", "FACTURA", ioFACTURA.Text, "=", "AND", , , True
End If

If Trim(ioSUCODIGO.Text) <> "" Then
    cr.AñadeCampo "CABPEDPRO", "SUCODIGO", ioSUCODIGO.Text, "=", "AND", , , True
End If


If Trim(ioTRANSPORTI.Text) <> "" Then
    cr.AñadeCampo "CABPEDPRO", "TRNSPORTI", ioTRANSPORTI.Text, "like", "AND", , , True
End If

If ioFECHAINI.Text <> "" And ioFECHAFIN.Text <> "" Then
    cr.AñadeCampo "CABPEDPRO", "FMODI", ioFECHAINI.Text, ">=", "AND", True
    cr.AñadeCampo "CABPEDPRO", "FMODI", ioFECHAFIN.Text, "<=", "AND", True
End If

If cbCODPROV.Text <> "" Then
    cr.AñadeCampo "CABPEDPRO", "CODPROV", CInt(cbCODPROV.Text), "=", "AND"
End If

If cbEN.Text <> "" Then
    cr.AñadeCampo "CABPEDPRO", "DESTINO", cbEN.Text, "=", "AND", , True
End If

If cbTEMPOR.Text <> "" Then
    cr.AñadeCampo "DETPEDPRO", "TEMPOR", CInt(cbTEMPOR.Text), "=", "AND"
End If

If opTotal.Value = True Then
    Call procesa_informes(4, False, cr.formula)
Else
    Call procesa_informes(3, False, cr.formula)
End If

   On Error GoTo 0
   Exit Sub

cbAceptar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbAceptar_Click de Formulario frmTotPed"

End Sub

Private Sub Form_Load()

Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2

  
With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
End With

'fg.Clear
'fg.Rows = 1

With cbCODALMORIG
      .LenCodigo = 4
      .CodigoWidth = 670
      .ConexionString = locCnn
      .SQLString = "SELECT CODIGO, DESCRIPCION FROM ALMACENES WHERE MBAJA = 0 ORDER BY CODIGO"
      .carga
      DoEvents
      .Text = AlmacenActual
End With

With cbCODPROV
      .LenCodigo = 6
      .CodigoWidth = 800
      .ConexionString = locCnn
      .SQLString = "SELECT CODIGO, NOMBRE FROM MAPROV WHERE MBAJA = 0 ORDER BY CODIGO"
      .carga
      DoEvents
End With

With cbEN
    .añade_item "0  A", 1
    .añade_item "1  B", 2
    .LenCodigo = 1
    .CodigoWidth = 300
End With

With ioFECHAINI
    .dspFormat = "dd/mm/yyyy"
    .LongMaxima = 10
End With

With ioSUCODIGO
    .dspFormat = "dd/mm/yyyy"
    .LongMaxima = 10
End With

With ioALBARAN
    .LongMaxima = 10
End With

With ioFACTURA
    .LongMaxima = 10
End With


With ioFECHAFIN
    .dspFormat = "dd/mm/yyyy"
    .LongMaxima = 10
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




Private Sub ioSUCODIGO_Validate(Cancel As Boolean)
If Trim(ioSUCODIGO.Text) = "" Then Exit Sub

If Not IsDate(ioSUCODIGO.Text) Then
    ioSUCODIGO.CancelarValidacion
    Cancel = True
Else
    ioSUCODIGO.Text = Format(ioSUCODIGO.Text, "dd/mm/yyyy")
End If

End Sub

