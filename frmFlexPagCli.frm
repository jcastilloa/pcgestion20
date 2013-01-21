VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmFlexPagCli 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Deudas de Clientes ..."
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11475
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
   ScaleHeight     =   6915
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin PCGestion.miCombo cbCAJAS 
      Height          =   495
      Left            =   6562
      TabIndex        =   2
      Top             =   -15
      Width           =   4140
      _ExtentX        =   7303
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
      Height          =   450
      Left            =   2227
      TabIndex        =   0
      Top             =   0
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   794
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
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   6015
      Left            =   15
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   870
      Visible         =   0   'False
      Width           =   11430
      _cx             =   20161
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
      FormatString    =   $"frmFlexPagCli.frx":0000
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
   Begin PCGestion.ucGrdBttn cbLista 
      Height          =   405
      Left            =   9420
      TabIndex        =   4
      Top             =   465
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   714
      Caption         =   "&Consultar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "frmFlexPagCli.frx":00DE
   End
   Begin PCGestion.ucGrdBttn cbBorrar 
      Height          =   405
      Left            =   10695
      TabIndex        =   5
      Top             =   465
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   714
      Caption         =   "&Borrar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "frmFlexPagCli.frx":00FA
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel4 
      Height          =   345
      Left            =   4680
      Top             =   495
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   609
      Caption         =   "-F4- Consultar   -F5- Ir a Rejilla    -F8- Salir  -C- Asig. Cliente"
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
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
   Begin PCGestion.miText ioFECHAFIN 
      Height          =   450
      Left            =   4687
      TabIndex        =   1
      Top             =   0
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   794
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
   Begin PCGestion.bsGradientLabel lblCliente 
      Height          =   345
      Left            =   30
      Top             =   495
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   609
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA FIN"
      Height          =   330
      Left            =   3592
      TabIndex        =   8
      Top             =   75
      Width           =   1065
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CAJA"
      Height          =   330
      Left            =   6007
      TabIndex        =   6
      Top             =   60
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA INICIO"
      Height          =   330
      Left            =   772
      TabIndex        =   7
      Top             =   60
      Width           =   1425
   End
End
Attribute VB_Name = "frmFlexPagCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim first As Boolean

Dim tmpstrcombo As String

Public miosql As New clsSmartSQL

Public miRc As New ADODB.Recordset
Public seleccionado As Boolean

'CODIGO DE CLIENTE -----------------
Dim Codigo_Cliente As Long
Dim Caja_Cliente As Byte
'-----------------------------------


'Dim nif As New clsNIF

Private Sub cbBorrar_click()

ioFECHAINI.Text = Date
ioFECHAFIN.Text = ""
cbCAJAS.Text = CajaActual

fg.Clear
fg.Rows = 1

'Call cbLista_click

End Sub

Private Sub cbCAJAS_Validate(Cancel As Boolean)
Call cbLista_click
End Sub

Private Sub cbLista_click()
Dim usa_where As Boolean

   On Error GoTo cbLista_click_Error

miosql.ClearWhereClause

If ((ioFECHAINI.Text <> "") And IsDate(ioFECHAINI.Text)) And ((ioFECHAFIN.Text <> "") And IsDate(ioFECHAFIN.Text)) Then

    'miosql.AddSimpleWhereClause "FALTA", ioFECHAINI.Text, , CLAUSE_GREATERTHANOREQUAL
    'miosql.AddSimpleWhereClause "FALTA", CStr(DateAdd("d", 1, ioFECHAINI.Text)), , CLAUSE_LESSTHAN, LOGIC_AND
    'miOsql.AddComplexWhereClause "Year(FALTA IN (" & masql.SQL & ")", LOGIC_AND
    
    '>= q el dia actual
    '< que el dia siguiente
    miosql.AddComplexWhereClause "FECIERRE >= '" & Format(ioFECHAINI.Text, "yyyymmdd") & "' AND FECIERRE <= '" & Format(ioFECHAFIN.Text, "yyyymmdd") & "'", LOGIC_AND
    usa_where = True
         
'solo buscar por fecha inicial
ElseIf ((ioFECHAINI.Text <> "") And IsDate(ioFECHAINI.Text)) And (ioFECHAFIN.Text = "") Then
         
    miosql.AddSimpleWhereClause "FECIERRE", Format(ioFECHAINI.Text, "yyyymmdd")
    usa_where = True
    'miosql.AddComplexWhereClause "FECIERRE = '" & Format(ioFECHAINI.Text, "mm/dd/yyyy") & "'", LOGIC_AND
         
End If

If cbCAJAS.Text <> "" Then
    miosql.AddSimpleWhereClause "CODCAJA", cbCAJAS.Text, , , LOGIC_AND
    usa_where = True
End If

'If cbESTADO.Text <> "" Then
'    miosql.AddSimpleWhereClause "ESTADO", cbESTADO.Text, , , LOGIC_AND
'    usa_where = True
'End If

'si deja todo en blanco, no mostrar ningun registro
If Not usa_where Then
    fg.Clear
    Exit Sub
End If


If miRc.State = 1 Then miRc.Close
miRc.Open miosql.SQL, locCnn, adOpenStatic, adLockOptimistic

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


   On Error GoTo 0
   Exit Sub

cbLista_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbLista_click de Formulario frmFlexArre"

End Sub


'---------------------------------------------------------------------------------------
' Procedimiento : carga_grid
' Fecha/Hora    : 26/01/2004 09:59
' Autor         : JCastillo
' Propósito     :   Cargar el grid con los arreglos
'---------------------------------------------------------------------------------------
'MODELO
'TALLA
'COLOR
'DESCRIPCION
'COSTURERA
'PVP
'USUARIO
'CAJA
'ESTADO 0=PENDIENTE, 1=SERVIDO, 2=CANCELADO
'---------------------------------------------------------------------------------------
Private Sub carga_grid()
Dim tmpcodcolor As Long
Dim Total_Efectivo As Double
Dim Saldo_Caja_Efectivo As Double
Dim Ventas_Netas As Double
Dim Cobros_Realizados As Double

   On Error GoTo carga_grid_Error

   With fg
   
    .Clear
    .Cols = 33
    
    .ColFormat(5) = "Currency"
    .ColFormat(7) = "Currency"
    .ColFormat(10) = "Currency"
    .ColFormat(12) = "Currency"
    .ColFormat(14) = "Currency"
    .ColFormat(16) = "Currency"
    .ColFormat(18) = "Currency"
    .ColFormat(20) = "Currency"
    .ColFormat(21) = "Currency"
    .ColFormat(22) = "Currency"
    .ColFormat(23) = "Currency"
    .ColFormat(24) = "Currency"
    .ColFormat(25) = "Currency"
    .ColFormat(26) = "Currency"
    .ColFormat(27) = "Currency"
    .ColFormat(28) = "Currency"
    .ColFormat(29) = "Currency"
    .ColFormat(30) = "Currency"
    .ColFormat(31) = "Currency"
    
    .ColHidden(0) = True
    .ColHidden(1) = True
    .Rows = 1
    
    .TextMatrix(0, 2) = "Codigo"
    .TextMatrix(0, 3) = "Caja"
    .TextMatrix(0, 4) = "Fecha"
    .TextMatrix(0, 5) = "Contado"
    
    .TextMatrix(0, 6) = "Nº Vales Ac"
    .TextMatrix(0, 7) = "Vales Ac"
    
    .TextMatrix(0, 8) = "Nº Vales Em"
    .TextMatrix(0, 9) = "Vales Em"
    
    .TextMatrix(0, 10) = "Nº Tarj."
    .TextMatrix(0, 11) = "Tarjetas"
    .TextMatrix(0, 12) = "Nº Devol."
    .TextMatrix(0, 13) = "Devolu."
    .TextMatrix(0, 14) = "Nº Arre."
    .TextMatrix(0, 15) = "Arreglos"
    .TextMatrix(0, 16) = "Nº Movi."
    .TextMatrix(0, 17) = "Movimien."
    .TextMatrix(0, 18) = "Nº Pagos"
    .TextMatrix(0, 19) = "Pagos"
    .TextMatrix(0, 20) = "Nº Deud.C."
    .TextMatrix(0, 21) = "Deudas Cli."
    .TextMatrix(0, 22) = "En Caja (Fic.)"
    .TextMatrix(0, 23) = "Total A"
    .TextMatrix(0, 24) = "Total B"
    .TextMatrix(0, 25) = "TOTAL"
    .TextMatrix(0, 26) = "En Caja (Real)"
    .TextMatrix(0, 27) = "Descuadre"
    .TextMatrix(0, 28) = "Cobros Real."
    .TextMatrix(0, 29) = "Ventas Netas"
    .TextMatrix(0, 30) = "Saldo Caja"
    .TextMatrix(0, 31) = "Total Efectivo"
    
    .TextMatrix(0, 32) = "Usuario"
    
    If miRc.RecordCount <= 0 Then Exit Sub
        
    Do
             .Rows = .Rows + 1
    
        If Not miRc.EOF Then
     
         
            'ID
            '.TextMatrix(.Rows - 1, 0) = miRc.Fields("ID")
            'CAJA
            '.TextMatrix(.Rows - 1, 1) = miRc.Fields("CODIGO") & Format(miRc.Fields("CODCAJA"), "000")
            
            Total_Efectivo = ((miRc.fields("T_CONTADO") - miRc.fields("T_ARREGLOS")) + miRc.fields("T_VALE_ACEP"))
            Cobros_Realizados = Total_Efectivo + miRc.fields("T_TARJETA")
            Ventas_Netas = Total_Efectivo - miRc.fields("T_VALE_ACEP")
            Saldo_Caja_Efectivo = Total_Efectivo - miRc.fields("T_PAGOS")
            
                        
            'CODIGO
            .TextMatrix(.Rows - 1, 2) = miRc.fields("CODIGO")
            
            'CAJA
            .TextMatrix(.Rows - 1, 3) = Trim(devuelve_campo("SELECT DESCRIPCION FROM CAJAS WHERE CODIGO = " & miRc.fields("CODCAJA"), locCnn))
            
            'FECHA
            .TextMatrix(.Rows - 1, 4) = miRc.fields("FECIERRE")
            
            'Total Contado
            .TextMatrix(.Rows - 1, 5) = miRc.fields("T_CONTADO") - miRc.fields("T_ARREGLOS")
            
            'Numero vales
            .TextMatrix(.Rows - 1, 6) = miRc.fields("NVALES_ACEP")
                           
            'Total Vales
            .TextMatrix(.Rows - 1, 7) = miRc.fields("T_VALE_ACEP")
            
            'Numero vales emitidos
            .TextMatrix(.Rows - 1, 8) = miRc.fields("NVALES_EMI")
            
            'Total Vales emitidos
            .TextMatrix(.Rows - 1, 9) = miRc.fields("T_VALE_EMI")
            
            'Num. Tarjetas
            .TextMatrix(.Rows - 1, 10) = miRc.fields("NVTAR")
                        

                         
            'Total Tarjetas
            .TextMatrix(.Rows - 1, 11) = miRc.fields("T_TARJETA")
            
             'devoluciones
            .TextMatrix(.Rows - 1, 12) = miRc.fields("NDEVOL")
            
            'Total devoluciones
            .TextMatrix(.Rows - 1, 13) = miRc.fields("T_DEVOLU")
            
            'Nº Arreglos
            .TextMatrix(.Rows - 1, 14) = miRc.fields("NARRE")
            
            'Total Arreglos
            .TextMatrix(.Rows - 1, 15) = miRc.fields("T_ARREGLOS")
                                
             'Nº Movi
            .TextMatrix(.Rows - 1, 16) = miRc.fields("N_MOVI")
            
            'Total Movimientos
            .TextMatrix(.Rows - 1, 17) = miRc.fields("T_MOVI")
            
            'Nº Pagos
            .TextMatrix(.Rows - 1, 18) = miRc.fields("N_PAGOS")
            
            'Total Pagos
            .TextMatrix(.Rows - 1, 19) = miRc.fields("T_PAGOS")
            
            'Nº Deudas Cliente
            .TextMatrix(.Rows - 1, 20) = miRc.fields("N_DEUDC")
            
            'Total Deudas Cliente
            .TextMatrix(.Rows - 1, 21) = miRc.fields("T_DEUDC")
     
            'En caja ficticio:
            'CONTADO + -  MOVIMIENTOS CAJA
            .TextMatrix(.Rows - 1, 22) = miRc.fields("T_CONTADO") + miRc.fields("T_MOVI")
            
            'total en caja a
            .TextMatrix(.Rows - 1, 23) = miRc.fields("T_CAJAA")
            
            'total en caja b
            .TextMatrix(.Rows - 1, 24) = miRc.fields("T_CAJAB")
            
            'TOTAL:
            .TextMatrix(.Rows - 1, 25) = (miRc.fields("T_CONTADO") - miRc.fields("T_ARREGLOS")) + miRc.fields("T_TARJETA")
            
            'EN caja REAL
            .TextMatrix(.Rows - 1, 26) = miRc.fields("T_ENCAJA")
            
            'Descuadre (REAL - FICTICIO)
            .TextMatrix(.Rows - 1, 27) = miRc.fields("T_ENCAJA") - .TextMatrix(.Rows - 1, 20)
     
            'cobros realizados
            .TextMatrix(.Rows - 1, 28) = Cobros_Realizados
            
            'Ventas Netas
            .TextMatrix(.Rows - 1, 29) = Ventas_Netas
            
            'Saldo de caja efectivo
            .TextMatrix(.Rows - 1, 30) = Saldo_Caja_Efectivo
            
            'Total Efectivo
            .TextMatrix(.Rows - 1, 31) = Total_Efectivo
            
            'usuario
            .TextMatrix(.Rows - 1, 32) = Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & miRc.fields("CODUSR"), locCnn))
     
        End If
    
    If Not miRc.EOF Then miRc.MoveNext
    
    Loop Until miRc.EOF
          
        .SubtotalPosition = flexSTAbove
        .subtotal flexSTSum, , 5, , vbBlue, vbWhite, True
        .subtotal flexSTSum, , 7, , vbBlue, vbWhite, True
        .subtotal flexSTSum, , 8, , vbBlue, vbWhite, True
        .subtotal flexSTSum, , 9, , vbBlue, vbWhite, True
        .subtotal flexSTSum, , 11, , vbBlue, vbWhite, True
        .subtotal flexSTSum, , 13, , vbBlue, vbWhite, True
        .subtotal flexSTSum, , 15, , vbBlue, vbWhite, True
        .subtotal flexSTSum, , 17, , vbBlue, vbWhite, True
        .subtotal flexSTSum, , 19, , vbBlue, vbWhite, True
        .subtotal flexSTSum, , 20, , vbBlue, vbGreen, True
        .subtotal flexSTSum, , 21, , vbBlue, vbGreen, True
        .subtotal flexSTSum, , 22, , vbBlue, vbGreen, True
        .subtotal flexSTSum, , 23, , vbBlue, vbGreen, True
        .subtotal flexSTSum, , 24, , vbBlue, vbGreen, True
        .subtotal flexSTSum, , 25, , vbBlue, vbGreen, True
        .subtotal flexSTSum, , 26, , vbBlue, vbGreen, True
        .subtotal flexSTSum, , 27, , vbBlue, vbGreen, True
        .subtotal flexSTSum, , 28, , vbBlue, vbGreen, True
        .subtotal flexSTSum, , 29, , vbBlue, vbGreen, True
        .subtotal flexSTSum, , 30, , vbBlue, vbGreen, True
        .subtotal flexSTSum, , 31, , vbBlue, vbGreen, True
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 3) = "Totales:"
        
    .AutoSize 1, .Cols - 1
    .Redraw = True

  End With
  
  
  
   On Error GoTo 0
   Exit Sub

carga_grid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_grid de Formulario frmFlexArre"
 
 End Sub


Private Sub fg_DblClick()
    seleccionado = True

If miRc.State = 0 Then Exit Sub
If miRc.RecordCount <= 0 Then Exit Sub

'si corresponde a algun ID
If fg.TextMatrix(fg.Row, 0) <> "" Then

With frmNuArr

    .Solo_Actualizar = True
    .Sel_Caja = fg.TextMatrix(fg.Row, 1)
    .Sel_ID = fg.TextMatrix(fg.Row, 0)
    
    .ioNOMBRE.Text = fg.TextMatrix(fg.Row, 7)
    .ioDescripcion.Text = fg.TextMatrix(fg.Row, 6)
    .ioPVP.Text = fg.TextMatrix(fg.Row, 8)
    
    Select Case fg.TextMatrix(fg.Row, 11)
            
            Case "PENDIENTE"
            
                    .cbESTADO.Text = 1
            Case "SERVIDO"
            
                    .cbESTADO.Text = 2
            Case "CANCELADO"
            
                    .cbESTADO.Text = 3
                      
    End Select
            
    '.cbESTADO.Text = fg.TextMatrix(fg.Row, 6)
    
    .Show 1
    Call cbLista_click
    
End With
    
End If
    
   ' Unload Me
End Sub

Private Sub fg_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

Case 13
    
    Call fg_DblClick
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
        
    If Codigo_Cliente = 0 Then
    
        
        Do
        
            Call Abre_Grid_Clientes
    
        Loop Until Codigo_Cliente > 0
    
    End If
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode

'Ir al grid, o regresar
Case vbKeyF5
    
    If fg.Rows > 1 Then
        If fg.TabStop Then
            fg.TabStop = False
            ioFECHAINI.SetFocus
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

    Call cbLista_click
    KeyCode = 0
    
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
  End With
  
  With ioFECHAINI
    .dspFormat = "dd/mm/yyyy"
   .LongMaxima = 10
  End With
  
  With ioFECHAFIN
    .dspFormat = "dd/mm/yyyy"
   .LongMaxima = 10
  End With
  
 '   With ioIMPORTE
 '   .dspFormat = "Currency"
  ' .LongMaxima = 10
 '  .Alineacion = 1
 ' End With
  
 'With ioMODELO
 '  .LongMaxima = 30
 'End With
  
 'With ioNOMBRE
 '    .LongMaxima = 50
 'End With
  
 'With cbESTADO
 '   .añade_item "1   - PENDIENTE"
 '   .añade_item "2   - SERVIDO"
  '  .añade_item "3   - CANCELADO"
  '  .LenCodigo = 1
 '   .CodigoWidth = 300
 '   .Text = "1"
 'End With
  
 'artsql.AddTable "MAARTIC"
 'artsql.AddField "CODIGO"
 'masql.AddTable "COSTURE"
 miosql.AddTable "CIERREDIA"
 'masql.AddField "CODIGO"
 
ioFECHAINI.Text = Date
ioFECHAFIN.Text = ""
cbCAJAS.Text = CajaActual

Select Case TipoPermiso
Case 0
    cbCAJAS.Locked = True
'Case 1
'    cbCAJAS.Locked = False
End Select
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    tmpstrcombo = ""
    'Set nif = Nothing
        
    'si hemos establecido un filtro que no devuelve ningun registro,
    'borrar filtro para que no de error al volver al formulario
    'If miRc.EOF Then Call cbBorrar_click
    
    'No descargar desde aqui, descargar desde el formulario desde donde
    'se llame
    Set frmFlexCie = Nothing
End Sub



Private Sub ioFECHAINI_Validate(Cancel As Boolean)

If ioFECHAINI.Text <> "" Then Call cbLista_click

End Sub


Private Sub ioFECHAFIN_Validate(Cancel As Boolean)

If ioFECHAFIN.Text <> "" Then Call cbLista_click

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

