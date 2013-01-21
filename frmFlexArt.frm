VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmFlexArt 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10830
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin PCGestion.miCombo cbFAMILIA 
      Height          =   495
      Left            =   6480
      TabIndex        =   3
      Top             =   435
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   5565
      Left            =   0
      TabIndex        =   9
      Top             =   2220
      Visible         =   0   'False
      Width           =   10830
      _cx             =   19103
      _cy             =   9816
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
      FormatString    =   $"frmFlexArt.frx":0000
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
   Begin PCGestion.miText ioMODELO 
      Height          =   495
      Left            =   6495
      TabIndex        =   5
      Top             =   885
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
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
   Begin PCGestion.ucGrdBttn cbLista 
      Height          =   345
      Left            =   7680
      TabIndex        =   8
      Top             =   1860
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   609
      Caption         =   "&Consultar"
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "frmFlexArt.frx":00DE
      GradientColor1  =   16761024
      Angle           =   269
      ImageCaptionPos =   2
   End
   Begin PCGestion.ucGrdBttn cbBorrar 
      Height          =   345
      Left            =   8985
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1860
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   609
      Caption         =   "&Borrar"
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "frmFlexArt.frx":00FA
      Angle           =   269
   End
   Begin PCGestion.miCombo cbSECCION 
      Height          =   495
      Left            =   1050
      TabIndex        =   2
      Top             =   435
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PCGestion.miCombo cbCODIGO 
      Height          =   495
      Left            =   1050
      TabIndex        =   0
      Top             =   -15
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PCGestion.miCombo cbSUBFAM 
      Height          =   495
      Left            =   1050
      TabIndex        =   4
      Top             =   885
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PCGestion.miCombo cbCODPROV 
      Height          =   495
      Left            =   1050
      TabIndex        =   6
      Top             =   1320
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PCGestion.ucGrdBttn cmCerrar 
      Height          =   345
      Left            =   9870
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1860
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   609
      Caption         =   "C&errar"
      ForeColor       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "frmFlexArt.frx":0116
      GradientColor1  =   -2147483636
      GradientColor2  =   -2147483638
      Angle           =   269
      FocusColor      =   255
   End
   Begin PCGestion.miCombo cbTemporada 
      Height          =   450
      Left            =   7845
      TabIndex        =   1
      Top             =   0
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PCGestion.miText ioREF 
      Height          =   495
      Left            =   6495
      TabIndex        =   7
      Top             =   1350
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
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
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TEMPORADA"
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
      Left            =   6780
      TabIndex        =   21
      Top             =   75
      Width           =   1095
   End
   Begin MSForms.CheckBox fwbajas 
      Height          =   375
      Left            =   1020
      TabIndex        =   13
      Top             =   1785
      Width           =   1755
      VariousPropertyBits=   746588183
      BackColor       =   12632256
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "3096;661"
      Value           =   "1"
      Caption         =   "Ocultar Bajas"
      FontName        =   "Trebuchet MS"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox fwhistorico 
      Height          =   375
      Left            =   2625
      TabIndex        =   20
      Top             =   1785
      Width           =   1815
      VariousPropertyBits=   746588183
      BackColor       =   12632256
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "3201;661"
      Value           =   "1"
      Caption         =   "Ocultar Historico"
      FontName        =   "Trebuchet MS"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REF."
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
      Left            =   6090
      TabIndex        =   19
      Top             =   1440
      Width           =   405
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROVEEDOR"
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
      Left            =   0
      TabIndex        =   18
      Top             =   1410
      Width           =   1050
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MODELO"
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
      Left            =   5730
      TabIndex        =   17
      Top             =   975
      Width           =   750
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
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
      Left            =   315
      TabIndex        =   16
      Top             =   90
      Width           =   705
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SUBFAMILIA"
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
      Left            =   15
      TabIndex        =   15
      Top             =   945
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FAMILIA"
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
      Left            =   5685
      TabIndex        =   14
      Top             =   525
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SECCION"
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
      Left            =   150
      TabIndex        =   12
      Top             =   525
      Width           =   915
   End
End
Attribute VB_Name = "frmFlexArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmFlexArt
' DateTime  : 02/11/2003 02:27
' Author    : Administrador
' Purpose   : Formulario de consulta para seleccionar un artículo.
'---------------------------------------------------------------------------------------
Option Explicit

Dim first As Boolean
Dim tmprc As New ADODB.Recordset

Dim consultado As Boolean

Dim tmpcodsec As String
Dim tmpcodfam As String
Dim tmpcodsubfam As String
Public tmpcodprov As String
Dim tmpcodtempor As String
Dim tmpcodiva As String

Const Caption_Form = "Consulta Artículos ..."

Public miosql As New clsSmartSQL
Public miRc As New ADODB.Recordset
Dim nif As New clsNIF

'Para cuando se llame desde el modulo de pedidos
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public add_en_detalle As Boolean
Public rc_detalle As New ADODB.Recordset
Public NumeroPedido As Long

'para almacenar la temporada por defecto
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Temporada_Defecto As Byte

'Linea q acabamos de añadir
'Public Linea_Creada As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'para guardar la conexion activa en ese momento (loccnn o loccnnsp)
Dim miConn As New ADODB.Connection


Private Sub cbBorrar_click()

DoEvents

cbSECCION.Text = ""
cbFAMILIA.Text = ""
cbSUBFAM.Text = ""

If add_en_detalle = False Then cbCODPROV.Text = ""

cbCODIGO.Text = ""
ioMODELO.Text = ""
ioREF.Text = ""
cbTemporada.Text = ""
fwbajas.Value = True
fwhistorico.Value = True

fg.Clear
'fg.Rows = 1

'Call cbLista_click

DoEvents

Call carga_combos_inicio

consultado = False

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'---------------------------------------------------------------------------------------
' Procedure : añade_en_detalle
' DateTime  : 06/11/2003 10:35
' Author    : Administrador
' Purpose   : Añadir en el rc_detalle el registro correspondiente al ultimo
'                 movimiento.
'---------------------------------------------------------------------------------------
Private Sub añade_en_detalle()
Dim tmpcodigo As Variant

   On Error GoTo añade_en_detalle_Error


'si es un articulo de nueva creación, añadir en el momento

With rc_detalle

    .AddNew
    
    tmpcodigo = devuelve_campo("select max(LINEA) + 1 from DETPEDPRO where NUMERO = " & NumeroPedido, miConn)
    
    If tmpcodigo = "@" Then tmpcodigo = 1
    
    .fields("NUMERO") = NumeroPedido
    .fields("LINEA") = tmpcodigo
    
    'Linea_Creada = tmpcodigo
        
    .fields("CODART") = miRc.fields("CODIGO")
    .fields("TEMPOR") = miRc.fields("TEMPOR")
    .fields("PRECOM") = miRc.fields("PRECOM")
    
    'meter las unidades por defecto
     
     .fields("UNIDADES") = miRc.fields("PEDIR")
     
     'meter los descuentos
    
     .fields("ALMORIG") = AlmacenActual
     .fields("DCTO") = miRc.fields("DCTO")
     
     'meter los % para iva y re
    
     .fields("IVA").Value = devuelve_campo("SELECT IVA FROM IVA WHERE CODIGO = " & miRc.fields("TIPOIVA"), miConn)
     .fields("RE").Value = devuelve_campo("SELECT RE FROM IVA WHERE CODIGO = " & miRc.fields("TIPOIVA"), miConn)
          
     .UpdateBatch
     
     Call frmPedProv.refresca_grid_externo(True)
     frmPedProv.Linea_Creada = tmpcodigo
     Set tmpcodigo = Nothing
     
End With


   On Error GoTo 0
   Exit Sub

añade_en_detalle_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure añade_en_detalle of Formulario FrmFlexArt"

End Sub

Private Sub cbCODIGO_Validate(Cancel As Boolean)

'si codigo y temporada es <> "" entonces cargar directamente
If (cbCODIGO.Text <> "") And (cbTemporada.Text <> "") Then
    Call cbLista_click
End If

End Sub

Private Sub cbLista_click()

miosql.ClearWhereClause

If cbCODIGO.Text <> "" Then
    miosql.AddSimpleWhereClause "CODIGO", CLng(cbCODIGO.Text)
End If

If cbSECCION.Text <> "" Then
    miosql.AddSimpleWhereClause "SECCION", CLng(cbSECCION.Text)
End If



If cbFAMILIA.Text <> "" Then
    miosql.AddSimpleWhereClause "FAMILIA", CLng(cbFAMILIA.Text)
End If

If cbSUBFAM.Text <> "" Then
    miosql.AddSimpleWhereClause "SUBFAM", CLng(cbSUBFAM.Text)
End If

If cbCODPROV.Text <> "" Then
    miosql.AddSimpleWhereClause "CODPROV", CLng(cbCODPROV.Text)
End If

If cbTemporada.Text <> "" Then
    miosql.AddSimpleWhereClause "TEMPOR", CByte(cbTemporada.Text)
End If

If ioMODELO.Text <> "" Then
    miosql.AddSimpleWhereClause "MODELO", ioMODELO.Text, , CLAUSE_LIKE
End If

If ioREF.Text <> "" Then
    miosql.AddSimpleWhereClause "REF", ioREF.Text, , CLAUSE_LIKE
End If

'si decimos que ocultar bajas ...
'MBAJA = FALSE
If fwbajas.Value = True Then
    miosql.AddSimpleWhereClause "MBAJA", 0
End If

'si decide ocultar historico  ...
If fwhistorico.Value = True Then
    miosql.AddSimpleWhereClause "HIST", 0
End If

miRc.Close
miRc.Open miosql.SQL, miConn, adOpenStatic, adLockOptimistic

If miRc.RecordCount <= 0 Then

    MsgBox "No se han encontrado Artículos", vbInformation
    Call cbBorrar_click
   
    Exit Sub
    
Else


fg.Visible = False

Set fg.DataSource = miRc
DoEvents

    With fg
    
        'ocultar la columna rowguid
        .ColHidden(.Cols - 1) = True
        
        If tmpcodsec <> "" Then .ColComboList(2) = tmpcodsec
        If tmpcodfam <> "" Then .ColComboList(3) = tmpcodfam
        If tmpcodsubfam <> "" Then .ColComboList(4) = tmpcodsubfam
        If tmpcodiva <> "" Then .ColComboList(16) = tmpcodiva
        If tmpcodprov <> "" Then .ColComboList(18) = tmpcodprov
        If tmpcodtempor <> "" Then .ColComboList(19) = tmpcodtempor
        .ColFormat(1) = "00000"
        .AutoSize 1, fg.Cols - 1
    End With

'ocultar la columna comentario
fg.ColHidden(21) = True
        

fg.Visible = True

consultado = True

End If

End Sub


Private Sub cbSECCION_Validate(Cancel As Boolean)
If cbSECCION.Text <> "" Then

  'cargar codigo
With cbCODIGO
    .ConexionString = miConn
    .SQLString = "SELECT CODIGO, MODELO FROM MAARTIC WHERE MBAJA = 0 AND SECCION = " & CInt(cbSECCION.Text) & " ORDER BY CODIGO"
    .LenCodigo = 5
    .CodigoWidth = 800
    .carga
    DoEvents
  End With
  
With cbFAMILIA
    .ConexionString = miConn
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

  'cargar codigo
With cbCODIGO
    .ConexionString = miConn
    .SQLString = "SELECT CODIGO, MODELO FROM MAARTIC WHERE MBAJA = 0 AND FAMILIA = " & CInt(cbFAMILIA.Text) & " ORDER BY CODIGO"
    .LenCodigo = 5
    .CodigoWidth = 800
    .carga
    DoEvents
  End With
  
With cbSUBFAM
    .ConexionString = miConn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM SUBFAM WHERE MBAJA = 0 AND CODFAM = " & CInt(cbFAMILIA.Text) & " ORDER BY CODIGO"
    .LenCodigo = 5
    .CodigoWidth = 800
    .carga
    DoEvents
End With

End If

End Sub


Private Sub cbsubfam_Validate(Cancel As Boolean)


If cbSUBFAM.Text <> "" Then

  'cargar codigo
With cbCODIGO
    .ConexionString = miConn
    .SQLString = "SELECT CODIGO, MODELO FROM MAARTIC WHERE MBAJA = 0 AND FAMILIA = " & CInt(cbSUBFAM.Text) & " ORDER BY CODIGO"
    .LenCodigo = 5
    .CodigoWidth = 800
    .carga
    DoEvents
  End With
  
End If

End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : cbTemporada_Validate
' Fecha/Hora     : 03/12/2003 10:22
' Autor             : JCastillo
' Propósito       : Validación de temporada
'---------------------------------------------------------------------------------------
Private Sub cbTemporada_Validate(Cancel As Boolean)

   On Error GoTo cbTemporada_Validate_Error

If (cbTemporada.Text <> "") Then
  'cargar codigo
  With cbCODIGO
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, MODELO FROM MAARTIC WHERE MBAJA = 0 AND TEMPOR = " & CLng(cbTemporada.Text) & " ORDER BY CODIGO"
    .LenCodigo = 5
    .CodigoWidth = 800
    .carga
    .Refresh
  End With
  
End If

'si codigo y temporada es <> "" entonces cargar directamente
If (cbCODIGO.Text <> "") And (cbTemporada.Text <> "") Then Call cbLista_click

   On Error GoTo 0
   Exit Sub

cbTemporada_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbTemporada_Validate de Formulario frmFlexArt"
    
End Sub

Private Sub cmCerrar_Click()
Unload Me
End Sub

Private Sub fg_dblClick()

    If fg.Row > 0 Then Unload Me
    
End Sub

Private Sub fg_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

Case 13, vbKeyEscape
    KeyAscii = 0
    Unload Me
    
End Select

End Sub

Private Sub Form_Activate()
    
    Me.Caption = Caption_Form
    
    DoEvents
    
    If Not first Then
     
        DoEvents
        fg.AutoSearch = flexSearchFromCursor
        fg.ExplorerBar = flexExSortAndMove
        fg.Sort = flexSortStringAscending
        
           
        'ocultar la columna rowguid
        fg.ColHidden(fg.Cols - 1) = True
    
        'cargar strings para los colcomblist
    
        If locCnn.State <> 0 Then
            Set miConn = locCnn
        Else
            Set miConn = locCnnSP
        End If
    
        With tmprc
            .Open "SELECT CODIGO, DESCRIPCION FROM SECCIONES WHERE MBAJA = 0 ORDER BY CODIGO", miConn, adOpenDynamic, adLockReadOnly
            tmpcodsec = fg.BuildComboList(tmprc, "DESCRIPCION", "CODIGO", vbBlue)
            .Close
            .Open "SELECT CODIGO, DESCRIPCION FROM FAMILIAS WHERE MBAJA = 0 ORDER BY CODIGO", miConn, adOpenDynamic, adLockReadOnly
            tmpcodfam = fg.BuildComboList(tmprc, "DESCRIPCION", "CODIGO", vbBlue)
            .Close
            .Open "SELECT CODIGO, DESCRIPCION FROM SUBFAM WHERE MBAJA = 0 ORDER BY CODIGO", miConn, adOpenDynamic, adLockReadOnly
            tmpcodsubfam = fg.BuildComboList(tmprc, "DESCRIPCION", "CODIGO", vbBlue)
            .Close
            .Open "SELECT CODIGO, NOMBRE FROM MAPROV WHERE MBAJA = 0 ORDER BY CODIGO", miConn, adOpenDynamic, adLockReadOnly
            tmpcodprov = fg.BuildComboList(tmprc, "NOMBRE", "CODIGO", vbBlue)
            .Close
            .Open "SELECT IDTEM, TEMPORADA FROM TEMPOR WHERE MBAJA = 0 ORDER BY IDTEM", miConn, adOpenDynamic, adLockReadOnly
            tmpcodtempor = fg.BuildComboList(tmprc, "TEMPORADA", "IDTEM", vbBlue)
            .Close
            .Open "SELECT CODIGO, CAST(IVA as CHAR(3)) + ' %' AS DESCRIP FROM IVA WHERE MBAJA = 0 ORDER BY CODIGO", miConn, adOpenDynamic, adLockReadOnly
            tmpcodiva = fg.BuildComboList(tmprc, "DESCRIP", "CODIGO", vbBlue)
        End With
        
   ' With fg
    '    .ColComboList(2) = tmpcodsec
     '   .ColComboList(3) = tmpcodfam
      '  .ColComboList(4) = tmpcodsubfam
       ' .ColComboList(18) = tmpcodiva
   '     .ColComboList(19) = tmpcodprov
    '    .ColComboList(20) = tmpcodtempor
     '   .ColFormat(1) = "00000"
      '  .AutoSize 1, fg.Cols - 1
    'End With
    
    tmprc.Close
    Set tmprc = Nothing
        

    first = True

    End If
  
    Me.Caption = Caption_Form & " (total: " & (miRc.RecordCount) & ")"
    
  
'Set miConn = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Procedure : carga_combos_inicio
' DateTime  : 02/11/2003 13:56
' Author    : Administrador
' Purpose   : Cargar los combos con los datos iniciales
'---------------------------------------------------------------------------------------
Private Sub carga_combos_inicio()

 
  'cargar el sql para codigo

  

  'cargar codigo
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
With cbCODIGO
    .ConexionString = locCnn
    
    If Temporada_Defecto = 0 Then
        .SQLString = "SELECT CODIGO, MODELO FROM MAARTIC WHERE MBAJA = 0 ORDER BY CODIGO"
    Else
        .SQLString = "SELECT CODIGO, MODELO FROM MAARTIC WHERE MBAJA = 0 AND TEMPOR = " & Temporada_Defecto & " ORDER BY CODIGO"
    End If
    
    .LenCodigo = 5
    .CodigoWidth = 800
    .carga
    .Refresh
        
End With
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
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

'no hacer nada
If add_en_detalle And cbCODPROV.Text <> "" Then

Else

With cbCODPROV
    .ConexionString = locCnn
    .LenCodigo = 5
    .SQLString = "SELECT CODIGO, NOMBRE FROM MAPROV WHERE MBAJA = 0 ORDER BY CODIGO"
    .CodigoWidth = 800
    .carga
    DoEvents
    
    'cargar valor por defecto recibido del formulario frmPedProv
    If add_en_detalle Then
        .Text = Format(tmpcodprov, "00000")
        .Enabled = False
    End If
    
End With

End If

'no enlazar a datos
With cbTemporada
    .ConexionString = locCnn
    .SQLString = "SELECT IDTEM, AÑO + '  ' + TEMPORADA FROM TEMPOR WHERE MBAJA = 0 ORDER BY ACTUAL DESC, IDTEM"
    .LenCodigo = 3
    .CodigoWidth = 500
    .carga
    .TabStop = False
    
    If add_en_detalle Then
        .Text = Temporada_Defecto
        .Enabled = False
    Else
    'establecer la temprada actual como temporada de trabajo por defecto
        .Text = TemporadaActual
    End If
    
End With

'cbTemporada.Text = TemporadaActual

End Sub

Private Sub Form_Load()

  Move (Screen.Width - Width) \ 2, Separacion_MDIForm

  fg.Visible = False
  
    
  Call carga_combos_inicio
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    tmpcodsec = ""
    tmpcodfam = ""
    tmpcodsubfam = ""
    tmpcodprov = ""
    tmpcodtempor = ""
    tmpcodiva = ""
           
    'si hemos establecido un filtro que no devuelve ningun registro,
    'borrar filtro para que no de error al volver al formulario
    If miRc.EOF Then
        Call cbBorrar_click
    Else
        
        If add_en_detalle And consultado Then
            Call añade_en_detalle
        End If
        
    End If
    
    consultado = False
    Temporada_Defecto = 0
    
    Set miConn = Nothing
    Set frmFlexArt = Nothing
End Sub
    
Private Sub ioMODELO_Validate(Cancel As Boolean)

If ioMODELO.Text <> "" Then Call cbLista_click

End Sub



Private Sub ioREF_Validate(Cancel As Boolean)

If ioREF.Text <> "" Then Call cbLista_click

End Sub
