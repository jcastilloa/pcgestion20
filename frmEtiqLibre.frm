VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmEtiqLibre 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sacar etiquetas ..."
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10185
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
   ScaleHeight     =   7125
   ScaleWidth      =   10185
   Begin PCGestion.miText ioCODBAR 
      Height          =   525
      Left            =   1260
      TabIndex        =   0
      Top             =   45
      Width           =   2595
      _ExtentX        =   4577
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
   Begin PCGestion.chameleonButton cbAceptar 
      Height          =   795
      Left            =   8250
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6285
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
      MICON           =   "frmEtiqLibre.frx":0000
      PICN            =   "frmEtiqLibre.frx":001C
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
      Left            =   9225
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6285
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
      MICON           =   "frmEtiqLibre.frx":0CF6
      PICN            =   "frmEtiqLibre.frx":0D12
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.miCombo cbCODTALLA 
      Height          =   495
      Left            =   6135
      TabIndex        =   4
      Top             =   600
      Width           =   4005
      _ExtentX        =   7064
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
   Begin PCGestion.miCombo cbCODCOL 
      Height          =   465
      Left            =   765
      TabIndex        =   5
      Top             =   1095
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   820
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
   Begin PCGestion.miText ioCODART 
      Height          =   525
      Left            =   4605
      TabIndex        =   1
      Top             =   75
      Width           =   1110
      _ExtentX        =   4233
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
   Begin PCGestion.miCombo cbTEMPOR 
      Height          =   480
      Left            =   6945
      TabIndex        =   2
      Top             =   75
      Width           =   3195
      _ExtentX        =   7011
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
   Begin PCGestion.miText ioUNIDADES 
      Height          =   525
      Left            =   4785
      TabIndex        =   6
      Top             =   1095
      Width           =   750
      _ExtentX        =   1323
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
      Height          =   4230
      Left            =   30
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1605
      Width           =   10140
      _cx             =   17886
      _cy             =   7461
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
      FocusRect       =   2
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmEtiqLibre.frx":15EC
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
      DataMode        =   0
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
   Begin PCGestion.chameleonButton cbGenerar 
      Height          =   825
      Left            =   30
      TabIndex        =   18
      Top             =   6270
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1455
      BTYPE           =   9
      TX              =   "&Generar"
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
      MICON           =   "frmEtiqLibre.frx":1691
      PICN            =   "frmEtiqLibre.frx":16AD
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
      Height          =   375
      Left            =   30
      Top             =   5865
      Width           =   10140
      _ExtentX        =   17886
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
   Begin PCGestion.chameleonButton cbBorrarUltima 
      Height          =   825
      Left            =   990
      TabIndex        =   19
      Top             =   6270
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   1455
      BTYPE           =   9
      TX              =   "&Borrar Ultima"
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
      MICON           =   "frmEtiqLibre.frx":2387
      PICN            =   "frmEtiqLibre.frx":23A3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.miCombo cbCATTALL 
      Height          =   495
      Left            =   1230
      TabIndex        =   3
      Top             =   585
      Width           =   4185
      _ExtentX        =   7382
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
   Begin PCGestion.chameleonButton cbInsertarPedido 
      Height          =   405
      Left            =   8460
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1125
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "&Insertar Pedido"
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
      MICON           =   "frmEtiqLibre.frx":307D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.miText ioDIGITOS 
      Height          =   525
      Left            =   7695
      TabIndex        =   8
      Top             =   1095
      Width           =   750
      _ExtentX        =   1323
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
   Begin PCGestion.miText ioSALTAR 
      Height          =   525
      Left            =   6255
      TabIndex        =   7
      Top             =   1095
      Width           =   750
      _ExtentX        =   1323
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Saltar"
      Height          =   360
      Left            =   5430
      TabIndex        =   23
      Top             =   1185
      Width           =   795
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Digitos"
      Height          =   360
      Left            =   6870
      TabIndex        =   22
      Top             =   1185
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORIA TALLA"
      Height          =   540
      Left            =   -195
      TabIndex        =   20
      Top             =   555
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TEMPORADA"
      Height          =   285
      Left            =   5730
      TabIndex        =   17
      Top             =   150
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      Height          =   360
      Left            =   3825
      TabIndex        =   16
      Top             =   150
      Width           =   795
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Uds."
      Height          =   360
      Left            =   4290
      TabIndex        =   15
      Top             =   1185
      Width           =   510
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TALLA"
      Height          =   300
      Left            =   5415
      TabIndex        =   14
      Top             =   690
      Width           =   690
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "COLOR"
      Height          =   285
      Left            =   0
      TabIndex        =   13
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO DE BARRAS"
      Height          =   555
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   1230
   End
End
Attribute VB_Name = "frmEtiqLibre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const fichero = "c:\TempEtiquetasDB.mdb"
Const sCnn = strCnnMdb & fichero
Dim cn As New ADODB.Connection

Dim etiqrc As New ADODB.Recordset

Private Sub cbAceptar_Click()
    Unload Me
End Sub

Private Sub cbCancelar_Click()
    Unload Me
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

Private Sub cbGenerar_Click()

 If ioSALTAR.Text = "" Then ioSALTAR.Text = "0"
  
 DoEvents
 Call Añade_Etiquetas_En_Blanco(ioSALTAR.Text)
 DoEvents
 Call procesa_informes(1, False)

End Sub


Private Sub cbTEMPOR_Validate(Cancel As Boolean)

   On Error GoTo cbTEMPOR_Validate_Error

If ioCODART.Text <> "" And cbTEMPOR.Text <> "" Then
 
    If devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & CLng(ioCODART.Text) & " AND TEMPOR = " & CLng(cbTEMPOR.Text)) = "@" Then
        
                lblstatus.Caption = "No existe el artículo para esa temporada!"
                cbTEMPOR.SetFocus
                Cancel = True
                Exit Sub
                
    Else
    
        lblstatus.Caption = ""
        'Call carga_almacenes_origen(cbCODALMORIG)
                
    End If
 
 End If

   On Error GoTo 0
   Exit Sub

cbTEMPOR_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbTEMPOR_Validate de Formulario frmEtiqLibre"
 
End Sub

Private Sub cbBorrarUltima_Click()

   On Error GoTo cbBorrarUltima_Click_Error


If etiqrc.RecordCount <= 0 Then Exit Sub

If MsgBox("¿Desea borrar la última etiqueta?", vbQuestion + vbYesNo, titulo) = vbNo Then Exit Sub

etiqrc.MoveLast
etiqrc.Delete

Call cargar_grid

   On Error GoTo 0
   Exit Sub

cbBorrarUltima_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbBorrarUltima_Click de Formulario frmEtiqLibre"

End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : cbInsertarPedido_Click
' Fecha/Hora    : 29/01/2004 17:24
' Autor         : JCastillo
' Propósito     :   Insertar las etiquetas del pedido
'---------------------------------------------------------------------------------------
Private Sub cbInsertarPedido_Click()
Dim rcP As New ADODB.Recordset
Dim tmpp As Long

   'On Error GoTo cbInsertarPedido_Click_Error
   
   tmpp = InputBox("Introduzca Pedido", "Sacar etiquetas")
   
   If ioDIGITOS.Text = "" Then ioDIGITOS.Text = "11"
   
   If Trim(tmpp) = "" Or Not IsNumeric(tmpp) Then Exit Sub
   
   lblstatus.Caption = "Espere por favor ..."
   
   rcP.Open "SELECT * FROM DETPEDPRO WHERE NUMERO = " & tmpp & " AND ALMORIG = " & AlmacenActual, locCnn, adOpenDynamic, adLockReadOnly

   Do Until rcP.EOF
   
      'insertar una etiqueta
      Call inserta_etiqueta(Format(rcP.fields("CODART"), "00000") & Format(rcP.fields("TEMPOR"), "000") & Format(rcP.fields("CODTALLA"), "00") & Format(rcP.fields("CODCOL"), "000"), rcP.fields("UNIDADES"), ioDIGITOS.Text)
      DoEvents
   
    If Not rcP.EOF Then rcP.MoveNext
   Loop
   
   etiqrc.Requery
   Call cargar_grid
   
   lblstatus.Caption = ""

rcP.Close
Set rcP = Nothing

   On Error GoTo 0
   Exit Sub

cbInsertarPedido_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbInsertarPedido_Click de Formulario frmEtiqLibre"

End Sub

Private Sub Form_Load()

Move (Screen.Width - Width) \ 2, Separacion_MDIForm
   
   With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
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
    .SQLString = "SELECT IDTEM, AÑO + '  ' + TEMPORADA FROM TEMPOR WHERE MBAJA = 0 ORDER BY ACTUAL DESC, IDTEM"
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

With ioUNIDADES
    .LongMaxima = 5
    .SoloNumeros = True
    .PermitirBlanco = False
    .Alineacion = 1
End With

With ioCODBAR
    .LongMaxima = LenCodBar
    .SoloNumeros = True
End With

With ioSALTAR
    .LongMaxima = 3
    .SoloNumeros = True
    .Alineacion = 1
End With

With ioDIGITOS
    .LongMaxima = 2
    .SoloNumeros = True
    .Alineacion = 1
End With

Call CreateDBEtiquetas
etiqrc.Open "SELECT * FROM ETIQUETAS", cn, adOpenStatic, adLockOptimistic



End Sub




'---------------------------------------------------------------------------------------
' Procedure : CreateDBEtiquetas
' DateTime  : 10/11/2003 20:36
' Author    : Administrador
' Purpose   : Rutina que crea la base de datos temporal donde se alma
'             cenaran los registros que se van al imprimir como etiquetas
'             (un registro por cada unidad articulo/talla/color.
'---------------------------------------------------------------------------------------
'
Private Sub CreateDBEtiquetas()
On Error GoTo ErrorCreateDB

Dim Cat     As New ADOX.Catalog
Dim Tbl(6) As ADOX.Table
Dim Idx()   As ADOX.Index
Dim msgErrR As Integer
'Dim sCnn    As String

Dim tmpean As String * LenCodBar


ChDir ("c:\")

'si existe uno previo, borrar
If Dir(fichero) <> "" Then Kill fichero


Cat.Create sCnn

  '----------* Table Definition of ETIQUETAS *----------
  Set Tbl(0) = New ADOX.Table
  Tbl(0).ParentCatalog = Cat
  With Tbl(0)
    .Name = "ETIQUETAS"
    .Columns.Append "ABREVIA", adVarWChar, 20
    .Columns.Append "CODCOLOR", adVarWChar, 3
      .Columns("CODCOLOR").Properties("Default").Value = "0"
    .Columns.Append "CODIGO", adVarWChar, 5
      .Columns("CODIGO").Properties("Default").Value = "0"
    .Columns.Append "CODTALLA", adVarWChar, 2
      .Columns("CODTALLA").Properties("Default").Value = "0"
    .Columns.Append "DESCOLOR", adVarWChar, 15
    .Columns.Append "DESCTALLA", adVarWChar, 15
    .Columns.Append "MODELO", adVarWChar, 30
    .Columns.Append "REFERENCIA", adVarWChar, 15
    .Columns.Append "Id", adInteger
      .Columns("Id").Properties("AutoIncrement").Value = True
      .Columns("Id").Properties("Nullable").Value = False
    
    .Columns.Append "PRECOM", adVarWChar, 6
   
    
    .Columns.Append "PROVEEDOR", adVarWChar, 3
    '.Columns("PROVEEDOR").Properties("Default").Value = ""
      
    .Columns.Append "PVP", adCurrency
      .Columns("PVP").Properties("Default").Value = "0"
    .Columns.Append "TEMPOR", adVarWChar, 3
      .Columns("TEMPOR").Properties("Default").Value = "0"
      
      .Columns.Append "IMAGEN", adLongVarBinary
      .Columns("IMAGEN").Properties("Description").Value = "IMAGEN"

  End With
  '----------* Index Defini4tions of ETIQUETAS *----------
  ReDim Idx(0)
  Set Idx(0) = New ADOX.Index
    Idx(0).Name = "PrimaryKey"
    Idx(0).PrimaryKey = True
    Idx(0).Unique = True
      Idx(0).Columns.Append "Id"
  Tbl(0).Indexes.Append Idx(0)

  Cat.tables.Append Tbl(0)

  Set Cat = Nothing
  
'tempor
'codart
'codtalla
'codcol
'preven

cn.Open sCnn
  
  
  
  'BarCodefrm.BarCode1.DataField = "IMAGEN"
  'Set BarCodefrm.BarCode1.DataSource = etiqrc
  
  'meterle etiquetas en blanco si hay para saltar
   

 ' frmBarcode.Show
  DoEvents
  

  tmpean = ""
  
  Exit Sub

ErrorCreateDB:
    msgErrR = MsgBox("    Error No. " & Err & " " & vbCrLf & Error, vbCritical + vbAbortRetryIgnore, "Code Gen Error")
    Select Case msgErrR
      Case Is = vbAbort
      If Not (Cat Is Nothing) Then
        Set Cat = Nothing
      End If
      Exit Sub
     Case Is = vbRetry
       Resume Next
     Case Is = vbIgnore
       Resume
    End Select

End Sub


Private Sub Añade_Etiquetas_En_Blanco(saltar As Byte)
 
 Dim tmpvar As Byte
  
  'borras las q esten en blanco de antes
  cn.Execute "DELETE FROM ETIQUETAS WHERE PROVEEDOR = ' '"
  
  If saltar > 0 Then
  
    If Not etiqrc.EOF Then etiqrc.MoveFirst
    For tmpvar = 0 To saltar - 1
        
            etiqrc.AddNew
            etiqrc.fields("ABREVIA") = " "
            etiqrc.fields("TEMPOR") = " "
            etiqrc.fields("CODIGO") = " "
            etiqrc.fields("CODTALLA") = " "
            etiqrc.fields("CODCOLOR") = " "
            etiqrc.fields("PRECOM") = " "
            etiqrc.fields("PVP") = "0"
            etiqrc.fields("PROVEEDOR") = " "
            etiqrc.fields("MODELO") = " "
            etiqrc.fields("DESCOLOR") = " "
            etiqrc.fields("DESCTALLA") = " "
            etiqrc.fields("REFERENCIA") = " "
            GuardarArchivo etiqrc.fields("IMAGEN"), App.Path & "\Blanca.bmp"
           
        etiqrc.Update
        DoEvents
    
    Next tmpvar
  
  End If
  
End Sub


Private Sub inserta_etiqueta(Barcode As String, unidades As Long, digitos As String)
Dim codb As MiCodBar
Dim tmpprov As Long
Dim nveces As Long
Dim tmp_precio As Single

  'ahora meterle los datos ....
  'With rc_detped
    
        'si estan pendientes de meter en el almacén
        'If .Fields("METIDO") = 0 Then
            'un registro por cada unidad para el mismo articulo
            For nveces = 1 To unidades
    
            'CODIGO DE BARRAS 13 DIGITOS:
            
            'ARTICULO: 5 digitos
            'TEMPORADA:3 digitos
            'TALLA:    2 digitos
            'COLOR:    3 digitos
                        
            'tmpean = barcode
            codb = Descompone_CBAR(Barcode)
            'tmpean = Format(.Fields("CODART"), "00000") & Format(.Fields("TEMPOR"), "000") & Format(.Fields("CODTALLA"), "00") & Format(.Fields("CODCOL"), "000")
   
            'Debug.Print tmpean
            
            'PaintCode BarCodefrm, Mid$(tmpean, 1, 1), Mid$(tmpean, 2, 6), Mid$(tmpean, 8, 6)

           ' frmBarcode.txtData.Text = tmpean
           ' frmBarcode.DrawBarCode "128"
           ' frmBarcode.cmdBMP_Click
           ' DoEvents
            
            If Trim(codb.TALLA_ART) = "" Or Trim(codb.COLOR_ART) = "" Then
            lblstatus.Caption = "El pedido seleccionado tiene TALLAS/COLORES en blanco"
            Exit Sub
            End If
            
            
            
            'tmp_precio = Obtiene_Precom_Pedido(CLng(codb.CODIGO_ART), CLng(codb.TEMPORADA_ART), CLng(codb.TALLA_ART), CLng(codb.COLOR_ART), locCnn)
            
            'If tmp_precio = 0 Then
             '   MsgBox "No se encuentra el precio de compra en ningun pedido. Imposible hacer etiqueta", vbExclamation, titulo
            '    Exit Sub
            'End If
                        
            BarCodefrm.BarCode1.DataToEncode = Barcode
            DoEvents
            Set BarCodefrm.Picture1.Picture = BarCodefrm.BarCode1.Picture
            DoEvents
            'BarCodefrm.barcode1.Picture
            
            SavePicture BarCodefrm.Picture1.Image, App.Path & "\Barcode.bmp"
            'BarCodefrm.BarCode1.SaveBarCode "c:\pruebasav.wmf"
            DoEvents
            
            'Set BarCodefrm.Picture1.Picture = LoadPicture("c:\pruebasav.wmf")
            
            DoEvents

            'Set BarCodefrm.Picture1.Picture = BarCodefrm.BarCode1.Picture
            
            etiqrc.AddNew
            
            GuardarArchivo etiqrc.fields("IMAGEN"), App.Path & "\Barcode.bmp"
            'GuardarBinary etiqrc.Fields("IMAGEN"), BarCodefrm.Picture1
            'GuardarBinary etiqrc.Fields("IMAGEN"), BarCodefrm.BarCode1.Picture
            
            tmpprov = devuelve_campo("SELECT CODPROV FROM MAARTIC WHERE CODIGO =" & codb.CODIGO_ART & " and TEMPOR = " & codb.TEMPORADA_ART, locCnn)
            tmp_precio = devuelve_campo("SELECT PRECOM FROM MAARTIC WHERE CODIGO =" & codb.CODIGO_ART & " and TEMPOR = " & codb.TEMPORADA_ART, locCnn)
            
            etiqrc.fields("ABREVIA") = devuelve_campo("SELECT ABREVIA FROM TEMPOR WHERE IDTEM =" & codb.TEMPORADA_ART, locCnn)
            etiqrc.fields("TEMPOR") = Format(codb.TEMPORADA_ART, "000")
            etiqrc.fields("CODIGO") = Format(codb.CODIGO_ART, "00000")
            etiqrc.fields("CODTALLA") = Format(codb.TALLA_ART, "00")
            etiqrc.fields("CODCOLOR") = Format(codb.COLOR_ART, "000")
            etiqrc.fields("PRECOM") = Mid(digitos, 1, 2) & Format(tmp_precio, "00")
            etiqrc.fields("PVP") = devuelve_campo("SELECT PREVEN FROM MAARTIC WHERE CODIGO =" & codb.CODIGO_ART & " and TEMPOR = " & codb.TEMPORADA_ART, locCnn)
            etiqrc.fields("PROVEEDOR") = Mid(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & tmpprov, locCnn), 1, 3)
            etiqrc.fields("MODELO") = devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO =" & codb.CODIGO_ART & " and TEMPOR = " & codb.TEMPORADA_ART, locCnn)
            etiqrc.fields("DESCOLOR") = devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO =" & codb.COLOR_ART, locCnn)
            etiqrc.fields("DESCTALLA") = devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO =" & codb.TALLA_ART, locCnn)
            etiqrc.fields("REFERENCIA") = devuelve_campo("SELECT REF FROM MAARTIC WHERE CODIGO =" & codb.CODIGO_ART & " and TEMPOR = " & codb.TEMPORADA_ART, locCnn)
            etiqrc.Update
        
            DoEvents
            Next
       '  End If
         
       ' .MoveNext
    
    'Loop

 ' End With
  
'  etiqrc.Close
'  Set etiqrc = Nothing
  


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If etiqrc.State = 1 Then etiqrc.Close
  If cn.State = 1 Then cn.Close
  Set cn = Nothing
  Set etiqrc = Nothing
  Set frmEtiqLibre = Nothing
End Sub

Private Sub cargar_grid()
Dim tmpcodcolor As Long
'Dim rc As New ADODB.Recordset

On Error GoTo cargar_grid_Error

'rc.Open "SELECT * FROM ETIQUETAS", cn, adOpenDynamic, adLockOptimistic

fg.Clear
fg.Rows = 2
fg.Cols = 3
fg.Redraw = flexRDNone

  'poner títulos al grid
  With fg
        .Clear
        .Rows = 1
        .Cols = 7
       ' .ColHidden(1) = True
        .TextMatrix(0, 1) = "Codigo"
        .TextMatrix(0, 2) = "Modelo"
        .TextMatrix(0, 3) = "Talla"
        .TextMatrix(0, 4) = "Color"
        .TextMatrix(0, 5) = "Precio Compra"
        .TextMatrix(0, 6) = "Precio Venta"
  End With
  
  If etiqrc.RecordCount <= 0 Then Exit Sub
    
  etiqrc.MoveFirst
  
  DoEvents

Do Until etiqrc.EOF

    fg.Rows = fg.Rows + 1
    
    fg.TextMatrix(fg.Rows - 1, 1) = etiqrc.fields("CODIGO")
    fg.TextMatrix(fg.Rows - 1, 2) = etiqrc.fields("MODELO")
    fg.TextMatrix(fg.Rows - 1, 3) = etiqrc.fields("DESCTALLA")
    fg.TextMatrix(fg.Rows - 1, 4) = etiqrc.fields("DESCOLOR")
    
    fg.TextMatrix(fg.Rows - 1, 5) = etiqrc.fields("PRECOM")
    fg.TextMatrix(fg.Rows - 1, 6) = etiqrc.fields("PVP")
      
    tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & etiqrc.fields("CODCOLOR"))
    fg.Col = 4
    fg.Row = fg.Rows - 1
    fg.CellBackColor = tmpcodcolor
    fg.Col = 2
              

    DoEvents

    If Not etiqrc.EOF Then etiqrc.MoveNext

Loop

fg.SubtotalPosition = flexSTAbove
fg.subtotal flexSTCount, , 2, , vbBlue, vbWhite
fg.TextMatrix(1, 1) = "Total Etiquetas:"

fg.HighLight = flexHighlightWithFocus
fg.FocusRect = flexFocusHeavy
fg.AllowBigSelection = False
fg.AllowSelection = True

fg.AutoSize 1, fg.Cols - 1

fg.Redraw = True
'etiqrc.Close
'Set rc = Nothing

   On Error GoTo 0
   Exit Sub

cargar_grid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cargar_grid de Formulario frmEtiqLibre"

End Sub

Private Sub ioCODBAR_LostFocus()
Dim codb As MiCodBar
 
 If Len(ioCODBAR.Text) = LenCodBar Then
 
    codb = Descompone_CBAR(ioCODBAR.Text)
    
    ioCODART.Text = Format(codb.CODIGO_ART, "00000")
    cbTEMPOR.Text = Format(codb.TEMPORADA_ART, "000")
    cbCODTALLA.Text = Format(codb.TALLA_ART, "00")
    cbCODCOL.Text = Format(codb.COLOR_ART, "000")
    
    inserta_etiqueta ioCODART.Text & cbTEMPOR.Text & cbCODTALLA.Text & cbCODCOL.Text, 1, "11"
    
    DoEvents
    
    Call cargar_grid
    DoEvents
    ioCODBAR.Text = ""
    
    ioCODART.Text = ""
    cbTEMPOR.Text = ""
    cbCODTALLA.Text = ""
    cbCODCOL.Text = ""
    
    ioCODBAR.SetFocus
    
 End If

End Sub



'---------------------------------------------------------------------------------------
' Procedimiento : ioUNIDADES_Validate
' Fecha/Hora    :  21/01/2004 18:00
' Autor         : JCastillo
' Propósito     :
'---------------------------------------------------------------------------------------
'
Private Sub ioUNIDADES_Validate(Cancel As Boolean)
    
   On Error GoTo ioUNIDADES_Validate_Error

    If ioDIGITOS.Text = "" Then ioDIGITOS.Text = "11"
    
    inserta_etiqueta Format(ioCODART.Text, "00000") & Format(cbTEMPOR.Text, "000") & Format(cbCODTALLA.Text, "00") & Format(cbCODCOL.Text, "000"), ioUNIDADES.Text, ioDIGITOS.Text
    
    Call cargar_grid
    DoEvents
    ioCODBAR.Text = ""
    
    ioCODART.Text = ""
    cbTEMPOR.Text = ""
    cbCODTALLA.Text = ""
    cbCODCOL.Text = ""

   On Error GoTo 0
   Exit Sub

ioUNIDADES_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioUNIDADES_Validate de Formulario frmEtiqLibre"
    
End Sub

