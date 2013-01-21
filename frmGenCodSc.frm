VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmGenCodSc 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Generar códigos de Seguridad ..."
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10410
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
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   Begin PCGestion.miText ioCODIGO 
      Height          =   525
      Left            =   945
      TabIndex        =   2
      Top             =   645
      Width           =   1665
      _ExtentX        =   3757
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
      Left            =   3855
      TabIndex        =   4
      Top             =   6585
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
      MICON           =   "frmGenCodSc.frx":0000
      PICN            =   "frmGenCodSc.frx":001C
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
      Left            =   5640
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6585
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
      MICON           =   "frmGenCodSc.frx":0CF6
      PICN            =   "frmGenCodSc.frx":0D12
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.miCombo cbCODALMORIG 
      Height          =   510
      Left            =   930
      TabIndex        =   0
      Top             =   45
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   900
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
   Begin PCGestion.bsGradientLabel lblCodigo 
      Height          =   465
      Left            =   2700
      Top             =   660
      Width           =   3555
      _ExtentX        =   4657
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
      CaptionAlignment=   1
   End
   Begin PCGestion.chameleonButton cbImprimir 
      Height          =   795
      Left            =   4815
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6585
      Width           =   780
      _ExtentX        =   1376
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
      MICON           =   "frmGenCodSc.frx":15EC
      PICN            =   "frmGenCodSc.frx":1608
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   4755
      Left            =   1703
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1275
      Width           =   7005
      _cx             =   12356
      _cy             =   8387
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
      FormatString    =   $"frmGenCodSc.frx":22E2
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
   Begin PCGestion.chameleonButton cbIntroducirDesde 
      Height          =   420
      Left            =   6345
      TabIndex        =   3
      Top             =   690
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   741
      BTYPE           =   9
      TX              =   "Introducir Todas desde ..."
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
      MICON           =   "frmGenCodSc.frx":23C0
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
      Left            =   0
      Top             =   6135
      Width           =   10395
      _ExtentX        =   11165
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
   Begin PCGestion.miCombo cbCODALMDEST 
      Height          =   510
      Left            =   6135
      TabIndex        =   1
      Top             =   60
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   900
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DESTINO"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5160
      TabIndex        =   10
      Top             =   105
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ORIGEN "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   75
      TabIndex        =   7
      Top             =   90
      Width           =   795
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO TRANS."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   -90
      TabIndex        =   5
      Top             =   570
      Width           =   1065
   End
End
Attribute VB_Name = "frmGenCodSc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Módulo       : frmGenCodSc
' Fecha/Hora : 18/06/2004 10:27
' Autor         : JCastillo
' Propósito    :Genera códigos de seguridad para una transferencia, o serie de transferencias
'                   dadas. Carga un grid y luego permite imprimirlo
'---------------------------------------------------------------------------------------


Option Explicit

Public CodigoTrn As Double
Public AlmacenTrn As Byte

Private Sub cbAceptar_Click()

    If MsgBox("¿Desea cerrar la ventana actual?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Unload Me
    
End Sub

Private Sub cbCancelar_Click()
    Unload Me
End Sub


Private Sub cbCODALMORIG_Validate(Cancel As Boolean)

If cbCODALMORIG.Text = "" Then
    lblstatus.Caption = "ORIGEN no puede estar en blanco"
    Cancel = True
End If

End Sub

Private Sub cbcodalmdest_Validate(Cancel As Boolean)

If cbCODALMORIG.Text = cbCODALMDEST.Text Then
    lblstatus.Caption = "No se permite ORIGEN y DESTINO iguales"
    Cancel = True
End If

End Sub

Private Sub cbImprimir_Click()
Dim linea1 As String
Dim linea2 As String

    On Error GoTo cbImprimir_Click_Error

    If fg.Rows <= 1 Then Exit Sub
    
    linea1 = "Códigos de Seguridad. Origen: " & Trim(devuelve_campo("select descripcion from almacenes where codigo = " & cbCODALMORIG.Text, locCnn))
    
    If cbCODALMDEST.Text <> "" Then linea1 = linea1 & ". Destino: " & Trim(devuelve_campo("select descripcion from almacenes where codigo = " & cbCODALMDEST.Text, locCnn))
    
    linea2 = "Usuario: " & Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & UsuarioActual, locCnn)) & ". Fecha: " & Now
    Call PrintFlexGrid(fg, 1, 1, 1, linea1, linea2, 12, 2)

   On Error GoTo 0
   Exit Sub

cbImprimir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbImprimir_Click de Formulario frmGenCodSc"

End Sub

Private Sub cbIntroducirDesde_Click()
Dim rc As New ADODB.Recordset

'pasarle el validate
On Error GoTo cbIntroducirDesde_Click_Error

Call ioCODIGO_Validate(False)


   With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
   End With
   


If cbCODALMDEST.Text = "" Then
   rc.Open "SELECT CODIGO, CODALMDEST FROM PTRANS WHERE ESTADO = 1 AND CODALMORIG = " & cbCODALMORIG.Text & " AND CODIGO >= " & ioCODIGO.Text, locCnn, adOpenStatic, adLockReadOnly
'si filtra por almacén de destino
Else
   rc.Open "SELECT CODIGO, CODALMDEST FROM PTRANS WHERE ESTADO = 1 AND CODALMORIG = " & cbCODALMORIG.Text & " AND CODALMDEST = " & cbCODALMDEST.Text & " AND CODIGO >= " & ioCODIGO.Text, locCnn, adOpenStatic, adLockReadOnly
End If

With fg
    .Clear
    .Cols = 4
    .Rows = 1
    .TextMatrix(0, 1) = "Transf."
    .TextMatrix(0, 2) = "Cod. Seguridad"
    .TextMatrix(0, 3) = "Destino"
    
    Do Until rc.EOF
    
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 1) = rc.fields("CODIGO")
        .TextMatrix(.Rows - 1, 2) = CodigoSeguridad_TRN(Format(rc.fields("CODIGO"), "000000000") & Format(cbCODALMORIG.Text, "000"))
        .TextMatrix(.Rows - 1, 3) = Trim(devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & rc.fields("CODALMDEST"), locCnn))
        
        rc.MoveNext
        
    Loop

    .SubtotalPosition = flexSTAbove
    .subtotal flexSTCount, , 3, , vbBlue, vbWhite, True
    
    If .Rows > 1 Then
    
    .TextMatrix(1, 1) = ""
    .TextMatrix(1, 3) = "Total : " & .TextMatrix(1, 3)
    
    End If
    
    .AutoSize 1, .Cols - 1

End With


rc.Close
Set rc = Nothing

   On Error GoTo 0
   Exit Sub

cbIntroducirDesde_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbIntroducirDesde_Click de Formulario frmGenCodSc"

End Sub

Private Sub Form_Load()

   On Error GoTo Form_Load_Error

Move (Screen.Width - Width) \ 2, Separacion_MDIForm

   With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
   End With
   
With cbCODALMORIG
    .ConexionString = locCnn
    .LenCodigo = 3
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM ALMACENES WHERE MBAJA = 0 ORDER BY CODIGO"
    .CodigoWidth = 500
     .carga
    .Text = AlmacenActual
    If TipoPermiso = 0 Then .Locked = True
End With

With cbCODALMDEST
    .ConexionString = locCnn
    .LenCodigo = 3
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM ALMACENES WHERE MBAJA = 0 ORDER BY CODIGO"
    .CodigoWidth = 500
     .carga
End With

With ioCODIGO
    .Alineacion = 1
    .LongMaxima = 10
    .SoloNumeros = True
End With

If CodigoTrn > 0 Then
    ioCODIGO.Text = CodigoTrn
    cbCODALMORIG.Text = AlmacenTrn
    Call ioCODIGO_Validate(False)
End If

With fg
    .Clear
    .Rows = 1
    .Cols = 3
End With

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Form_Load de Formulario frmGenCodSc"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Set frmGenCodSc = Nothing

End Sub



Private Sub ioCODIGO_Validate(Cancel As Boolean)
Dim m As Double

'comprobar almacen de origen
   On Error GoTo ioCODIGO_Validate_Error

lblCodigo.Caption = ""

If cbCODALMORIG.Text = "" Then
    lblstatus.Caption = "ORIGEN no puede estar en blanco"
    Cancel = False
    cbCODALMORIG.SetFocus
    Exit Sub
End If

If cbCODALMORIG.Text = cbCODALMDEST.Text Then
    lblstatus.Caption = "No se permite ORIGEN y DESTINO iguales"
    Cancel = False
    cbCODALMDEST.SetFocus
    Exit Sub
End If

'comprobar código de transferencia
If Trim(ioCODIGO.Text) <> "" Then
    If Not IsNumeric(ioCODIGO.Text) Then
    lblstatus.Caption = "CODIGO incorrecto"
    Cancel = True
    End If
Else
    lblstatus.Caption = "No se permite CODIGO en blanco"
    Cancel = True
End If

'---------------------------------------------------------------------------------------
'                     9 digitos codigo de transferencia
'                     3 digitos codigo de almacen
'                     Es un codigo de seguridad para poder aceptar la transferencia
'                     incluso si no coinciden las prendas en la comprobación.
'---------------------------------------------------------------------------------------
 m = CodigoSeguridad_TRN(Format(ioCODIGO.Text, "000000000") & Format(cbCODALMORIG.Text, "000"))
 
 lblCodigo.Caption = "Código: " & CStr(m)

   On Error GoTo 0
   Exit Sub

ioCODIGO_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioCODIGO_Validate de Formulario frmGenCodSc"

End Sub

