VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMsgPtrans 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mensajes de Transferencias"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11400
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox TextoRTF 
      Height          =   135
      Left            =   3975
      TabIndex        =   9
      Top             =   180
      Visible         =   0   'False
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   238
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMsgPtrans.frx":0000
   End
   Begin PCGestion.bsGradientLabel lblstatus 
      Height          =   375
      Left            =   0
      Top             =   5460
      Width           =   11370
      _ExtentX        =   16854
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
   Begin PCGestion.chameleonButton cbAgregar 
      Height          =   795
      Left            =   30
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "&Agregar F1"
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
      MICON           =   "frmMsgPtrans.frx":008A
      PICN            =   "frmMsgPtrans.frx":00A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbActualizar 
      Height          =   795
      Left            =   1125
      TabIndex        =   3
      Top             =   5880
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "&Actualizar F2"
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
      MICON           =   "frmMsgPtrans.frx":0D80
      PICN            =   "frmMsgPtrans.frx":0D9C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbEdicion 
      Height          =   795
      Left            =   2355
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5880
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "&Edicion F3"
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
      MICON           =   "frmMsgPtrans.frx":1676
      PICN            =   "frmMsgPtrans.frx":1692
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
      Left            =   8265
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5880
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
      MICON           =   "frmMsgPtrans.frx":1EF0
      PICN            =   "frmMsgPtrans.frx":1F0C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbEliminar 
      Height          =   795
      Left            =   9225
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "E&liminar F9"
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
      MICON           =   "frmMsgPtrans.frx":27E6
      PICN            =   "frmMsgPtrans.frx":2802
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbCerrar 
      Height          =   795
      Left            =   10320
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "Cerrar ESC"
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
      MICON           =   "frmMsgPtrans.frx":33D4
      PICN            =   "frmMsgPtrans.frx":33F0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   4980
      Left            =   30
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   465
      Width           =   11370
      _cx             =   20055
      _cy             =   8784
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
      FormatString    =   $"frmMsgPtrans.frx":40CA
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   4
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
   Begin PCGestion.bsGradientLabel lblAlmacen_Actual 
      Height          =   375
      Left            =   6315
      Top             =   30
      Width           =   5055
      _ExtentX        =   8916
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
      Colour1         =   11513775
      Colour2         =   16558731
      CaptionAlignment=   1
   End
   Begin VB.Label ioCODIGO 
      Alignment       =   2  'Center
      BackColor       =   &H00AC998C&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1875
      TabIndex        =   1
      Top             =   60
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSFERENCIA"
      Height          =   330
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   1620
   End
End
Attribute VB_Name = "frmMsgPtrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Plantilla
' DateTime  : 31/10/2003 10:08
' Author    : Administrador
' Purpose   : Plantilla de código para los formularios de maestros.
'---------------------------------------------------------------------------------------
'·································································································································
' Convenio:
'·································································································································
'
' Para los campos de texto:  usar miText
' Para los combos:              usar miCombo.
'
'
' - Instrucciones:
'
' Enlazar los controles a los campos en Form_Load()   (ver ). Y especificar la tabla
' y orden (y otras cosas que se pudieran necesitar) mediante los parametros del
' oSQL (objecto SmartSQL):
'
'  oSQL.AddTable "SECCIONES"
'  oSQL.AddOrderClause "CODIGO"

'---------------------------------------------------------------------------------------
' - Cambiar en:
'
' Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
' set plantilla = nothing
' por el nombre del formulario. Ej:   set frmMntPrue = nothing

'---------------------------------------------------------------------------------------
' - Si se utiliza algun campo simulando que sea incremental (que se incremente en cada
' registro) cambia en Private Sub cbAgregar_Click()
'
' tmpcodigo = devuelve_campo("select max(codigo) + 1 from secciones")
'
' y poner el SQL correcto para que nos devuelva el proximo codigo para nuestro campo
'
'
'
'---------------------------------------------------------------------------------------
' - Colocar 2 tipos de validaciones para los datos.
'
'  Una validación a nivel de campo. Por ejemplo, comprobar al salir del campo
'  que la información es correcta, usando el evento validate. (si es > X, <> "", etc)
'
'- Otra validación es en:
'
'Private Sub rc_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'
' El evento cuando cambia un registro del recordset, pones el codigo por ejemplo en

'        Case adRsnUpdate
'
' Mira el formulario frmMntCli para ver un ejemplo de esto.
'
'---------------------------------------------------------------------------------------
' - Formularios de Lista. Para llamar al formulario de lista estandar FrmFlexSimple, ver el
' codigo de cbLista_Click. Cambiar los colformats y otras cosas que puedan ser
' necesarias, para adecuar a cada formulario
'
'
'---------------------------------------------------------------------------------------
'Otras notas:
'
' - comprobar el orden correcto de los tabindex para permitir recorrer miText y miCombo
' del formulario con el teclado (soportan el avance con ENTER). Desde el primero hasta
' el ultimo.
'
'- cambiar en cbAgregar_clik y cbEditar_click
'
' ioDescripcion.setfocus
'
' por el nombre del control que tengamos que activar en primer lugar.
'
' cambiar en Private Sub ioCODIGO_Change(), y poner el numero de 000
' correcto en cada caso




Option Explicit
Dim WithEvents rc As ADODB.Recordset
Attribute rc.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim prime As Boolean

'codigo de la transferencia actual, para cargar los registros
Public TRNSF_ACTUAL As Long


Public miCODALMORIG As Byte

Dim oSQL As New clsSmartSQL

'---------------------------------------------------------------------------------------
' Subrutina   : fg_dblClick
' Fecha/Hora  : 08/12/2003 17:11
' Autor       : JCASTILLO
' Propósito   : Editar el mensaje seleccionado
'---------------------------------------------------------------------------------------
Private Sub fg_DblClick()

   On Error GoTo fg_dblClick_Error

With fg
    
    
    If rc.RecordCount = 0 Then Exit Sub
    
    'si esta editando o añadiendo entonces ...
    If mbAddNewFlag Or mbEditFlag Then
        lblstatus.Caption = "Debe guardar o cancelar los cambios antes de seleccionar otro mensaje"
        Exit Sub
    End If
    
    'si no selecciona ninguna salir
    If .TextMatrix(.Row, .Cols - 1) = "" Or .TextMatrix(.Row, .Cols - 1) = "ID" Then Exit Sub
       

    If Not rc.BOF Then rc.MoveFirst
    rc.Find "ID =" & .TextMatrix(.Row, .Cols - 1)
        
    DoEvents
        
    Call cbedicion_Click

End With

   On Error GoTo 0
   Exit Sub

fg_dblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento fg_dblClick de Formulario frmMsgPtrans"

End Sub

Private Sub ioCODIGO_Change()
ioCODIGO.Caption = Format(ioCODIGO.Caption, "00000")
End Sub

Private Sub Form_Activate()

If Not prime Then

 ' If rc.RecordCount = 0 Then
        
      '  If MsgBox("No se encuentran Mensajes para la transferencia actual. ¿Crear?", vbYesNo + vbQuestion, "Mensajes") = vbNo Then
         '   Unload Me
      '  Else
            'Call carga_grid
       '     Call cbAgregar_Click
      '  End If
        
 ' Else
'       Call cmdFirst_Click
        Call carga_grid
        Call cbCancelar_Click
        
  'End If


prime = True
End If
    
End Sub

Private Sub Form_Load()
  
  Move (Screen.Width - Width) \ 2, Separacion_MDIForm
  
   With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
   End With
   
  Set rc = New Recordset
  oSQL.AddTable "PTRANSMSG"
  oSQL.AddOrderClause "FMODI", True
  oSQL.AddOrderClause "ID"
  oSQL.AddSimpleWhereClause "CODIGO", TRNSF_ACTUAL
  rc.Open oSQL.SQL, locCnn, adOpenStatic, adLockOptimistic
    

''''''''
'////////////////////////////////////////////////////////
' Enlazar controles:
'////////////////////////////////////////////////////////

  With ioCODIGO
  Set .DataSource = rc
        .DataField = "CODIGO"
  End With
  
  'cargar datos en el grid
  Call carga_grid

  lblAlmacen_Actual.Caption = devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & AlmacenActual, locCnn)
       
  mbDataChanged = False
  
  fg.Rows = 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      
      If MsgBox("¿Desea cerrar la ventana actual?", vbQuestion + vbYesNo, "Cerrar") = vbYes Then
        cbcerrar_Click
      End If
      
      
       Case vbKeyF1
            Call cbAgregar_Click
        
      Case vbKeyF2
            Call cbactualizar_Click
        
      Case vbKeyF3
            Call cbedicion_Click
        
 
           
      
  End Select
  KeyCode = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error Resume Next

If rc.EditMode <> adEditNone Then rc.CancelUpdate

rc.Close
Set rc = Nothing

Set oSQL = Nothing

With frmPtrans
    .lblMensajes.Caption = "Mensajes: " & devuelve_campo("SELECT COUNT(CODIGO) FROM PTRANSMSG WHERE CODIGO = " & rc.fields("CODIGO") & " AND CODALMORIG = " & rc.fields("CODALMORIG"), locCnn)
    .WindowState = vbNormal
End With
    

Set frmMsgPtrans = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub





Private Sub rc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  If rc.AbsolutePosition > 0 Then
    lblstatus.Caption = "Registro: " & CStr(rc.AbsolutePosition)
    
  End If
End Sub

Private Sub rc_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Aquí se coloca el código de validación
  'Se llama a este evento cuando ocurre la siguiente acción
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cbAgregar_Click()
  Dim tmpcodigo As Variant
  
  
   On Error GoTo cbAgregar_Click_Error

  With rc
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    
    tmpcodigo = devuelve_campo("select max(ID) + 1 from PTRANSMSG where CODALM = " & AlmacenActual)
    
    .fields("CODIGO") = TRNSF_ACTUAL
    
    'Si devuelve @ esque ha habido un error
    If tmpcodigo = "@" Then tmpcodigo = 1
    
    .fields("ID") = tmpcodigo
    
    'End If

    .fields("FMODI") = Now
    .fields("CODUSR") = UsuarioActual
    .fields("CODALM") = AlmacenActual
    .fields("CODALMORIG") = miCODALMORIG
    .fields("MSG") = ""

    Set tmpcodigo = Nothing
    
    lblstatus.Caption = "Agregar registro"
    mbAddNewFlag = True
    SetButtons False
    cbActualizar.Visible = True

  End With

    
 Call carga_grid
 
 Me.WindowState = vbMinimized
 FrmInicio.Editor.carga "Mensaje de Transferencia", rc.fields("MSG"), ""
 
 DoEvents
 Me.WindowState = vbNormal
 DoEvents
 Call cbactualizar_Click
 
 DoEvents
 
 fg.AutoSize 1, fg.Cols - 1


   On Error GoTo 0
   Exit Sub

cbAgregar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbAgregar_Click de Formulario frmMsgPtrans"

End Sub

Private Sub cbEliminar_Click()
    On Error GoTo DeleteErr
  With rc
    '.Delete
    '.MoveNext
    .fields("mbaja") = True
    .fields("FBAJA") = Date
    If .EOF Then .MoveLast
  End With
 
  Call cbactualizar_Click
     
Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : cbedicion_Click
' Fecha/Hora  : 08/12/2003 21:07
' Autor       : JCASTILLO
' Propósito   : Editar o solo visualizar el mensaje (depende de si el usuario actual fue el
'               creador del mensaje.
'---------------------------------------------------------------------------------------
Private Sub cbedicion_Click()
  
  On Error GoTo cbedicion_Click_Error

  If rc.RecordCount = 0 Then Exit Sub
  
 rc.Move 0
  If rc.fields("CODUSR") = UsuarioActual Then
     
     Me.WindowState = vbMinimized
     FrmInicio.Editor.carga "Mensaje de Transferencia", rc.fields("MSG")
     Me.WindowState = vbNormal
     
     lblstatus.Caption = "Modificar registro"
     mbEditFlag = True
     SetButtons False
     cbActualizar.Visible = True
     
  Else
  
     Me.WindowState = vbMinimized
     'pasarle el texto del mensaje en vez del objeto field
     FrmInicio.Editor.carga "Mensaje de Transferencia (Solo lectura)", , rc.fields("MSG").Value, True
     Me.WindowState = vbNormal
     
  End If
  
   On Error GoTo 0
   Exit Sub

cbedicion_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbedicion_Click de Formulario frmMsgPtrans"

End Sub
Private Sub cbCancelar_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  rc.CancelUpdate
  If mvBookMark > 0 Then
    rc.Bookmark = mvBookMark
  Else
    rc.MoveFirst
  End If
  mbDataChanged = False

End Sub

Private Sub cbactualizar_Click()
  On Error GoTo UpdateErr
 
  rc.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    rc.MoveLast              'va al nuevo registro
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  
  lblstatus.Caption = ""
  Call carga_grid
  
 
  Exit Sub
UpdateErr:
'  If Err.Number = -2147217887 Then Exit Sub
  MsgBox Err.Description, vbInformation, "Atención"
End Sub

Private Sub cbcerrar_Click()
  Unload Me
End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : carga_Grid
' Fecha/Hora  : 08/12/2003 16:17
' Autor       : JCASTILLO
' Propósito   : Carga el grid de mensajes con los mensajes para la transferencia actual
'---------------------------------------------------------------------------------------
Private Sub carga_grid()

   On Error GoTo carga_grid_Error
    
    If rc.RecordCount = 0 Then Exit Sub
    
    
    If mbAddNewFlag Or mbEditFlag Then Exit Sub
    
    With fg
        .Clear
        .Rows = 1
        .Cols = 6
        .TextMatrix(0, 1) = "FECHA"
        .TextMatrix(0, 2) = "USUARIO"
        .TextMatrix(0, 3) = "ALMACEN"
        .TextMatrix(0, 4) = "MENSAJE"
        .TextMatrix(0, 5) = "ID"
    End With
    
    rc.Requery
    
    If Not rc.BOF Then rc.MoveFirst
    'tmprc.Open "SELECT * FROM PTRANSMSG WHERE CODIGO = " & TRNSF_ACTUAL, locCnn, adOpenDynamic, adLockReadOnly
    With fg
    
    Do Until rc.EOF
    
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 1) = rc.fields("FMODI")
        .TextMatrix(.Rows - 1, 2) = Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & rc.fields("CODUSR"), locCnn))
        .TextMatrix(.Rows - 1, 3) = Trim(devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & rc.fields("CODALM"), locCnn))
    
         TextoRTF.TextRTF = rc.fields("MSG")
        .TextMatrix(.Rows - 1, 4) = TextoRTF.Text
        
        .TextMatrix(.Rows - 1, 5) = rc.fields("ID")
    
    
    rc.MoveNext
    
    Loop
    
        .AutoSize 1, .Cols - 1
    End With

    rc.MoveFirst
    
   On Error GoTo 0
   Exit Sub

carga_grid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_Grid de Formulario frmMsgPtrans"
End Sub



Private Sub SetButtons(bVal As Boolean)
  cbAgregar.Visible = bVal
  cbEdicion.Visible = bVal
  cbActualizar.Visible = Not bVal
  cbCancelar.Visible = Not bVal
  cbEliminar.Visible = bVal
  cbCerrar.Visible = bVal
 ' cbLista.Visible = bVal
   
  'cbActualizar.Visible = bVal
 
End Sub
