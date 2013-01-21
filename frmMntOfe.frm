VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMntOfe 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ofertas ..."
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9765
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
   ScaleHeight     =   4665
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   Begin PCGestion.chameleonButton cmdPrevious 
      Height          =   630
      Left            =   1065
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3165
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1111
      btype           =   9
      tx              =   "F6"
      enab            =   -1  'True
      font            =   "frmMntOfe.frx":0000
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntOfe.frx":002C
      picn            =   "frmMntOfe.frx":004A
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblstatus 
      Height          =   375
      Left            =   30
      Top             =   2745
      Width           =   9690
      _extentx        =   17092
      _extenty        =   661
      caption         =   ""
      fount           =   "frmMntOfe.frx":0D1E
      captioncolour   =   0
      colour1         =   15640462
      colour2         =   7177785
      captionalignment=   1
   End
   Begin PCGestion.chameleonButton cmdFirst 
      Height          =   630
      Left            =   30
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3165
      Width           =   1005
      _extentx        =   1773
      _extenty        =   1111
      btype           =   9
      tx              =   "F5"
      enab            =   -1  'True
      font            =   "frmMntOfe.frx":0D4C
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntOfe.frx":0D78
      picn            =   "frmMntOfe.frx":0D96
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cbLista 
      Height          =   630
      Left            =   4485
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3165
      Width           =   1065
      _extentx        =   1879
      _extenty        =   1111
      btype           =   9
      tx              =   "Lista F4"
      enab            =   -1  'True
      font            =   "frmMntOfe.frx":1ACE
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntOfe.frx":1AFA
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cmdNext 
      Height          =   630
      Left            =   7605
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3165
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1111
      btype           =   9
      tx              =   "F7"
      enab            =   -1  'True
      font            =   "frmMntOfe.frx":1B18
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntOfe.frx":1B44
      picn            =   "frmMntOfe.frx":1B62
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   1
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cmdLast 
      Height          =   630
      Left            =   8670
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3165
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1111
      btype           =   9
      tx              =   "F8"
      enab            =   -1  'True
      font            =   "frmMntOfe.frx":2836
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntOfe.frx":2862
      picn            =   "frmMntOfe.frx":2880
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   1
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cbAgregar 
      Height          =   795
      Left            =   15
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3825
      Width           =   1065
      _extentx        =   1879
      _extenty        =   1402
      btype           =   9
      tx              =   "&Agregar F1"
      enab            =   -1  'True
      font            =   "frmMntOfe.frx":35B8
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntOfe.frx":35E4
      picn            =   "frmMntOfe.frx":3602
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cbActualizar 
      Height          =   795
      Left            =   1110
      TabIndex        =   7
      Top             =   3825
      Width           =   1200
      _extentx        =   2117
      _extenty        =   1402
      btype           =   9
      tx              =   "&Actualizar F2"
      enab            =   -1  'True
      font            =   "frmMntOfe.frx":42DE
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntOfe.frx":430A
      picn            =   "frmMntOfe.frx":4328
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cbEdicion 
      Height          =   795
      Left            =   2340
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3825
      Width           =   990
      _extentx        =   1746
      _extenty        =   1402
      btype           =   9
      tx              =   "&Edicion F3"
      enab            =   -1  'True
      font            =   "frmMntOfe.frx":4C04
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntOfe.frx":4C30
      picn            =   "frmMntOfe.frx":4C4E
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cbCancelar 
      Height          =   795
      Left            =   6555
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3825
      Width           =   930
      _extentx        =   1640
      _extenty        =   1402
      btype           =   9
      tx              =   "&Cancelar"
      enab            =   -1  'True
      font            =   "frmMntOfe.frx":54AE
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntOfe.frx":54DA
      picn            =   "frmMntOfe.frx":54F8
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cbEliminar 
      Height          =   795
      Left            =   7545
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3825
      Width           =   1065
      _extentx        =   1879
      _extenty        =   1402
      btype           =   9
      tx              =   "E&liminar F9"
      enab            =   -1  'True
      font            =   "frmMntOfe.frx":5DD4
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntOfe.frx":5E00
      picn            =   "frmMntOfe.frx":5E1E
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cbCerrar 
      Height          =   795
      Left            =   8670
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3825
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1402
      btype           =   9
      tx              =   "Cerrar ESC"
      enab            =   -1  'True
      font            =   "frmMntOfe.frx":69F2
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntOfe.frx":6A1E
      picn            =   "frmMntOfe.frx":6A3C
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.miText ioFECHAFIN 
      Height          =   525
      Left            =   4335
      TabIndex        =   3
      Top             =   1125
      Width           =   1365
      _extentx        =   2408
      _extenty        =   926
      font            =   "frmMntOfe.frx":7718
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioIMPORTE 
      Height          =   525
      Left            =   6660
      TabIndex        =   4
      Top             =   1140
      Width           =   1305
      _extentx        =   2302
      _extenty        =   926
      font            =   "frmMntOfe.frx":7744
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioDCTO 
      Height          =   540
      Left            =   8625
      TabIndex        =   5
      Top             =   1140
      Width           =   1005
      _extentx        =   1773
      _extenty        =   953
      font            =   "frmMntOfe.frx":7770
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioDescripcion 
      Height          =   525
      Left            =   1530
      TabIndex        =   6
      Top             =   1680
      Width           =   8100
      _extentx        =   14288
      _extenty        =   926
      font            =   "frmMntOfe.frx":779C
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioFECHAINI 
      Height          =   525
      Left            =   1530
      TabIndex        =   2
      Top             =   1110
      Width           =   1365
      _extentx        =   2408
      _extenty        =   926
      font            =   "frmMntOfe.frx":77C8
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miCombo cbCAJAS 
      Height          =   540
      Left            =   1530
      TabIndex        =   0
      Top             =   555
      Width           =   4140
      _extentx        =   7303
      _extenty        =   953
      font            =   "frmMntOfe.frx":77F4
   End
   Begin PCGestion.miCombo cbTIPO 
      Height          =   540
      Left            =   6255
      TabIndex        =   1
      Top             =   585
      Width           =   3345
      _extentx        =   7303
      _extenty        =   953
      font            =   "frmMntOfe.frx":7820
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO"
      Height          =   330
      Left            =   5685
      TabIndex        =   27
      Top             =   660
      Width           =   540
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
      Left            =   1575
      TabIndex        =   26
      Top             =   90
      Width           =   1200
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CAJA"
      Height          =   330
      Left            =   915
      TabIndex        =   25
      Top             =   630
      Width           =   540
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "IMPORTE"
      Height          =   330
      Left            =   5670
      TabIndex        =   24
      Top             =   1215
      Width           =   945
   End
   Begin VB.Label lblDcto 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DCTO"
      Height          =   315
      Left            =   7980
      TabIndex        =   23
      Top             =   1230
      Width           =   600
   End
   Begin MSForms.CheckBox ioMBAJA 
      Height          =   435
      Left            =   900
      TabIndex        =   22
      Top             =   2235
      Width           =   840
      VariousPropertyBits=   746596375
      BackColor       =   12632256
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "1482;767"
      Value           =   "0"
      Caption         =   "Baja"
      FontName        =   "Trebuchet MS"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      Height          =   330
      Left            =   645
      TabIndex        =   11
      Top             =   135
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA FIN"
      Height          =   315
      Left            =   3135
      TabIndex        =   10
      Top             =   1200
      Width           =   1110
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA INI"
      Height          =   330
      Left            =   285
      TabIndex        =   9
      Top             =   1200
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION"
      Height          =   360
      Left            =   30
      TabIndex        =   8
      Top             =   1770
      Width           =   1410
   End
End
Attribute VB_Name = "frmMntOfe"
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

Dim oSQL As New clsSmartSQL


Private Sub cbTIPO_Validate(Cancel As Boolean)

Select Case cbTIPO.Text
    
  Case 3 'importe
  
    ioDCTO.Visible = False
    ioIMPORTE.Visible = True
    
    lblImporte.Visible = True
    lblDcto.Visible = False
  
  Case 2 '%
  
    ioDCTO.Visible = True
    ioIMPORTE.Visible = False
    
    lblImporte.Visible = False
    lblDcto.Visible = True
    
  Case 1  '2x1
  
    ioIMPORTE.Visible = False
    ioDCTO.Visible = False
    
    lblImporte.Visible = False
    lblDcto.Visible = False
  
End Select

End Sub

Private Sub ioCODIGO_Change()
ioCODIGO.Caption = Format(ioCODIGO.Caption, "00000")
End Sub

Private Sub Form_Activate()

If Not prime Then

  If rc.RecordCount = 0 Then
        
        If MsgBox("No se encuentran Ofertas. ¿Crear?", vbYesNo + vbQuestion, "Ofertas") = vbNo Then
        Unload Me
        Else
        Call cbAgregar_Click
        End If
        
  Else
        Call cmdFirst_Click
        Call cbCancelar_Click
        
  End If

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
  oSQL.AddTable "OFERTAS"
  oSQL.AddOrderClause "CODIGO"
  
  If TipoPermiso = 0 Then
    oSQL.AddSimpleWhereClause "CODCAJA", CByte(CajaActual), , CLAUSE_EQUALS
  End If
  
  rc.Open oSQL.SQL, locCnn, adOpenStatic, adLockOptimistic
    
''''''''
'////////////////////////////////////////////////////////
' Enlazar controles:
'////////////////////////////////////////////////////////

  With ioCODIGO
  Set .DataSource = rc
        .DataField = "CODIGO"
  End With
  
  With ioDescripcion
  Set .DataSource = rc
        .PermitirBlanco = True
        .DataField = "DESCRIPCION"
        .LongMaxima = 15
  End With
  
  With ioDCTO
        .DataField = "DCTO"
        Set .DataSource = rc
        .LongMaxima = 3
        .Alineacion = 1
  End With
  
  With ioFECHAINI
  Set .DataSource = rc
        .DataField = "FINICIO"
        .dspFormat = "dd/mm/yyyy"
        .LongMaxima = 10
  End With
  
  With ioFECHAFIN
  Set .DataSource = rc
        .DataField = "FFIN"
        .dspFormat = "dd/mm/yyyy"
        .LongMaxima = 10
  End With
  
  With ioIMPORTE
        .dspFormat = "Currency"
        .LongMaxima = 10
        .Alineacion = 1
  End With
  
   'Cargar el micombo cajas
   With cbCAJAS
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM CAJAS WHERE MBAJA = 0 ORDER BY CODIGO"
    .LenCodigo = 1
    .CodigoWidth = 500
    .carga
    Set .DataSource = rc
    .DataField = "CODCAJA"
    .Refresh
    DoEvents
  End With
  
  With cbTIPO
    .LenCodigo = 1
    .CodigoWidth = 500
    .añade_item "1  Oferta 2x1", 1
    .añade_item "2  Oferta %", 2
    .añade_item "3  Oferta Importe", 3
    Set .DataSource = rc
    .DataField = "TIPO"
  End With
  
  With ioMBAJA
  Set .DataSource = rc
        .DataField = "MBAJA"
  End With
              
  mbDataChanged = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      
      If MsgBox("¿Desea cerrar la ventana actual?", vbQuestion + vbYesNo, "Cerrar") = vbYes Then
        cbcerrar_Click
      End If
      
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
      
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
      
      Case vbKeyF1
            Call cbAgregar_Click
        
      Case vbKeyF2
            Call cbactualizar_Click
        
      Case vbKeyF3
            Call cbedicion_Click
        
      Case vbKeyF4
            Call cbLista_click
      
      Case vbKeyF5
            Call cmdFirst_Click
    
       Case vbKeyF6
            Call cmdPrevious_Click
      
       Case vbKeyF7
            Call cmdNext_Click
    
       Case vbKeyF8
        Call cmdLast_Click
      
      
  End Select
  KeyCode = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error Resume Next

If rc.EditMode <> adEditNone Then rc.CancelUpdate

rc.Close
Set rc = Nothing

Set oSQL = Nothing

'If Me.MDIChild = True Then frmMenuTactil.Show
Set frmMntOfe = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub



Private Sub cbLista_click()

If rc.EditMode = adEditNone Then

With frmFlexSimple

    .Caption = "Ofertas ..."
        
    With .fg
            Set .DataSource = rc
            .ColFormat(1) = "000"
            DoEvents
            .AutoSize 1, .Cols - 1
            .Refresh
    End With
    
    .Show 1

End With

Else

    MsgBox "Debe guardar o cancelar cambios antes de seleccionar un nuevo registro", vbInformation, "Atención"

End If

End Sub








Private Sub ioDCTO_Validate(Cancel As Boolean)

    If ioDCTO.Text > 100 Then
        lblstatus.Caption = "No se permite DCTO mayor del 100%"
        ioDCTO.Text = 100
        ioDCTO.SetFocus
        Cancel = True
    End If

End Sub

Private Sub rc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  If rc.AbsolutePosition > 0 Then
  
    Call cbTIPO_Validate(False)
    
    lblstatus.Caption = "Registro: " & CStr(rc.AbsolutePosition)
    ioIMPORTE.Text = rc.fields("IMPORTE")
    
  Else
    ioIMPORTE.Text = ""
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
  
  On Error GoTo AddErr
  With rc
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    
    tmpcodigo = devuelve_campo("select max(codigo) + 1 from OFERTAS WHERE CODCAJA = " & CajaActual)
    
    'Si devuelve @ esque ha habido un error
    If tmpcodigo = "@" Then tmpcodigo = 1
    
    .fields("CODIGO") = tmpcodigo
    
    'End If

    Set tmpcodigo = Nothing
    
    lblstatus.Caption = "Agregar Oferta"
    mbAddNewFlag = True
    SetButtons False
    cbActualizar.Visible = True
    cbCAJAS.SetFocus
  End With

    
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cbEliminar_Click()
     
  On Error GoTo cbEliminar_Click_Error

  If MsgBox("¿Desea dar de baja la oferta actual?", vbQuestion + vbYesNo, titulo) = vbNo Then Exit Sub

  With rc
    .fields("mbaja") = True
    .UpdateBatch adAffectAll
    .Requery
    .Move 0
  End With
     
  On Error GoTo 0
  Exit Sub

cbEliminar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbEliminar_Click de Formulario frmMntOfe"

End Sub



Private Sub cbedicion_Click()
  On Error GoTo EditErr

  lblstatus.Caption = "Modificar registro"
  mbEditFlag = True
  SetButtons False
  cbActualizar.Visible = True
  
  cbCAJAS.SetFocus
  
  Exit Sub

EditErr:
  MsgBox Err.Description
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
  'On Error GoTo UpdateErr
   
    
  With cbTIPO
        
        If .Text = "" Then
            lblstatus.Caption = "No se permite TIPO de Oferta en blanco"
            .SetFocus
            Exit Sub
        End If
  
  End With
   
    
  Select Case cbTIPO.Text
  
  Case 3
  
    With ioIMPORTE
        If .Text = "" Then
            lblstatus.Caption = "No se permite IMPORTE en blanco"
            .SetFocus
            Exit Sub
        End If
    End With

  Case 2
  
    With ioDCTO
        If .Text = "" Then
            lblstatus.Caption = "No se permite DESCUENTO en blanco"
            .SetFocus
            Exit Sub
        End If
    End With
    
   End Select
   
    With ioFECHAINI
        If .Text = "" Then
            lblstatus.Caption = "No se permite FECHA INICIO en blanco"
            .SetFocus
            Exit Sub
        End If
    End With
        
    With ioFECHAFIN
        If .Text = "" Then
            lblstatus.Caption = "No se permite FECHA FIN en blanco"
            .SetFocus
            Exit Sub
        End If
    End With
   
  If ioIMPORTE.Visible = True Then
  rc.fields("IMPORTE") = ioIMPORTE.Text
  End If
  rc.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    rc.MoveLast              'va al nuevo registro
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  
  lblstatus.Caption = ""
  
  cbAgregar.SetFocus

  Exit Sub
UpdateErr:
  If Err.Number = -2147217887 Then Exit Sub
  MsgBox Err.Description, vbInformation, "Atención"
End Sub

Private Sub cbcerrar_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  rc.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
If Err.Number = -2147217842 Then Exit Sub
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  rc.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
If Err.Number = -2147217842 Then Exit Sub
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  If Not rc.EOF Then rc.MoveNext
  If rc.EOF And rc.RecordCount > 0 Then
    Beep
     'ha sobrepasado el final; vuelva atrás
    rc.MoveLast
  End If
  'muestra el registro actual
  mbDataChanged = False

  Exit Sub
GoNextError:
If Err.Number = -2147217842 Then Exit Sub
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  If Not rc.BOF Then rc.MovePrevious
  If rc.BOF And rc.RecordCount > 0 Then
    Beep
    'ha sobrepasado el final; vuelva atrás
    rc.MoveFirst
  End If
  'muestra el registro actual
  mbDataChanged = False

  Exit Sub

GoPrevError:
If Err.Number = -2147217842 Then Exit Sub
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cbAgregar.Visible = bVal
  cbEdicion.Visible = bVal
  cbActualizar.Visible = Not bVal
  cbCancelar.Visible = Not bVal
  cbEliminar.Visible = bVal
  cbCerrar.Visible = bVal
  cbLista.Visible = bVal
   
  'cbActualizar.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub
