VERSION 5.00
Begin VB.Form frmMovCaja 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movimientos de Caja ..."
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7425
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
   ScaleHeight     =   3390
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optSalida 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Salida"
      Height          =   300
      Left            =   4680
      TabIndex        =   20
      Top             =   1065
      Width           =   990
   End
   Begin VB.OptionButton optEntrada 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Entrada"
      Height          =   300
      Left            =   3450
      TabIndex        =   19
      Top             =   1065
      Width           =   1215
   End
   Begin PCGestion.chameleonButton cmdPrevious 
      Height          =   630
      Left            =   1080
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1905
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   1111
      BTYPE           =   9
      TX              =   "F6"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
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
      MICON           =   "FrmMovCaja.frx":0000
      PICN            =   "FrmMovCaja.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblstatus 
      Height          =   375
      Left            =   30
      Top             =   1470
      Width           =   7365
      _ExtentX        =   12991
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
   Begin PCGestion.miText ioComentario 
      Height          =   525
      Left            =   1500
      TabIndex        =   0
      Top             =   480
      Width           =   5925
      _ExtentX        =   10451
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
   Begin PCGestion.chameleonButton cmdFirst 
      Height          =   630
      Left            =   45
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1905
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   1111
      BTYPE           =   9
      TX              =   "F5"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
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
      MICON           =   "FrmMovCaja.frx":0CEE
      PICN            =   "FrmMovCaja.frx":0D0A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbLista 
      Height          =   630
      Left            =   3240
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1875
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1111
      BTYPE           =   9
      TX              =   "Lista F4"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
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
      MICON           =   "FrmMovCaja.frx":1A40
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cmdNext 
      Height          =   630
      Left            =   5265
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1905
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   1111
      BTYPE           =   9
      TX              =   "F7"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
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
      MICON           =   "FrmMovCaja.frx":1A5C
      PICN            =   "FrmMovCaja.frx":1A78
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cmdLast 
      Height          =   630
      Left            =   6330
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1905
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   1111
      BTYPE           =   9
      TX              =   "F8"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
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
      MICON           =   "FrmMovCaja.frx":274A
      PICN            =   "FrmMovCaja.frx":2766
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbAgregar 
      Height          =   795
      Left            =   30
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2565
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
      MICON           =   "FrmMovCaja.frx":349C
      PICN            =   "FrmMovCaja.frx":34B8
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
      TabIndex        =   13
      Top             =   2565
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
      MICON           =   "FrmMovCaja.frx":4192
      PICN            =   "FrmMovCaja.frx":41AE
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
      Left            =   2340
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2565
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
      MICON           =   "FrmMovCaja.frx":4A88
      PICN            =   "FrmMovCaja.frx":4AA4
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
      Left            =   4215
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2565
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
      MICON           =   "FrmMovCaja.frx":5302
      PICN            =   "FrmMovCaja.frx":531E
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
      Left            =   5205
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2565
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
      MICON           =   "FrmMovCaja.frx":5BF8
      PICN            =   "FrmMovCaja.frx":5C14
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
      Left            =   6330
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2565
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
      MICON           =   "FrmMovCaja.frx":67E6
      PICN            =   "FrmMovCaja.frx":6802
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.miText ioIMPORTE 
      Height          =   525
      Left            =   1500
      TabIndex        =   1
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "IMPORTE"
      Height          =   360
      Left            =   420
      TabIndex        =   18
      Top             =   1050
      Width           =   1020
   End
   Begin VB.Label ioFCIERRE 
      Alignment       =   2  'Center
      BackColor       =   &H00EEA78E&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3900
      TabIndex        =   6
      Top             =   75
      Width           =   1245
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
      Left            =   1515
      TabIndex        =   5
      Top             =   75
      Width           =   1605
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NUMERO"
      Height          =   330
      Left            =   600
      TabIndex        =   4
      Top             =   105
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA"
      Height          =   315
      Left            =   3135
      TabIndex        =   3
      Top             =   105
      Width           =   720
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CONCEPTO"
      Height          =   360
      Left            =   270
      TabIndex        =   2
      Top             =   585
      Width           =   1185
   End
End
Attribute VB_Name = "FrmMovCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Modulo      : FrmMovCaja
' Fecha/Hora  : 14/02/2004 14:14
' Autor       : JCASTILLO
' Propósito   : Introducir movimientos de caja
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

Dim tmpcodigo As Variant

Dim oSQL As New clsSmartSQL

Private Sub ioCODIGO_Change()
ioCODIGO.Caption = Format(ioCODIGO.Caption, "000000000")
End Sub

Private Sub Form_Activate()

If Not prime Then

  If rc.RecordCount = 0 Then
        
        If MsgBox("No se encuentran Movimientos. ¿Crear?", vbYesNo + vbQuestion, "Movimientos") = vbNo Then
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
  oSQL.AddTable "MOVCAJA"
  oSQL.AddSimpleWhereClause "CODCAJA", CajaActual
  oSQL.AddSimpleWhereClause "FCIERRE", Format(Date, "yyyymmdd")
  oSQL.AddOrderClause "CODIGO"
  rc.Open oSQL.SQL, locCnn, adOpenStatic, adLockOptimistic
    

''''''''
'////////////////////////////////////////////////////////
' Enlazar controles:
'////////////////////////////////////////////////////////

  With ioCODIGO
  Set .DataSource = rc
        .DataField = "CODIGO"
  End With
  
  With ioComentario
  Set .DataSource = rc
        .PermitirBlanco = False
        .DataField = "COMENTARIO"
        .LongMaxima = 50
  End With
  
  With ioIMPORTE
  'Set .DataSource = rc
        .PermitirBlanco = False
        .dspFormat = "Currency"
       ' .DataField = "COMENTARIO"
        .LongMaxima = 10
        .Alineacion = 1
  End With
  
  With ioFCIERRE
  Set .DataSource = rc
        .DataField = "FCIERRE"
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

Set FrmMovCaja = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub



Private Sub cbLista_click()
Dim tmpstrcombo As String
Dim tmpstrcaja As String
Dim tmprc As New ADODB.Recordset

'If rc.EditMode = adEditNone Then

   On Error GoTo cbLista_click_Error

With tmprc
    .Open "SELECT CODIGO, NOMBRE FROM PERSONAL ORDER BY CODIGO", locCnn, adOpenDynamic, adLockReadOnly
    tmpstrcombo = frmFlexSimple.fg.BuildComboList(tmprc, "NOMBRE", "CODIGO", vbBlue)
    .Close
    .Open "SELECT CODIGO, DESCRIPCION FROM CAJAS ORDER BY CODIGO", locCnn, adOpenDynamic, adLockReadOnly
    tmpstrcaja = frmFlexSimple.fg.BuildComboList(tmprc, "DESCRIPCION", "CODIGO", vbBlue)
    .Close
End With

With frmFlexSimple

    .Caption = "Movimientos Caja ..."
        
    With .fg
            Set .DataSource = rc
            .ColFormat(1) = "000000000"
            .ColFormat(4) = "Currency"
            .ColComboList(2) = tmpstrcaja
            .ColComboList(3) = tmpstrcombo
            .TextMatrix(0, 2) = "CAJA"
            .TextMatrix(0, 3) = "USUARIO"
            DoEvents
            .subtotal flexSTSum, , 4, , vbBlue, vbWhite, True
            .subtotal flexSTCount, , 2, , vbBlue, vbWhite, True
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 4) = "Total(" & .TextMatrix(1, 2) & ") " & .TextMatrix(1, 4)
            .TextMatrix(1, 2) = ""
            .AutoSize 1, .Cols - 1
            .Refresh
    End With
    
    .Show 1
    
    DoEvents
    Set frmFlexSimple = Nothing

End With

'Else

'    MsgBox "Debe guardar o cancelar cambios antes de seleccionar un nuevo registro", vbInformation, "Atención"

'End If

tmpstrcombo = ""
tmpstrcaja = ""
Set tmprc = Nothing

   On Error GoTo 0
   Exit Sub

cbLista_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbLista_click de Formulario FrmMovCaja"

End Sub


Private Sub rc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  If rc.AbsolutePosition > 0 Then
     
    If rc.fields("IMPORTE") <> 0 Then
        ioIMPORTE.Text = rc.fields("IMPORTE")
        ioFCIERRE.Caption = Format(rc.fields("FCIERRE"), "dd/mm/yyyy")
    Else
        ioIMPORTE.Text = 0
        ioFCIERRE.Caption = ""
    End If
    
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
  
  
  On Error GoTo AddErr
  With rc
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    
   tmpcodigo = devuelve_campo("select max(codigo) + 1 from MOVCAJA where CODCAJA =" & CajaActual, locCnn)

    .AddNew
    
   
   


       
    lblstatus.Caption = "Agregar registro"
    mbAddNewFlag = True
    SetButtons False
    cbActualizar.Visible = True
  ioComentario.SetFocus
  End With

    
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cbEliminar_Click()
   


   On Error GoTo cbEliminar_Click_Error

If rc.RecordCount <= 0 Then Exit Sub

If MsgBox("¿Desea borrar el movimiento de caja actual?", vbQuestion + vbYesNo, titulo) = vbNo Then Exit Sub
 
rc.Delete
  
If rc.RecordCount > 0 Then Call cmdFirst_Click
   
Exit Sub


   On Error GoTo 0
   Exit Sub

cbEliminar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbEliminar_Click de Formulario FrmMovCaja"

End Sub



Private Sub cbedicion_Click()
  On Error GoTo EditErr

  lblstatus.Caption = "Modificar registro"
  mbEditFlag = True
  SetButtons False
  cbActualizar.Visible = True
  
  ioComentario.SetFocus
  
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

  
  
   On Error GoTo cbactualizar_Click_Error

  'comprobaciones de importe
  
  If ioIMPORTE.Text = "" Then
    lblstatus.Caption = "Importe no puede estar en blanco"
    ioIMPORTE.SetFocus
    Exit Sub
  End If
  
  If ioIMPORTE.Text = 0 Then
    lblstatus.Caption = "Importe no puede ser cero"
    ioIMPORTE.SetFocus
    Exit Sub
  End If
  
  If (optEntrada.Value = False) And (optSalida.Value = False) Then
    lblstatus.Caption = "Debe seleccionar Entrada o Salida"
    Exit Sub
  End If
    
  If tmpcodigo = "@" Then tmpcodigo = 1

  rc.fields("LINEA") = 0
  rc.fields("CODIGO") = tmpcodigo
    
  'End If
 
  If optEntrada.Value Then
  rc.fields("IMPORTE") = ioIMPORTE.Text
  Else 'si es una salida, poner el importe como negativo
  rc.fields("IMPORTE") = (ioIMPORTE.Text * -1)
  End If
  
  rc.fields("CODCAJA") = CajaActual
  rc.fields("CODUSR") = UsuarioActual
  rc.fields("FCIERRE") = Date
 
  rc.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    rc.MoveLast              'va al nuevo registro
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  
  lblstatus.Caption = ""
  Set tmpcodigo = Nothing
  
  cbAgregar.SetFocus

   On Error GoTo 0
   Exit Sub

cbactualizar_Click_Error:
    If Err.Number = -2147217887 Then Exit Sub
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbactualizar_Click de Formulario FrmMovCaja"
 
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
