VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMntDev 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Devoluciones ..."
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8340
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
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin PCGestion.chameleonButton cmdPrevious 
      Height          =   630
      Left            =   1050
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1111
      btype           =   9
      tx              =   "F6"
      enab            =   -1  'True
      font            =   "frmMntDev.frx":0000
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntDev.frx":002C
      picn            =   "frmMntDev.frx":004A
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
      Top             =   2820
      Width           =   8295
      _extentx        =   14631
      _extenty        =   661
      caption         =   ""
      fount           =   "frmMntDev.frx":0D1E
      captioncolour   =   0
      colour1         =   15640462
      colour2         =   7177785
      captionalignment=   1
   End
   Begin PCGestion.chameleonButton cmdFirst 
      Height          =   630
      Left            =   15
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1005
      _extentx        =   1773
      _extenty        =   1111
      btype           =   9
      tx              =   "F5"
      enab            =   -1  'True
      font            =   "frmMntDev.frx":0D4C
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntDev.frx":0D78
      picn            =   "frmMntDev.frx":0D96
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
      Left            =   6210
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1111
      btype           =   9
      tx              =   "F7"
      enab            =   -1  'True
      font            =   "frmMntDev.frx":1ACE
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntDev.frx":1AFA
      picn            =   "frmMntDev.frx":1B18
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
      Left            =   7275
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1111
      btype           =   9
      tx              =   "F8"
      enab            =   -1  'True
      font            =   "frmMntDev.frx":27EC
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntDev.frx":2818
      picn            =   "frmMntDev.frx":2836
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3900
      Width           =   1065
      _extentx        =   1879
      _extenty        =   1402
      btype           =   9
      tx              =   "&Agregar F1"
      enab            =   -1  'True
      font            =   "frmMntDev.frx":356E
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntDev.frx":359A
      picn            =   "frmMntDev.frx":35B8
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
      TabIndex        =   4
      Top             =   3900
      Width           =   1200
      _extentx        =   2117
      _extenty        =   1402
      btype           =   9
      tx              =   "&Actualizar F2"
      enab            =   -1  'True
      font            =   "frmMntDev.frx":4294
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntDev.frx":42C0
      picn            =   "frmMntDev.frx":42DE
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3900
      Width           =   990
      _extentx        =   1746
      _extenty        =   1402
      btype           =   9
      tx              =   "&Edicion F3"
      enab            =   -1  'True
      font            =   "frmMntDev.frx":4BBA
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntDev.frx":4BE6
      picn            =   "frmMntDev.frx":4C04
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
      Left            =   5220
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3900
      Width           =   930
      _extentx        =   1640
      _extenty        =   1402
      btype           =   9
      tx              =   "&Cancelar"
      enab            =   -1  'True
      font            =   "frmMntDev.frx":5464
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntDev.frx":5490
      picn            =   "frmMntDev.frx":54AE
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
      Left            =   6180
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3900
      Width           =   1065
      _extentx        =   1879
      _extenty        =   1402
      btype           =   9
      tx              =   "E&liminar"
      enab            =   -1  'True
      font            =   "frmMntDev.frx":5D8A
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntDev.frx":5DB6
      picn            =   "frmMntDev.frx":5DD4
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
      Left            =   7275
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3900
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1402
      btype           =   9
      tx              =   "Cerrar ESC"
      enab            =   -1  'True
      font            =   "frmMntDev.frx":69A8
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntDev.frx":69D4
      picn            =   "frmMntDev.frx":69F2
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblArticulo 
      Height          =   390
      Left            =   1215
      Top             =   450
      Width           =   7080
      _extentx        =   12621
      _extenty        =   688
      caption         =   ""
      fount           =   "frmMntDev.frx":76CE
      captioncolour   =   0
      colour1         =   12632256
      colour2         =   16558731
      captionalignment=   1
   End
   Begin PCGestion.miText ioCODBAR 
      Height          =   525
      Left            =   1200
      TabIndex        =   0
      Top             =   1365
      Width           =   3000
      _extentx        =   5292
      _extenty        =   926
      font            =   "frmMntDev.frx":76FC
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioIMPORTE 
      Height          =   525
      Left            =   5145
      TabIndex        =   1
      Top             =   1365
      Width           =   1080
      _extentx        =   1905
      _extenty        =   926
      font            =   "frmMntDev.frx":7728
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.chameleonButton cbSeleccionarVenta 
      Height          =   630
      Left            =   2295
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1785
      _extentx        =   3149
      _extenty        =   1111
      btype           =   9
      tx              =   "F9 &Seleccionar Venta"
      enab            =   -1  'True
      font            =   "frmMntDev.frx":7754
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntDev.frx":7780
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cbImprimir 
      Height          =   630
      Left            =   4125
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1905
      _extentx        =   3360
      _extenty        =   1111
      btype           =   9
      tx              =   "&Imprimir Vale"
      enab            =   -1  'True
      font            =   "frmMntDev.frx":779E
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntDev.frx":77CA
      picn            =   "frmMntDev.frx":77E8
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
      Height          =   795
      Left            =   3720
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3915
      Width           =   1065
      _extentx        =   1879
      _extenty        =   1402
      btype           =   9
      tx              =   "Lista F4"
      enab            =   -1  'True
      font            =   "frmMntDev.frx":84C4
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntDev.frx":84F0
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblCliente 
      Height          =   375
      Left            =   1215
      Top             =   900
      Width           =   7080
      _extentx        =   12541
      _extenty        =   661
      caption         =   ""
      fount           =   "frmMntDev.frx":850E
      captioncolour   =   0
      colour1         =   12632256
      colour2         =   16558731
      captionalignment=   1
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel4 
      Height          =   315
      Left            =   2640
      Top             =   2430
      Width           =   3330
      _extentx        =   6376
      _extenty        =   556
      caption         =   "-C- Asignar Cliente  -N- Nuevo Cliente"
      fount           =   "frmMntDev.frx":853C
      captioncolour   =   0
      colour1         =   12632256
      colour2         =   16558731
   End
   Begin PCGestion.bsGradientLabel lblImprimirVale 
      Height          =   315
      Left            =   6015
      Top             =   2430
      Width           =   1455
      _extentx        =   2566
      _extenty        =   556
      caption         =   "-I- Imprimir Vale"
      fount           =   "frmMntDev.frx":856A
      captioncolour   =   0
      colour1         =   12632256
      colour2         =   16558731
   End
   Begin PCGestion.bsGradientLabel lblCodVale 
      Height          =   345
      Left            =   1215
      Top             =   2415
      Width           =   1350
      _extentx        =   2381
      _extenty        =   609
      caption         =   ""
      fount           =   "frmMntDev.frx":8598
      captioncolour   =   0
      colour1         =   12632256
      colour2         =   16558731
      captionalignment=   1
   End
   Begin PCGestion.miText ioMOTIVO 
      Height          =   540
      Left            =   1200
      TabIndex        =   3
      Top             =   1890
      Width           =   7140
      _extentx        =   12594
      _extenty        =   953
      font            =   "frmMntDev.frx":85C6
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.miText ioTICKET 
      Height          =   525
      Left            =   7005
      TabIndex        =   2
      Top             =   1365
      Visible         =   0   'False
      Width           =   1320
      _extentx        =   2328
      _extenty        =   926
      font            =   "frmMntDev.frx":85F2
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TICKET"
      Height          =   360
      Left            =   6150
      TabIndex        =   28
      Top             =   1455
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "VALE"
      Height          =   360
      Left            =   360
      TabIndex        =   27
      Top             =   2430
      Width           =   765
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ARTICULO"
      Height          =   300
      Left            =   120
      TabIndex        =   26
      Top             =   495
      Width           =   1035
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENTE"
      Height          =   300
      Left            =   285
      TabIndex        =   25
      Top             =   945
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "IMPORTE"
      Height          =   360
      Left            =   4125
      TabIndex        =   21
      Top             =   1455
      Width           =   990
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CBARRAS"
      Height          =   360
      Left            =   165
      TabIndex        =   20
      Top             =   1470
      Width           =   975
   End
   Begin MSForms.CheckBox ioMBAJA 
      Height          =   435
      Left            =   7455
      TabIndex        =   19
      Top             =   2370
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
   Begin VB.Label ioFMODI 
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
      Left            =   5850
      TabIndex        =   9
      Top             =   30
      Width           =   2445
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
      Left            =   1200
      TabIndex        =   8
      Top             =   30
      Width           =   1545
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      Height          =   330
      Left            =   285
      TabIndex        =   7
      Top             =   60
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ultima modificación"
      Height          =   315
      Left            =   3645
      TabIndex        =   6
      Top             =   60
      Width           =   2160
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MOTIVO"
      Height          =   360
      Left            =   270
      TabIndex        =   5
      Top             =   1980
      Width           =   870
   End
End
Attribute VB_Name = "frmMntDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Módulo     : frmMntDev
' Fecha/Hora : 26/01/2004 17:58
' Autor      : JCastillo
' Propósito  : Devoluciones de Mercancía
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
Dim TmpUsr  As Long


Public desde_ventas As Boolean
Dim oSQL As New clsSmartSQL


Public codigo_devol As Long
Public Caja_Devol As Byte



Private Sub cbLista_click()

If rc.EditMode = adEditNone Then

With frmFlexDev

    '.Caption = "Secciones ..."
        
    .desde_mnt = True

    Set .miRc = rc
    
    'Set ioIMPORTE.DataSource = Nothing
    Set ioMOTIVO.DataSource = Nothing
    Set ioCODIGO.DataSource = Nothing
    Set ioFMODI.DataSource = Nothing
    Set ioMBAJA.DataSource = Nothing
    
    .Show 1
    
    DoEvents
    
    Set frmFlexVal = Nothing
    
    Set ioMOTIVO.DataSource = rc
    Set ioCODIGO.DataSource = rc
    Set ioFMODI.DataSource = rc
    Set ioMBAJA.DataSource = rc
    
End With

Else

    MsgBox "Debe guardar o cancelar cambios antes de seleccionar un nuevo registro", vbInformation, "Atención"

End If

End Sub

Private Sub cbImprimir_Click()
Dim tmpvale As Long
Dim entrans As Boolean

   On Error GoTo cbImprimir_Click_Error

If mbEditFlag Or mbAddNewFlag Then Exit Sub

'si no se ha creado el vale, hacerlo ahora ...
If rc.fields("CODVAL") = 0 Then

  Set rc.ActiveConnection = Nothing
  
  With locCnn
    .Close
    .CursorLocation = adUseServer
    .Open strLocCnn
    .BeginTrans
    entrans = True
    DoEvents
  End With

      tmpvale = añadir_vale(0, UsuarioActual, rc.fields("CODCLI"), rc.fields("CAJACLI"), rc.fields("IMPORTE"), 0, 0, 2, Null, locCnn)
      lblstatus.Caption = "Se ha creado el vale: " & tmpvale & Format(CajaActual, "000")
      
      'asignarle el codigo de vale a la devolucion actual ...
      locCnn.Execute "UPDATE DEVOL SET CODVAL = " & tmpvale & " WHERE CODIGO = " & rc.fields("CODIGO") & " AND CODCAJA = " & rc.fields("CODCAJA")
      DoEvents
            
    With locCnn
    'realizar cambios
        If entrans Then
            .CommitTrans
        End If
    
         entrans = False
    
        .Close
        .CursorLocation = adUseClient
        .Open strLocCnn
    End With
  
    Set rc.ActiveConnection = locCnn
    DoEvents
  
    rc.Requery
    If Not rc.BOF Then rc.MoveLast

End If

'imprimir el vale actual
Call Imprime_Vale(rc.fields("CODVAL").Value, CajaActual, locCnn)

   On Error GoTo 0
   Exit Sub

cbImprimir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbImprimir_Click de Formulario frmMntDev"

End Sub

Private Sub cbSeleccionarVenta_Click()
   On Error GoTo cbSeleccionarVenta_Click_Error

If (Not mbEditFlag) And (Not mbAddNewFlag) Then Exit Sub

If rc.fields("CODART") <> 0 And rc.fields("TEMPOr") <> 0 Then
    lblstatus.Caption = "Ya se ha asignado un artículo a esta devolución"
    Exit Sub
End If

With frmFlexVen
    .Desde_Devol = True
    .Show 1
    
    If .D_Cancelado = False Then
    
        rc.fields("CODART") = .D_Codart
        rc.fields("TEMPOR") = .D_Tempor
        rc.fields("CODTALLA") = .D_CodTalla
        rc.fields("CODCOL") = .D_CodCol
        rc.fields("MOTIVO") = " "
        
        'sacar el importe unitario
        rc.fields("IMPORTE") = (.D_Importe / .D_Unidades)
        ioIMPORTE.Text = .D_Importe
        
        rc.Update
    
    End If

    Set frmFlexVen = Nothing

End With

   On Error GoTo 0
   Exit Sub

cbSeleccionarVenta_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbSeleccionarVenta_Click de Formulario frmMntDev"

End Sub

Private Sub ioCODBAR_Validate(Cancel As Boolean)
Dim mic As MiCodBar
Dim campos As Variant
Dim tart As Variant
Dim cadena As String

   On Error GoTo ioCODBAR_Click_Error
   
   If Not mbEditFlag And Not mbAddNewFlag Then Exit Sub
    
    
    If Len(ioCODBAR.Text) = LenCodBar Then
        
        
        mic = Descompone_CBAR(ioCODBAR.Text)

    
    ElseIf (Len(Trim(ioCODBAR.Text)) = 1) Then
        
    'si es un codigo de barras con la longitud válidad
    'o un codigo de un digito para los restos
    'RES1
    'buscar por referencia "RES" + el codigo de un digito
    'introducido
    
        
        'comprobar si existe el artículo/temporada
        
        'de aqui quite preven
        cadena = "SELECT MODELO, CODIGO FROM MAARTIC WHERE REF = 'RES" & Trim(ioCODBAR.Text) & "' AND TEMPOR = " & TemporadaActual
       
        tart = devuelve_matriz(cadena, locCnn)
        
        If Not IsArray(tart) Then
        
                lblstatus.Caption = "No existe el artículo para esa temporada!, Codigo de Barras no Válido"
                ioCODBAR.Text = ""
                ioCODBAR.CancelarValidacion
                Cancel = True
                       
                Beep
                Call Espera(1)
                Beep
                Call Espera(1)
                Beep
                
                Exit Sub
           
        End If
                   
           mic.CODIGO_ART = tart(1)
           mic.TEMPORADA_ART = TemporadaActual
           mic.TALLA_ART = "0"
           mic.COLOR_ART = "0"
           
          ' ioIMPORTE.Text = tart(2)
           
           Set tart = Nothing
                       
        Else
               
            
            If Trim(ioCODBAR.Text) = "" Then
                Beep
                lblstatus.Caption = "Atención ha dejado CODIGO DE BARRAS en blanco"
                Exit Sub
            End If
            
            Cancel = True
            Beep
            Call Espera(1)
            Beep
            Call Espera(1)
            Beep
            
            lblstatus.Caption = "Código de Barras Incorrecto"
            Exit Sub
        
        End If


      If (Len(ioCODBAR.Text) = LenCodBar) Or (Len(Trim(ioCODBAR.Text)) = 1) Then
      
        rc.fields("codart") = mic.CODIGO_ART
        rc.fields("tempor") = mic.TEMPORADA_ART
        rc.fields("codtalla") = mic.TALLA_ART
        rc.fields("codcol") = mic.COLOR_ART

        campos = devuelve_matriz("SELECT MODELO, REF, PREVEN FROM MAARTIC WHERE CODIGO = " & mic.CODIGO_ART & " AND TEMPOR = " & mic.TEMPORADA_ART, locCnn)
             'codigo de artículo
             
        If Not IsArray(campos) Then
            lblArticulo.Caption = "Error al leer el artículo"
            Exit Sub
        End If
        
        lblArticulo.Caption = Trim(campos(1)) & " - " & Format(mic.CODIGO_ART, "00000") & " " & Trim(campos(0)) & " " & Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & mic.TALLA_ART, locCnn)) & "  " & Trim(devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & mic.COLOR_ART, locCnn))
        ioIMPORTE.Text = campos(2)
        
        Set campos = Nothing
        Set tart = Nothing
        
      End If
    
   On Error GoTo 0
   Exit Sub

ioCODBAR_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioCODBAR_Click de Formulario frmMntDev"

   On Error GoTo 0
   Exit Sub

ioCODBAR_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioCODBAR_Validate de Formulario frmMntDev"

End Sub

Private Sub ioCODIGO_Change()
ioCODIGO.Caption = Format(ioCODIGO.Caption, "00000000")
End Sub

Private Sub Form_Activate()

If Not prime Then

    
  'If rc.RecordCount = 0 Then
        
        'If MsgBox("No se encuentran Devoluciones. ¿Crear?", vbYesNo + vbQuestion, "Devoluciones") = vbNo Then
        'Unload Me
        'Else
        

      If codigo_devol = 0 Then
        
        Call cbAgregar_Click
        
      End If
        'End If
        
  'Else
    '    Call cmdFirst_Click
    '    Call cbCancelar_Click
        
  'End If

prime = True
frmVerFecha.Show 1
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
  oSQL.AddTable "DEVOL"
  oSQL.AddOrderClause "CODIGO"
  
  If codigo_devol > 0 And Caja_Devol > 0 Then
    
        oSQL.AddSimpleWhereClause "CODCAJA", Caja_Devol
        oSQL.AddSimpleWhereClause "CODIGO", codigo_devol
        
  Else
  
        oSQL.AddSimpleWhereClause "CODCAJA", CajaActual
        'que solo saque las devoluciones del dia de hoy
        oSQL.AddComplexWhereClause "FMODI >= '" & Format(Now, "yyyymmdd") & "'", LOGIC_AND
  
  End If
  
  'oSQL.AddSimpleWhereClause "FMODI", CDate(Format(Now, "yyyymmdd")), , CLAUSE_GREATERTHANOREQUAL

  rc.Open oSQL.SQL, locCnn, adOpenStatic, adLockOptimistic
    

''''''''
'////////////////////////////////////////////////////////
' Enlazar controles:
'////////////////////////////////////////////////////////

  With ioCODIGO
  Set .DataSource = rc
        .DataField = "CODIGO"
  End With
  
  With ioMOTIVO
  Set .DataSource = rc
        '.PermitirBlanco = False
        .DataField = "MOTIVO"
        .LongMaxima = 30
  End With
  
  With ioIMPORTE
'  Set .DataSource = rc
        .dspFormat = "Currency"
        '.PermitirBlanco = False
       ' .DataField = "IMPORTE"
        .LongMaxima = 10
        .Alineacion = 1
  End With
  
  With ioCODBAR
    .LongMaxima = LenCodBar
    .SoloNumeros = True
    .PermitirBlanco = True
  End With
  
  With ioFMODI
  Set .DataSource = rc
        .DataField = "FMODI"
  End With
  
  With ioMBAJA
  Set .DataSource = rc
        .DataField = "MBAJA"
  End With
             
  mbDataChanged = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If (mbEditFlag Or mbAddNewFlag) And (KeyCode <> vbKeyF9) And (KeyCode <> vbKeyC) And (KeyCode <> vbKeyN) And (KeyCode <> vbKeyI) Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      
      If MsgBox("¿Desea cerrar la ventana actual?", vbQuestion + vbYesNo, "Cerrar") = vbYes Then
        cbcerrar_Click
      End If
      
'    Case vbKeyEnd
'      cmdLast_Click
'    Case vbKeyHome
'      cmdFirst_Click
  '  Case vbKeyUp, vbKeyPageUp
 '     If Shift = vbCtrlMask Then
'        cmdFirst_Click'
  '    Else
 '       cmdPrevious_Click
'      End If
      
'    Case vbKeyDown, vbKeyPageDown
'      If Shift = vbCtrlMask Then
'        cmdLast_Click
'      Else
'        cmdNext_Click
'      End If
      
      Case vbKeyF1
            Call cbAgregar_Click
        
      Case vbKeyF2
            Call cbactualizar_Click
        
      Case vbKeyF3
            Call cbedicion_Click
        
      'Case vbKeyF4
            'Call cbSeleccionaArticulo_Click
      
      Case vbKeyF5
            Call cmdFirst_Click
    
       Case vbKeyF6
            Call cmdPrevious_Click
      
       Case vbKeyF7
            Call cmdNext_Click
    
       Case vbKeyF8
        Call cmdLast_Click
        
        Case vbKeyF9
        Call cbSeleccionarVenta_Click
        
         'Asignar Cliente ...
      Case vbKeyC
      
       If Screen.ActiveControl.Name = "ioMOTIVO" Then Exit Sub
       
       If Not (mbEditFlag Or mbAddNewFlag) Then
        lblstatus.Caption = "No esta creando ninguna devolucion en este momento"
        KeyCode = 0
        Exit Sub
       End If
       
       If (IsNull(rc.fields("CODART")) Or IsNull(rc.fields("TEMPOR"))) Then
        lblstatus.Caption = "Debe especificar un artículo antes"
        KeyCode = 0
        Exit Sub
       End If
            
       If mbAddNewFlag And ioCODBAR.Text = "" Then
        lblstatus.Caption = "Debe especificar un artículo antes"
        KeyCode = 0
        Exit Sub
       End If
    
       'abre el grid de los clientes
       Call Abre_Grid_Clientes
        KeyCode = 0
        lblstatus.Caption = ""
        
      'crear nuevo cliente rapido
      Case vbKeyN
        
        'If Not (mbEditFlag Or mbAddNewFlag) Then Exit Sub
        
        If Screen.ActiveControl.Name = "ioMOTIVO" Then Exit Sub
        
        If rc.RecordCount <= 0 Then
            KeyCode = 0
            Exit Sub
        End If
        
       If (mbEditFlag Or mbAddNewFlag) Then
       
       If (IsNull(rc.fields("CODART")) Or IsNull(rc.fields("TEMPOR"))) Then
        lblstatus.Caption = "Debe especificar un artículo antes"
        KeyCode = 0
        Exit Sub
       End If
        
       End If
        
        With frmNuCliRap
        
            .Show 1
            
            DoEvents
            Me.SetFocus
            
            If .ID_Cliente_Creado > 0 Then
            
                If (mbEditFlag Or mbAddNewFlag) Then
            
                    rc.fields("CODCLI") = .ID_Cliente_Creado
                    rc.fields("CAJACLI") = .Caja_Cliente
                    rc.fields("MOTIVO") = "."
                    rc.Update
                    lblCliente.Caption = .RAZO_Creado
                    
                End If
                        
            End If
            
        
        End With
        
        Set frmNuCliRap = Nothing
         KeyCode = 0
        
        
        Case vbKeyI
     
        If Screen.ActiveControl.Name = "ioMOTIVO" Then Exit Sub
        
        If Not (mbEditFlag Or mbAddNewFlag) Then Call cbImprimir_Click
        KeyCode = 0
     
            
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

desde_ventas = False
Set frmMntDev = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub


Private Sub ioMOTIVO_Validate(Cancel As Boolean)
    'If Not mbEditFlag And Not mbAddNewFlag Then Exit Sub
    'Call cbactualizar_Click
End Sub







'Private Sub cbSeleccionaArticulo_Click()
'Dim sqlart As New clsSmartSQL
''Dim miRc As New ADODB.Recordset
'
    
'    With frmFlexArt
'
'    sqlart.AddTable "MAARTIC"
'    sqlart.AddSimpleWhereClause "MBAJA", 0
'
'    miRc.Open sqlart.SQL, locCnn, adOpenDynamic, adLockOptimistic
'
'    Set .miosql = sqlart
'
'    With .fg
'            .ColFormat(1) = "00000"
'             Set frmFlexArt.miRc = miRc
'    End With
'
'        .Show 1
    
'    End With
'
'    miRc.Close
'    Set miRc = Nothing
'
'    Exit Sub
'
'If rc.EditMode = adEditNone Then

'With frmFlexSimple
'
'    .Caption = "Secciones ..."
        
'    With .fg
'            Set .DataSource = rc
'            .ColFormat(1) = "000"
'            DoEvents
'            .AutoSize 1, .Cols - 1
'            .Refresh
'    End With
'
'    .Show 1
'
'End With

'Else
'
'    MsgBox "Debe guardar o cancelar cambios antes de seleccionar un nuevo registro", vbInformation, "Atención"
'
'End If

'End Sub


Private Sub rc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Dim campos As String
  'Esto mostrará la posición de registro actual para este Recordset
  If rc.AbsolutePosition > 0 Then
  
  lblstatus.Caption = "Registro: " & CStr(rc.AbsolutePosition)
  ioIMPORTE.Text = rc.fields("IMPORTE")
  
  If rc.fields("CODART") > 0 And rc.fields("TEMPOR") > 0 Then
  
        campos = devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & rc.fields("CODART") & " AND TEMPOR = " & rc.fields("TEMPOR"), locCnn)
        lblArticulo.Caption = Format(rc.fields("CODART"), "00000") & " " & campos & " " & Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & rc.fields("CODTALLA"), locCnn)) & "  " & Trim(devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & rc.fields("CODCOL"), locCnn))
        
  Else
        
        lblArticulo.Caption = ""
  
  End If
  
  If ((rc.fields("CODCLI") > 0 And Not IsNull(rc.fields("CODCLI"))) And (rc.fields("CAJACLI") > 0) And Not IsNull(rc.fields("CAJACLI"))) Then
        lblCliente.Caption = devuelve_campo("SELECT RAZO FROM CLIENTES WHERE CODIGO = " & rc.fields("CODCLI") & " AND CODCAJA = " & rc.fields("CAJACLI"), locCnn)
  Else
        lblCliente.Caption = ""
  End If
  
  If (rc.fields("CODVAL") > 0) And Not IsNull(rc.fields("CODVAL")) Then
    lblCodVale.Caption = rc.fields("CODVAL") & Format(rc.fields("CODCAJA"), "000")
  Else
    lblCodVale.Caption = ""
  End If
  
  
  End If
campos = ""

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
          
         TmpUsr = 0
         
         Do
         
            With frmSelDep
                .Show 1
                TmpUsr = .ID_Dependiente
                Unload frmSelDep
            End With
        
         Loop Until TmpUsr <> 0
  
         Set frmSelDep = Nothing
         
         Me.Caption = "Devoluciones ... Usuario [" & Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & TmpUsr, locCnn)) & "]"
         
  
  With rc
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    
    tmpcodigo = devuelve_campo("select max(codigo) + 1 from DEVOL where CODCAJA =" & CajaActual)
    
    'Si devuelve @ esque ha habido un error
    If tmpcodigo = "@" Then If tmpcodigo = "@" Then tmpcodigo = 1
    
    .fields("CODIGO") = tmpcodigo
    .fields("CODCAJA") = CajaActual
    '.fields("CODUSR") = UsuarioActual
    
    'End If

    Set tmpcodigo = Nothing
    
    lblstatus.Caption = "Agregar registro"
    mbAddNewFlag = True
    SetButtons False
    cbActualizar.Visible = True
    ioCODBAR.Visible = True
    
    ioCODBAR.Text = ""
    ioCODBAR.SetFocus
  End With

    
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cbEliminar_Click()
Dim tmpnumval As Long
Dim tmpcodart As Integer
Dim tmptempor As Byte
Dim tmpcodtalla As Integer
Dim tmpcodcol As Integer
Dim entrans As Boolean

  On Error GoTo DeleteErr
    
  'si no quiere salir
  If MsgBox("¿Desea cancelar la devolución actual?. Un importe de: " & ioIMPORTE.Text & "." & Chr(13) & "Nota: se quitara una unidad del almacén y si ha creado un vale, se borrará de la base de datos.", vbQuestion + vbYesNo, titulo) = vbNo Then Exit Sub
  
  With rc
    '.Delete
    '.MoveNext
    tmpnumval = .fields("CODVAL")
    
    tmpcodart = .fields("CODART")
    tmptempor = .fields("TEMPOR")
    tmpcodtalla = .fields("CODTALLA")
    tmpcodcol = .fields("CODCOL")
    
    
    .fields("mbaja") = True
  '  .Fields("FBAJA") = Date
    
    'guardar cambios en el registro
    rc.UpdateBatch adAffectAll
    DoEvents
  
    If .EOF Then .MoveLast
  End With
 
     'preparar para la transacción:
  Set rc.ActiveConnection = Nothing
    
  With locCnn
    .Close
    .CursorLocation = adUseServer
    .Open strLocCnn
    .BeginTrans
    entrans = True
    DoEvents
  End With
  
  'quitar las unidades al stock (se cancela la devolución)
  Call stock(tmpcodart, tmptempor, tmpcodtalla, tmpcodcol, AlmacenActual, 1, False, locCnn)
 
  'borrar el vale correspondiente a esta devolución ...
  If tmpnumval > 0 Then locCnn.Execute "DELETE FROM VALES WHERE CODIGO = " & tmpnumval & " AND CODCAJA = " & CajaActual
  
  With locCnn
    'realizar cambios
    .CommitTrans
    
    entrans = False
    
    .Close
    .CursorLocation = adUseClient
    .Open strLocCnn
  End With
  
  Set rc.ActiveConnection = locCnn
  DoEvents
  
Exit Sub
DeleteErr:

  If entrans Then locCnn.RollbackTrans
  
  MsgBox Err.Description
End Sub



Private Sub cbedicion_Click()
  On Error GoTo EditErr

  lblstatus.Caption = "Modificar registro"
  mbEditFlag = True
  SetButtons False
  cbActualizar.Visible = True
  
  ioCODBAR.Visible = False
  ioMOTIVO.SetFocus
  
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
  
  If desde_ventas Then
    Unload Me
  End If

End Sub

Private Sub cbactualizar_Click()
Dim entrans As Boolean
Dim tmpcodigo As Long
Dim tmpcodcli As Variant
Dim tmpcajacli As Variant
    
   On Error GoTo cbactualizar_Click_Error

  If ioIMPORTE.Text = "" Then ioIMPORTE.Text = "0"
    
  If ioIMPORTE.Text = "0" Then
    lblstatus.Caption = "El importe no puede ser 0"
    ioIMPORTE.SetFocus
    Exit Sub
  End If
  
 If ioMOTIVO.Text = "" Then
    lblstatus.Caption = "No se permite MOTIVO en blanco"
    ioMOTIVO.SetFocus
    Exit Sub
  End If
  
  If (rc.fields("CODART") = 0) Then
    lblstatus.Caption = "Debe asignar un artículo para la devolución"
    ioCODBAR.SetFocus
    Exit Sub
  End If
  
  If rc.fields("MOTIVO") = "" Then rc.fields("MOTIVO") = " "
  
  tmpcajacli = rc.fields("CAJACLI")
  tmpcodcli = rc.fields("CODCLI")
    
  rc.fields("IMPORTE") = ioIMPORTE.Text
 ' rc.Fields("CODCAJA") = CajaActual
  
  rc.fields("CODUSR") = TmpUsr 'UsuarioActual
  
  'recoger el codigo para el update de codval
  tmpcodigo = rc.fields("CODIGO")
  
  'guardar cambios en el registro
  rc.UpdateBatch adAffectAll
  DoEvents
    
  'si se esta creando el registro (luego se puede modificar el importe pero no el artículo)...
  If mbAddNewFlag Then
  
  'preparar para la transacción:
  Set rc.ActiveConnection = Nothing
    
  With locCnn
    .Close
    .CursorLocation = adUseServer
    .Open strLocCnn
    .BeginTrans
    entrans = True
    DoEvents
  End With
  
    'devolver 1 unidad de STOCK para ese artículo
    'si la condición de salida es 4 (condición de error), deshacer los cambios
    If stock(rc.fields("CODART"), rc.fields("TEMPOR"), rc.fields("CODTALLA"), rc.fields("CODCOL"), AlmacenActual, 1, True, locCnn) = 4 Then
           
      Call cbCancelar_Click
    
    'si todo esta ok, crear un vale
    Else
    
      'añadir un vale de devolución
      
    '  tmpvale = añadir_vale(0, UsuarioActual, tmpcodcli, tmpcajacli, CDbl(ioIMPORTE.Text), 0, 0, 2, Null, locCnn)
    lblstatus.Caption = "Se ha devuelto el artículo al almacén"
      
      'asignarle el codigo de vale a la devolucion actual ...
    '  locCnn.Execute "UPDATE DEVOL SET CODVAL = " & tmpvale & " WHERE CODIGO = " & tmpcodigo & " AND CODCAJA = " & CajaActual
     ' DoEvents
      
    End If
    
  
  
  With locCnn
    'realizar cambios
    
    If entrans Then
    .CommitTrans
    End If
    
    entrans = False
    
    .Close
    .CursorLocation = adUseClient
    .Open strLocCnn
  End With
  
  Set rc.ActiveConnection = locCnn
  DoEvents
  
  rc.Requery
  
  End If
    
  DoEvents

  If mbAddNewFlag Then
    rc.MoveLast              'va al nuevo registro
    
  End If
  
  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  
  ioCODBAR.Text = ""
  lblstatus.Caption = ""
  
  If desde_ventas Then
    Unload Me
  End If
  
  
  If Me.Visible Then cbAgregar.SetFocus
 
   On Error GoTo 0
   Exit Sub

cbactualizar_Click_Error:

    If entrans Then
    
        With locCnn
            .RollbackTrans
            .Close
            .CursorLocation = adUseClient
            .Open strLocCnn
        End With
            
        Set rc.ActiveConnection = locCnn
        DoEvents
        rc.Requery
            
    End If
    
    If Err.Number = -2147217887 Then Exit Sub
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbactualizar_Click de Formulario frmMntDev"
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
  lblImprimirVale.Visible = bVal
  Label8.Visible = bVal
  lblCodVale.Visible = bVal
   
  cbEdicion.Visible = bVal
  cbActualizar.Visible = Not bVal
  cbCancelar.Visible = Not bVal
  cbEliminar.Visible = bVal
  cbCerrar.Visible = bVal
  
'  cbSeleccionaArticulo.Visible = Not bVal
  cbSeleccionarVenta.Visible = Not bVal
  cbLista.Visible = bVal
  cbImprimir.Visible = bVal
   
  'cbActualizar.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
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
            
    .Desde_Devol = True
    
    '.desde_ventas = True
    Set .miRc = rccli
       
    DoEvents
  
    Me.Visible = False
  
    '.MDIChild = True
    .Show
        
    

    'Set frmFlexCli = Nothing
    
    DoEvents
    
End With

   Set cliSql = Nothing
   
  ' rccli.Close
  ' Set rccli = Nothing

   On Error GoTo 0
   Exit Sub


   On Error GoTo 0
   Exit Sub

Abre_Grid_Clientes_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Abre_Grid_Clientes de Formulario frmMntDev"

End Sub


'Asigna el cliente seleccionado en el flexgrid, para llamar desde el flexclientes
Public Sub Asignar_cliente_flex(Codigo_Cliente As Long, codcaja As Byte)

With frmFlexCli
    
    If .seleccionado Then
    
        'asignar valores ...
        rc.fields("CODCLI") = Codigo_Cliente 'rccli.Fields("CODIGO")
        rc.fields("CAJACLI") = codcaja 'rccli.Fields("CODCAJA")
        rc.fields("MOTIVO") = " "
        lblCliente.Caption = devuelve_campo("SELECT RAZO FROM CLIENTES WHERE CODIGO = " & rc.fields("CODCLI") & " AND CODCAJA = " & rc.fields("CAJACLI"), locCnn)
    
    'dejar como estaba
    Else
    
      '  rc.Fields("CODCLI") = Null
      '  rc.Fields("CAJACLI") = Null
      '  lblCliente.Caption = ""
        
    End If
    
End With
    
        rc.Update
        

End Sub


Private Sub Borrar_Apunte_De_Venta()

End Sub



