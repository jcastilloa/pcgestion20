VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMntVal 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vales"
   ClientHeight    =   4335
   ClientLeft      =   2310
   ClientTop       =   330
   ClientWidth     =   8955
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
   ScaleHeight     =   4335
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   Begin PCGestion.chameleonButton cmdPrevious 
      Height          =   630
      Left            =   1065
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2850
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1111
      btype           =   9
      tx              =   "F6"
      enab            =   -1  'True
      font            =   "frmMntVal.frx":0000
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntVal.frx":002C
      picn            =   "frmMntVal.frx":004A
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
      Left            =   15
      Top             =   2445
      Width           =   8940
      _extentx        =   15769
      _extenty        =   661
      caption         =   ""
      fount           =   "frmMntVal.frx":0D1E
      captioncolour   =   0
      colour1         =   15640462
      colour2         =   7177785
      captionalignment=   1
   End
   Begin PCGestion.miText ioIMPORTE 
      Height          =   525
      Left            =   1215
      TabIndex        =   0
      Top             =   1020
      Width           =   1275
      _extentx        =   2249
      _extenty        =   926
      font            =   "frmMntVal.frx":0D4C
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.chameleonButton cmdFirst 
      Height          =   630
      Left            =   30
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2850
      Width           =   1005
      _extentx        =   1773
      _extenty        =   1111
      btype           =   9
      tx              =   "F5"
      enab            =   -1  'True
      font            =   "frmMntVal.frx":0D78
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntVal.frx":0DA4
      picn            =   "frmMntVal.frx":0DC2
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
      Left            =   4665
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3510
      Width           =   1065
      _extentx        =   1879
      _extenty        =   1402
      btype           =   9
      tx              =   "Lista F4"
      enab            =   -1  'True
      font            =   "frmMntVal.frx":1AFA
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntVal.frx":1B26
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
      Left            =   6810
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2850
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1111
      btype           =   9
      tx              =   "F7"
      enab            =   -1  'True
      font            =   "frmMntVal.frx":1B44
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntVal.frx":1B70
      picn            =   "frmMntVal.frx":1B8E
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
      Left            =   7875
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2850
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1111
      btype           =   9
      tx              =   "F8"
      enab            =   -1  'True
      font            =   "frmMntVal.frx":2862
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntVal.frx":288E
      picn            =   "frmMntVal.frx":28AC
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
      Top             =   3510
      Width           =   1065
      _extentx        =   1879
      _extenty        =   1402
      btype           =   9
      tx              =   "&Agregar F1"
      enab            =   -1  'True
      font            =   "frmMntVal.frx":35E4
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntVal.frx":3610
      picn            =   "frmMntVal.frx":362E
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
      TabIndex        =   15
      Top             =   3510
      Width           =   1200
      _extentx        =   2117
      _extenty        =   1402
      btype           =   9
      tx              =   "&Actualizar F2"
      enab            =   -1  'True
      font            =   "frmMntVal.frx":430A
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntVal.frx":4336
      picn            =   "frmMntVal.frx":4354
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
      Left            =   2325
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3510
      Width           =   990
      _extentx        =   1746
      _extenty        =   1402
      btype           =   9
      tx              =   "&Edicion F3"
      enab            =   -1  'True
      font            =   "frmMntVal.frx":4C30
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntVal.frx":4C5C
      picn            =   "frmMntVal.frx":4C7A
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
      Left            =   5760
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3510
      Width           =   930
      _extentx        =   1640
      _extenty        =   1402
      btype           =   9
      tx              =   "&Cancelar"
      enab            =   -1  'True
      font            =   "frmMntVal.frx":54DA
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntVal.frx":5506
      picn            =   "frmMntVal.frx":5524
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
      Left            =   6750
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3510
      Width           =   1065
      _extentx        =   1879
      _extenty        =   1402
      btype           =   9
      tx              =   "E&liminar F9"
      enab            =   -1  'True
      font            =   "frmMntVal.frx":5E00
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntVal.frx":5E2C
      picn            =   "frmMntVal.frx":5E4A
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
      Left            =   7875
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3510
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1402
      btype           =   9
      tx              =   "Cerrar ESC"
      enab            =   -1  'True
      font            =   "frmMntVal.frx":6A1E
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntVal.frx":6A4A
      picn            =   "frmMntVal.frx":6A68
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.miText ioDCTO 
      Height          =   525
      Left            =   3420
      TabIndex        =   1
      Top             =   1020
      Width           =   1275
      _extentx        =   2249
      _extenty        =   926
      font            =   "frmMntVal.frx":7744
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.bsGradientLabel lblCliente 
      Height          =   375
      Left            =   1245
      Top             =   555
      Width           =   7635
      _extentx        =   13467
      _extenty        =   661
      caption         =   ""
      fount           =   "frmMntVal.frx":7770
      captioncolour   =   0
      colour1         =   12632256
      colour2         =   16558731
      captionalignment=   1
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel4 
      Height          =   315
      Left            =   2895
      Top             =   2970
      Width           =   3270
      _extentx        =   5768
      _extenty        =   556
      caption         =   "-C- Asignar Cliente  -N- Nuevo Cliente"
      fount           =   "frmMntVal.frx":779E
      captioncolour   =   0
      colour1         =   12632256
      colour2         =   16558731
   End
   Begin PCGestion.miCombo cbTIPO 
      Height          =   495
      Left            =   5340
      TabIndex        =   2
      Top             =   1020
      Width           =   3555
      _extentx        =   6271
      _extenty        =   873
      font            =   "frmMntVal.frx":77CC
   End
   Begin PCGestion.miText ioCADUCA 
      Height          =   525
      Left            =   1215
      TabIndex        =   3
      Top             =   1530
      Width           =   1530
      _extentx        =   2699
      _extenty        =   926
      font            =   "frmMntVal.frx":77F8
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.chameleonButton cbImprimir 
      Height          =   795
      Left            =   3345
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3510
      Width           =   1140
      _extentx        =   2037
      _extenty        =   1429
      btype           =   9
      tx              =   "&Imprimir"
      enab            =   -1  'True
      font            =   "frmMntVal.frx":7824
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmMntVal.frx":7850
      picn            =   "frmMntVal.frx":786E
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblORIGEN 
      Height          =   330
      Left            =   5700
      Top             =   2070
      Width           =   3240
      _extentx        =   5715
      _extenty        =   582
      caption         =   ""
      fount           =   "frmMntVal.frx":854A
      captioncolour   =   0
      colour1         =   15640462
      colour2         =   11513775
      captionalignment=   1
   End
   Begin PCGestion.bsGradientLabel lblEMITIDO 
      Height          =   330
      Left            =   1245
      Top             =   2055
      Width           =   3615
      _extentx        =   6376
      _extenty        =   582
      caption         =   ""
      fount           =   "frmMntVal.frx":8578
      captioncolour   =   0
      colour1         =   15640462
      colour2         =   11513775
      captionalignment=   1
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ORIGEN"
      Height          =   285
      Left            =   4875
      TabIndex        =   27
      Top             =   2070
      Width           =   780
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "EMITIDO"
      Height          =   285
      Left            =   285
      TabIndex        =   26
      Top             =   2055
      Width           =   870
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CADUCIDAD"
      Height          =   360
      Left            =   60
      TabIndex        =   24
      Top             =   1620
      Width           =   1140
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO"
      Height          =   300
      Left            =   4800
      TabIndex        =   23
      Top             =   1095
      Width           =   465
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENTE"
      Height          =   330
      Left            =   345
      TabIndex        =   22
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DCTO %"
      Height          =   360
      Left            =   2520
      TabIndex        =   21
      Top             =   1095
      Width           =   825
   End
   Begin MSForms.CheckBox ioMBAJA 
      Height          =   435
      Left            =   2790
      TabIndex        =   20
      Top             =   1560
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
      Left            =   6420
      TabIndex        =   8
      Top             =   90
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
      Left            =   1260
      TabIndex        =   7
      Top             =   90
      Width           =   1245
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      Height          =   330
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ultima modificación"
      Height          =   315
      Left            =   4170
      TabIndex        =   5
      Top             =   120
      Width           =   2160
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "IMPORTE"
      Height          =   360
      Left            =   210
      TabIndex        =   4
      Top             =   1110
      Width           =   975
   End
End
Attribute VB_Name = "frmMntVal"
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

Dim TmpUsr As Long

Public Codigo_Vale As Long
Public Caja_Vale As Byte

Private Sub cbImprimir_Click()

If mbEditFlag Or mbAddNewFlag Then Exit Sub

Call Imprime_Vale(rc.fields("CODIGO"), rc.fields("CODCAJA"), locCnn)

End Sub

Private Sub ioCODIGO_Change()
ioCODIGO.Caption = Format(ioCODIGO.Caption, "00000")
End Sub

Private Sub Form_Activate()

If Not prime Then
      
  If rc.RecordCount = 0 Then
        
        If MsgBox("No se encuentran Vales. ¿Crear?", vbYesNo + vbQuestion, "Vales") = vbNo Then
        Unload Me
        Else
        Call cbAgregar_Click
        End If
        
  Else
        Call cmdFirst_Click
        Call cbCancelar_Click
        
  End If
   

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
     
  oSQL.AddTable "VALES"
  oSQL.AddOrderClause "CODIGO"
  
  
  If Codigo_Vale > 0 And Caja_Vale > 0 Then
  
    oSQL.AddSimpleWhereClause "CODIGO", Codigo_Vale
    oSQL.AddSimpleWhereClause "CODCAJA", Caja_Vale
  
  Else
   
    oSQL.AddSimpleWhereClause "CODCAJA", CajaActual
    oSQL.AddSimpleWhereClause "FMODI", Format(Now, "yyyymmdd"), , CLAUSE_GREATERTHANOREQUAL, LOGIC_AND
    
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
  
  With ioIMPORTE
 ' Set .DataSource = rc
        .PermitirBlanco = True
        .SoloNumeros = True
        .dspFormat = "Currency"
        '.DataField = "IMPORTE"
        .LongMaxima = 10
        .Alineacion = 1
  End With
  
  With ioDCTO
'  Set .DataSource = rc
        .PermitirBlanco = True
        .SoloNumeros = True
      '  .DataField = "DCTO"
        .LongMaxima = 2
        .Alineacion = 1
  End With
  
  With ioCADUCA
    .dspFormat = "dd/mm/yyyy"
    .LongMaxima = 10
    .DataField = "CADUCA"
    Set .DataSource = rc
  End With
  
  With cbTIPO
      .ConexionString = locCnn
      .LenCodigo = 1
      
    '  .SQLString = "SELECT CODIGO, NOMBRE FROM TABLA WHERE MBAJA = 0 ORDER BY CODIGO"
       '1=VENTA, 2=DEVOLUCION, 3=SOBRANTE 4=ANULADO
       .añade_item "1  VENTA"
       .añade_item "2  DEVOLUCION"
       .añade_item "3  SOBRANTE"
       .añade_item "4  ANULADO"
       
      .DataField = "TIPO"
     ' .carga
      Set .DataSource = rc
      .CodigoWidth = 300
  End With
  
    With ioFMODI
  Set .DataSource = rc
        .DataField = "FMODI"
  End With
        
  mbDataChanged = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If (mbEditFlag Or mbAddNewFlag) And (KeyCode <> vbKeyC) And (KeyCode <> vbKeyN) Then Exit Sub

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
        
        'Asignar Cliente ...
      Case vbKeyC
      
       
       If Not (mbEditFlag Or mbAddNewFlag) Then Exit Sub
       'abre el grid de los clientes
       Call Abre_Grid_Clientes
       KeyCode = 0
      
      
      'crear nuevo cliente rapido
      Case vbKeyN
        
        If rc.RecordCount <= 0 Then Exit Sub
        With frmNuCliRap
        
            .Show 1
            
            If .ID_Cliente_Creado > 0 Then
            
            If Not (mbEditFlag Or mbAddNewFlag) Then
                rc.fields("CODCLI") = .ID_Cliente_Creado
                rc.fields("CAJACLI") = .Caja_Cliente
            End If
            
            lblCliente.Caption = .RAZO_Creado
            End If
            
        
        End With
        
        Set frmNuCliRap = Nothing
         KeyCode = 0
        
      
      
      
  End Select
  KeyCode = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error Resume Next

If rc.EditMode <> adEditNone Then rc.CancelUpdate

rc.Close
Set rc = Nothing

   'With locCnn
  '  If .State <> 0 Then .Close
  ' End With

Set oSQL = Nothing

'If Me.MDIChild = True Then frmMenuTactil.Show
Set Plantilla = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub



Private Sub cbLista_click()

If rc.EditMode = adEditNone Then

With frmFlexVal

    '.Caption = "Secciones ..."
        
    .desde_mnt = True

    Set .miRc = rc
    
    'Set ioIMPORTE.DataSource = Nothing
    Set ioDCTO.DataSource = Nothing
    Set cbTIPO.DataSource = Nothing
    Set ioCADUCA.DataSource = Nothing
    Set ioCODIGO.DataSource = Nothing
    Set ioFMODI.DataSource = Nothing
    Set ioMBAJA.DataSource = Nothing
    
    .Show 1
    
    DoEvents
    
    Set frmFlexVal = Nothing
    
    Set ioDCTO.DataSource = rc
    Set cbTIPO.DataSource = rc
    Set ioCADUCA.DataSource = rc
    Set ioCODIGO.DataSource = rc
    Set ioFMODI.DataSource = rc
    Set ioMBAJA.DataSource = rc
    
End With

Else

    MsgBox "Debe guardar o cancelar cambios antes de seleccionar un nuevo registro", vbInformation, "Atención"

End If

End Sub

Private Sub ioIMPORTE_Validate(Cancel As Boolean)

    If (ioIMPORTE.Text <> "0") And (ioIMPORTE.Text <> "") Then ioDCTO.Text = "0"

End Sub

Private Sub ioDCTO_Validate(Cancel As Boolean)

    If (ioDCTO.Text <> "0") And (ioDCTO.Text <> "") Then ioIMPORTE.Text = "0"

End Sub

Private Sub rc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  
  'Esto mostrará la posición de registro actual para este Recordset
  If rc.AbsolutePosition > 0 Then
        ioIMPORTE.Text = rc.fields("IMPORTE")
        ioDCTO.Text = rc.fields("DCTO")
        lblstatus.Caption = "Registro: " & CStr(rc.AbsolutePosition)
        If rc.fields("CODCAJA") > 0 Then lblORIGEN.Caption = Trim(devuelve_campo("SELECT DESCRIPCION FROM CAJAS WHERE CODIGO = " & rc.fields("CODCAJA"), locCnn))
        If rc.fields("CODPER") > 0 Then lblEMITIDO.Caption = Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & rc.fields("CODPER"), locCnn))
        
        If rc.fields("TIPO") = 4 Then
            ioMBAJA.Value = True
        Else
            ioMBAJA.Value = False
        End If
        
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
  
          TmpUsr = 0
          
          Do
  
            With frmSelDep
                .Show 1
                TmpUsr = .ID_Dependiente
                Unload frmSelDep
            End With
        
         Loop Until TmpUsr <> 0
  
         Set frmSelDep = Nothing
         
         Me.Caption = "Vales ... Usuario [" & Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & TmpUsr, locCnn)) & "]"
         
  
  With rc
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    
    tmpcodigo = devuelve_campo("select max(codigo) + 1 from VALES WHERE CODCAJA = " & CajaActual, locCnn)
     
    'Si devuelve @ esque ha habido un error
     If tmpcodigo = "@" Then tmpcodigo = 1
    
    .fields("CODIGO") = tmpcodigo
    
   ' End If

    Set tmpcodigo = Nothing
    
    lblstatus.Caption = "Agregar registro"
    mbAddNewFlag = True
    SetButtons False
    cbActualizar.Visible = True
  ioIMPORTE.SetFocus
  End With
  
  cbTIPO.Text = 1

    
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cbEliminar_Click()
    On Error GoTo DeleteErr
  
  'preguntar
  If MsgBox("¿Desea ANULAR el vale seleccionado? (por un importe de: " & Format(rc.fields("IMPORTE"), "Currency") & ")", vbQuestion + vbYesNo, titulo) = vbNo Then Exit Sub
  
  With rc

    .fields("TIPO") = 4
    
    rc.UpdateBatch adAffectAll
    
    .Move 0
  
  End With
 
      
Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub



Private Sub cbedicion_Click()
  On Error GoTo EditErr

  lblstatus.Caption = "Modificar registro"
  mbEditFlag = True
  SetButtons False
  cbActualizar.Visible = True
  
  ioIMPORTE.SetFocus
  
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
Dim tmpcodigo As Long
  On Error GoTo UpdateErr
 
  If ioIMPORTE.Text = "" Then ioIMPORTE.Text = "0"
  If ioDCTO.Text = "" Then ioDCTO.Text = "0"
   
  If CDbl(ioIMPORTE.Text) = 0 And CDbl(ioDCTO.Text) = 0 Then
    lblstatus.Caption = "Importe y DCTO no pueden ser ambos 0"
    ioIMPORTE.SetFocus
    Exit Sub
  End If
  
  If CDbl(ioIMPORTE.Text) <> 0 And CDbl(ioDCTO.Text) <> "0" Then
    lblstatus.Caption = "Debe introducir IMPORTE o DCTO no ambos"
    ioIMPORTE.SetFocus
    Exit Sub
  End If
  
  rc.fields("IMPORTE") = CDbl(ioIMPORTE.Text)
  rc.fields("CODPER") = TmpUsr 'UsuarioActual
  rc.fields("CODCAJA") = CajaActual
  rc.fields("DCTO") = ioDCTO.Text
    
  tmpcodigo = rc.fields("CODIGO")
  
  rc.UpdateBatch adAffectAll

  Call Imprime_Vale(tmpcodigo, CajaActual, locCnn)

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

'---------------------------------------------------------------------------------------
' Subrutina   : Abre_Grid_Clientes
' Fecha/Hora  : 18/01/2004 14:48
' Autor       : JCASTILLO
' Propósito   : Abre el grid de clientes, y obtiene un cliente para la venta
'---------------------------------------------------------------------------------------
Private Sub Abre_Grid_Clientes()
Dim cliSql As New clsSmartSQL
Dim rccli As New ADODB.Recordset


   'On Error GoTo Abre_Grid_Clientes_Error

cliSql.AddTable "CLIENTES"
cliSql.AddOrderClause "CODCAJA"
cliSql.AddOrderClause "CODIGO"

rccli.Open cliSql.SQL, locCnn, adOpenDynamic, adLockReadOnly

With frmFlexCli

    .Caption = "Clientes ..."
    Set .miosql = cliSql
            
    Set .miRc = rccli
       
    DoEvents
  
    
    .Show 1
    
    If .seleccionado Then
    
        'asignar valores ...
        rc.fields("CODCLI") = rccli.fields("CODIGO")
        rc.fields("CAJACLI") = rccli.fields("CODCAJA")
        lblCliente.Caption = devuelve_campo("SELECT RAZO FROM CLIENTES WHERE CODIGO = " & rc.fields("CODCLI") & " AND CODCAJA = " & rc.fields("CAJACLI"), locCnn)
    
    'dejar como estaba
    Else
    
      '  rc.Fields("CODCLI") = Null
      '  rc.Fields("CAJACLI") = Null
      '  lblCliente.Caption = ""
        
    End If
    
        rc.Update
        
    Set frmFlexCli = Nothing
    
    DoEvents
    
End With


rccli.Close
Set rccli = Nothing
Set cliSql = Nothing

   On Error GoTo 0
   Exit Sub

Abre_Grid_Clientes_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Abre_Grid_Clientes de Formulario frmCabVen"

End Sub



'Asigna el cliente seleccionado en el flexgrid, para llamar desde el flexclientes
Public Sub Asignar_cliente_flex(Codigo_Cliente As Long, codcaja As Byte)

With frmFlexCli
    
    If .seleccionado Then
    
        'asignar valores ...
        rc.fields("CODCLI") = Codigo_Cliente 'rccli.Fields("CODIGO")
        rc.fields("CAJACLI") = codcaja 'rccli.Fields("CODCAJA")
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


