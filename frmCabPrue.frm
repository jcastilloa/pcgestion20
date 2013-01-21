VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmCabPrue 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mercancía de Prueba"
   ClientHeight    =   7065
   ClientLeft      =   2235
   ClientTop       =   2250
   ClientWidth     =   10860
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   Begin PCGestion.chameleonButton cmdPrevious 
      Height          =   630
      Left            =   1080
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5580
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1111
      btype           =   9
      tx              =   "F6"
      enab            =   -1  'True
      font            =   "frmCabPrue.frx":0000
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmCabPrue.frx":002C
      picn            =   "frmCabPrue.frx":004A
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
      Top             =   5175
      Width           =   10815
      _extentx        =   17304
      _extenty        =   661
      caption         =   ""
      fount           =   "frmCabPrue.frx":0D1E
      captioncolour   =   0
      colour1         =   15640462
      colour2         =   7177785
      captionalignment=   1
   End
   Begin PCGestion.miText ioCODBAR 
      Height          =   525
      Left            =   1065
      TabIndex        =   0
      Top             =   420
      Width           =   2925
      _extentx        =   5159
      _extenty        =   926
      font            =   "frmCabPrue.frx":0D4C
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin PCGestion.chameleonButton cmdFirst 
      Height          =   630
      Left            =   45
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5580
      Width           =   1005
      _extentx        =   1773
      _extenty        =   1111
      btype           =   9
      tx              =   "F5"
      enab            =   -1  'True
      font            =   "frmCabPrue.frx":0D78
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmCabPrue.frx":0DA4
      picn            =   "frmCabPrue.frx":0DC2
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
      Left            =   7590
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5580
      Width           =   1065
      _extentx        =   1879
      _extenty        =   1111
      btype           =   9
      tx              =   "Lista F4"
      enab            =   -1  'True
      font            =   "frmCabPrue.frx":1AFA
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmCabPrue.frx":1B26
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
      Left            =   8730
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5580
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1111
      btype           =   9
      tx              =   "F7"
      enab            =   -1  'True
      font            =   "frmCabPrue.frx":1B44
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmCabPrue.frx":1B70
      picn            =   "frmCabPrue.frx":1B8E
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
      Left            =   9795
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5580
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1111
      btype           =   9
      tx              =   "F8"
      enab            =   -1  'True
      font            =   "frmCabPrue.frx":2862
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmCabPrue.frx":288E
      picn            =   "frmCabPrue.frx":28AC
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
      Left            =   30
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1065
      _extentx        =   1879
      _extenty        =   1402
      btype           =   9
      tx              =   "&Agregar F1"
      enab            =   -1  'True
      font            =   "frmCabPrue.frx":35E4
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmCabPrue.frx":3610
      picn            =   "frmCabPrue.frx":362E
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
      Left            =   1125
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1200
      _extentx        =   2117
      _extenty        =   1402
      btype           =   9
      tx              =   "&Actualizar F2"
      enab            =   -1  'True
      font            =   "frmCabPrue.frx":430A
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmCabPrue.frx":4336
      picn            =   "frmCabPrue.frx":4354
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
      Left            =   7695
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6240
      Width           =   930
      _extentx        =   1640
      _extenty        =   1402
      btype           =   9
      tx              =   "&Cancelar"
      enab            =   -1  'True
      font            =   "frmCabPrue.frx":4C30
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmCabPrue.frx":4C5C
      picn            =   "frmCabPrue.frx":4C7A
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
      Left            =   8685
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1065
      _extentx        =   1879
      _extenty        =   1402
      btype           =   9
      tx              =   "E&liminar F9"
      enab            =   -1  'True
      font            =   "frmCabPrue.frx":5556
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmCabPrue.frx":5582
      picn            =   "frmCabPrue.frx":55A0
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
      Left            =   9795
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1050
      _extentx        =   1852
      _extenty        =   1402
      btype           =   9
      tx              =   "Cerrar ESC"
      enab            =   -1  'True
      font            =   "frmCabPrue.frx":6174
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmCabPrue.frx":61A0
      picn            =   "frmCabPrue.frx":61BE
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblCliente 
      Height          =   375
      Left            =   5025
      Top             =   480
      Width           =   5790
      _extentx        =   10213
      _extenty        =   661
      caption         =   ""
      fount           =   "frmCabPrue.frx":6E9A
      captioncolour   =   0
      colour1         =   12632256
      colour2         =   16558731
      captionalignment=   1
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel4 
      Height          =   315
      Left            =   3285
      Top             =   5655
      Width           =   3630
      _extentx        =   6403
      _extenty        =   556
      caption         =   "-C- Asignar Cliente  -N- Nuevo Cliente"
      fount           =   "frmCabPrue.frx":6EC8
      captioncolour   =   0
      colour1         =   12632256
      colour2         =   16558731
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel3 
      Height          =   315
      Left            =   3285
      Top             =   6000
      Width           =   3630
      _extentx        =   6403
      _extenty        =   556
      caption         =   "-I- Ir a Rejilla         -B- Borrar Seleccion"
      fount           =   "frmCabPrue.frx":6EF6
      captioncolour   =   0
      colour1         =   12632256
      colour2         =   16558731
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   4215
      Left            =   30
      TabIndex        =   15
      Top             =   900
      Width           =   6300
      _cx             =   11112
      _cy             =   7435
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
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
      BackColorAlternate=   16626604
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
      FormatString    =   $"frmCabPrue.frx":6F24
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
   Begin VSFlex8Ctl.VSFlexGrid fgSel 
      Height          =   4215
      Left            =   6345
      TabIndex        =   16
      Top             =   900
      Width           =   4485
      _cx             =   7911
      _cy             =   7435
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
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
      BackColorAlternate=   16626604
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
      FormatString    =   $"frmCabPrue.frx":6FC1
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
   Begin PCGestion.bsGradientLabel lblUsuario 
      Height          =   375
      Left            =   5040
      Top             =   60
      Width           =   3600
      _extentx        =   6350
      _extenty        =   661
      caption         =   ""
      fount           =   "frmCabPrue.frx":705E
      captioncolour   =   0
      colour1         =   12632256
      colour2         =   16558731
      captionalignment=   1
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel1 
      Height          =   315
      Left            =   3300
      Top             =   6345
      Width           =   3630
      _extentx        =   6403
      _extenty        =   556
      caption         =   "-D- Asignar Dependiente"
      fount           =   "frmCabPrue.frx":708C
      captioncolour   =   0
      colour1         =   12632256
      colour2         =   16558731
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "U.Modi"
      Height          =   330
      Left            =   8655
      TabIndex        =   19
      Top             =   90
      Width           =   720
   End
   Begin VB.Label ioFMODI 
      Alignment       =   2  'Center
      BackColor       =   &H00AC998C&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   9420
      TabIndex        =   18
      Top             =   60
      Width           =   1395
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "USUARIO"
      Height          =   300
      Left            =   4035
      TabIndex        =   17
      Top             =   105
      Width           =   885
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENTE"
      Height          =   300
      Left            =   4050
      TabIndex        =   14
      Top             =   525
      Width           =   855
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
      Left            =   1080
      TabIndex        =   3
      Top             =   45
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      Height          =   330
      Left            =   150
      TabIndex        =   2
      Top             =   75
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CBARRAS"
      Height          =   360
      Left            =   30
      TabIndex        =   1
      Top             =   510
      Width           =   990
   End
End
Attribute VB_Name = "frmCabPrue"
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
Dim totimpor As Currency

Dim oSQL As New clsSmartSQL
Dim nif As New clsNIF

Dim rcdet As ADODB.Recordset
Dim conta_lineas As Long

Dim TmpUsr As Long


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public mCODCLI As Long
Public mCAJACLI As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




Private Sub ioCODBAR_Validate(Cancel As Boolean)
Dim mic As MiCodBar
Dim campos As String
Dim tart As Variant
Dim cadena As String
   
    'On Error GoTo ioCODBAR_Validate_Error
    
    
    If Trim(ioCODBAR.Text) = "" Then Exit Sub
    
    If Len(ioCODBAR.Text) = LenCodBar Then
                
        mic = Descompone_CBAR(ioCODBAR.Text)
    
    ElseIf (Len(Trim(ioCODBAR.Text)) = 1) Then
        
    'si es un codigo de barras con la longitud válidad
    'o un codigo de un digito para los restos
    'RES1
    'buscar por referencia "RES" + el codigo de un digito
    'introducido
    
        
        'comprobar si existe el artículo/temporada
        
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
        
        Else
                 
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
              
        'añadir una linea al grid
        Call añade_linea(CInt(mic.CODIGO_ART), CByte(mic.TEMPORADA_ART), CInt(mic.TALLA_ART), CInt(mic.COLOR_ART), 1)
        ioCODBAR.Text = ""
        
      End If
    

    Cancel = True

   On Error GoTo 0
   Exit Sub

ioCODBAR_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioCODBAR_Validate de Formulario frmCabPrue"

End Sub


'---------------------------------------------------------------------------------------
' Procedimiento : añade_linea
' Fecha/Hora    : 25/03/2004 17:03
' Autor         : JCastillo
' Propósito     :  Añade una linea al grid con los datos introducidos
'---------------------------------------------------------------------------------------
Private Sub añade_linea(codart As Integer, tempor As Byte, codtalla As Integer, codcol As Integer, uds As Single)

Dim tmpart As Variant
Dim tmpcodprov As Variant
Dim tmpcodcolor As Variant

 On Error GoTo añade_linea_Error

    'articulo no existe en la transferencia.
    tmpart = devuelve_matriz("SELECT CODPROV, PREVEN FROM MAARTIC WHERE CODIGO = " & codart & " AND TEMPOR = " & tempor, locCnn)
    totimpor = totimpor + (uds * tmpart(1))
    

    
'    With fg
'
'        '.subtotal flexSTCount, , 6, , vbBlue, vbWhite
'        .AddItem "", 2
'        .TextMatrix(2, 1) = conta_lineas
'        .TextMatrix(2, 2) = Trim(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & tmpart(0), locCnn))
'        .TextMatrix(2, 3) = Trim(devuelve_campo("SELECT REF FROM MAARTIC WHERE CODIGO = " & codart & " AND TEMPOR = " & tempor, locCnn))
'        .TextMatrix(2, 4) = Format(codart, "00000") & " " & Trim(devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & codart & " AND TEMPOR = " & tempor, locCnn))
'        .TextMatrix(2, 5) = Trim(devuelve_campo("SELECT ABREVIA FROM TEMPOR WHERE IDTEM = " & tempor, locCnn))
'
'        If codtalla > 0 Then .TextMatrix(2, 6) = Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & codtalla, locCnn))
'
'        If codcol > 0 Then
'
'            .TextMatrix(2, 7) = Trim(devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & codcol, locCnn))
'            tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & codcol, locCnn)
'
'            If tmpcodcolor <> "@" Then
'                .Row = 2
'                .Col = 7
'                .CellBackColor = tmpcodcolor
'                .Col = 3
'            End If
'
'        End If
'
'         totimpor = totimpor + (uds * tmpart(1))
'
'        .TextMatrix(1, 8) = uds
'        .TextMatrix(1, 9) = tmpart(1)
'
'        .subtotal flexSTSum, , 8, , vbBlue, vbWhite
'        .subtotal flexSTSum, , 9, , vbBlue, vbWhite
'
'      '  .TextMatrix(1, 5) = "Uds: "
'        .TextMatrix(1, 1) = ""
'     '   .TextMatrix(1, 9) = "Total: " & Format(totimpor, "Currency")
'       .TextMatrix(1, 10) = tmpcodigo
'        .AutoSize 1, .Cols - 1
'
'    End With
'
'    With rcdet
'        .AddNew
'        .fields("CODIGO") = rc.fields("CODIGO")
'        .fields("LINEA") = tmpcodigo
'        .fields("CODART") = codart
'        .fields("CODCOL") = codcol
'        .fields("CODTALLA") = codtalla
'        .fields("TEMPOR") = tempor
'        .fields("UNIDADES") = uds
'        .fields("CODCAJA") = rc.fields("CODCAJA")
'        .Update
'    End With
    
    'conta_lineas = conta_lineas + 1
    
    'Call carga_grid(rc.fields("CODIGO"))

   
   
   On Error GoTo 0
   Exit Sub

añade_linea_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento añade_linea de Formulario frmCabPrue"
End Sub



Private Sub ioCODIGO_Change()
ioCODIGO.Caption = Format(ioCODIGO.Caption, "000000000")
End Sub


Private Sub NuCliRapido()
'crear nuevo cliente rapido
        
   On Error GoTo NuCliRapido_Error

        With frmNuCliRap
        
            .Show 1
            
            DoEvents
            Me.SetFocus
            
            If .ID_Cliente_Creado > 0 Then
            rc.fields("CODCLI") = .ID_Cliente_Creado
            rc.fields("CAJACLI") = .Caja_Cliente
            rc.Update
            lblCliente.Caption = .RAZO_Creado
            End If
            
        
        End With
        
        Set frmNuCliRap = Nothing

   On Error GoTo 0
   Exit Sub

NuCliRapido_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento NuCliRapido de Formulario frmCabPrue"
         
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

Me.Visible = False

With frmFlexCli

    .Caption = "Clientes ..."
    Set .miosql = cliSql
            
    .desde_pruebas = True
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



Private Sub Form_Activate()

If Not prime Then

  If rc.RecordCount = 0 Then
        
        If MsgBox("No se encuentran Prestamos de Mercancía. ¿Crear?", vbYesNo + vbQuestion, "Prestamos de Mercancía") = vbNo Then
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
     
   
  Set rcdet = New ADODB.Recordset
  Set rc = New ADODB.Recordset
  oSQL.AddTable "CABPRESTA"
  oSQL.AddOrderClause "CODIGO"
  oSQL.AddSimpleWhereClause "CODCAJA", CajaActual
  'q solo saque las de hoy
  'oSQL.AddSimpleWhereClause "FMODI", Format(Date, "yyyymmdd")
  
  'abrir cabecera
  rc.Open oSQL.SQL, locCnn, adOpenStatic, adLockOptimistic
   
  If rc.RecordCount > 0 Then
  
  'abrir detalle
  If rcdet.State = 1 Then rcdet.Close
    
    rcdet.Open "SELECT * FROM DETPRESTA WHERE CODIGO = " & rc.fields("CODIGO") & " AND CODCAJA = " & CajaActual & " ORDER BY CODIGO", locCnn, adOpenStatic, adLockOptimistic
  
  End If
    
''''''''
'////////////////////////////////////////////////////////
' Enlazar controles:
'////////////////////////////////////////////////////////

  With ioCODIGO
  Set .DataSource = rc
        .DataField = "CODIGO"
  End With
  
  With ioFMODI
  Set .DataSource = rc
        .DataField = "FMODI"
  End With
  
  With fg
        .Clear
        .Cols = 11
        .Rows = 2
        .ColHidden(1) = True
        '.ColHidden(10) = True
        
        .ColAlignment(6) = flexAlignCenterCenter
        .TextMatrix(0, 2) = "Proveedor"
        .TextMatrix(0, 3) = "Referencia"
        .TextMatrix(0, 4) = "Modelo"
        .TextMatrix(0, 5) = "Temp."
        .TextMatrix(0, 6) = "Talla"
        .TextMatrix(0, 7) = "Color"
        .TextMatrix(0, 8) = "Uds."
        .TextMatrix(0, 9) = "PVP"
        .TextMatrix(0, 10) = "Linea"
  End With
                
  mbDataChanged = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  'que deje pasar las teclas para clientes ...
  If (mbEditFlag Or mbAddNewFlag) And ((KeyCode <> vbKeyC) And (KeyCode <> vbKeyN)) Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      
      If MsgBox("¿Desea cerrar la ventana actual?", vbQuestion + vbYesNo, "Cerrar") = vbYes Then
        cbcerrar_Click
      End If
      
'    Case vbKeyEnd
'      cmdLast_Click
'    Case vbKeyHome
'      cmdFirst_Click
'    Case vbKeyUp, vbKeyPageUp
'      If Shift = vbCtrlMask Then
'        cmdFirst_Click
'      Else
'        cmdPrevious_Click
'      End If
'
'    Case vbKeyDown, vbKeyPageDown
'      If Shift = vbCtrlMask Then
'        cmdLast_Click
'      Else
'        cmdNext_Click
'      End If
'
      Case vbKeyF1
            Call cbAgregar_Click
        
      Case vbKeyF2
            Call cbactualizar_Click
        
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
        
    Case vbKeyC
      
       'abre el grid de los clientes
       Call Abre_Grid_Clientes
        KeyCode = 0
      
      
      'crear nuevo cliente rapido
      Case vbKeyN
      
      Call NuCliRapido
      KeyCode = 0
      
    
    Case vbKeyB
    
     
     If fg.Rows <= 1 Then Exit Sub
     If Not IsNumeric(fg.TextMatrix(fg.Row, 10)) Then Exit Sub
     
     If MsgBox("¿Desea borrar el artículo seleccionado?: " & fg.TextMatrix(fg.Row, 4) & ". Linea: " & fg.TextMatrix(fg.Row, 10), vbQuestion + vbYesNo, titulo) = vbNo Then Exit Sub
     locCnn.Execute "DELETE FROM DETPRESTA WHERE CODIGO = " & rc.fields("CODIGO") & " AND CODCAJA = " & rc.fields("CODCAJA") & " AND LINEA =" & fg.TextMatrix(fg.Row, 10)
     
     DoEvents
     
     Call carga_grid(rc.fields("CODIGO"))
     
     KeyCode = 0
     
     
     Case vbKeyD
     
         TmpUsr = 0
         
         Do
         
            With frmSelDep
                .Show 1
                TmpUsr = .ID_Dependiente
                Unload frmSelDep
            End With
        
         Loop Until (TmpUsr <> 0)
         
      
  End Select
'  KeyCode = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error Resume Next

If rc.EditMode <> adEditNone Then rc.CancelUpdate

rc.Close
Set rc = Nothing

rcdet.Close
Set rcdet = Nothing

With locCnn
  If .State <> 0 Then .Close
End With

Set oSQL = Nothing
Set nif = Nothing

''If Me.MDIChild = True Then frmMenuTactil.Show
Set frmCabPrue = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub


Private Sub cbLista_click()

With frmFlexSimple

    .Caption = "Pruebas de Mercancía ..."
        
    With .fg
            Set .DataSource = rc
            .ColFormat(1) = "000000000"
            DoEvents
            .AutoSize 1, .Cols - 1
            .Refresh
    End With
    
    .Show 1

End With

End Sub




'---------------------------------------------------------------------------------------
' Procedimiento : ioNIF_Validate
' Fecha/Hora    : 25/03/2004 16:30
' Autor         : JCastillo
' Propósito     :
'---------------------------------------------------------------------------------------
'
'Private Sub ioNIF_Validate(Cancel As Boolean)
'Dim tmpdat As Variant
'
'   On Error GoTo ioNIF_Validate_Error
'
'If Trim(ioNIF.Text) = "" Then
'   ioNIF.CancelarValidacion
'    Cancel = True
'    Exit Sub
'End If
'
'nif.DarFormato = True
'nif.nif = ioNIF.Text
'
'If nif.Err Then
'    ioNIF.CancelarValidacion
'    Cancel = True
'    Exit Sub
'Else
'    ioNIF.Text = nif.nif
'End If
'
'tmpdat = devuelve_matriz("SELECT CODIGO, CODCAJA FROM CLIENTES WHERE NIF = '" & ioNIF.Text & "'", locCnn)
'
''se ha encontrado
'If IsArray(tmpdat) Then
'
'    'actualizar datos ...
'    rc.fields("CODCLI") = tmpdat(0)
'    rc.fields("CAJACLI") = tmpdat(1)
'    rc.Update
'
'    lblCliente.Caption = devuelve_campo("SELECT RAZO FROM CLIENTES WHERE CODIGO = " & rc.fields("CODCLI") & " AND CODCAJA = " & rc.fields("CAJACLI"), locCnn)
'
''no se ha encontrado
'Else
'
'    lblstatus.Caption = "No se ha encontrado el NIF: " & ioNIF.Text & ". Pulse -C- para crear un nuevo cliente."
'    Cancel = True
'
'End If
'
'
'   On Error GoTo 0
'   Exit Sub
'
'ioNIF_Validate_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioNIF_Validate de Formulario frmCabPrue"
'End Sub

Private Sub rc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  If rc.AbsolutePosition > 0 Then
  
  lblstatus.Caption = "Registro: " & CStr(rc.AbsolutePosition)
  
    If ((rc.fields("CODCLI") > 0 And Not IsNull(rc.fields("CODCLI"))) And (rc.fields("CAJACLI") > 0) And Not IsNull(rc.fields("CAJACLI"))) Then
        lblCliente.Caption = Trim(devuelve_campo("SELECT RAZO FROM CLIENTES WHERE CODIGO = " & rc.fields("CODCLI") & " AND CODCAJA = " & rc.fields("CAJACLI"), locCnn))
    Else
        lblCliente.Caption = ""
    End If
    
    If ((rc.fields("CODUSR") > 0 And Not IsNull(rc.fields("CODUSR")))) Then
        lblUsuario.Caption = Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & rc.fields("CODUSR"), locCnn))
    Else
        lblUsuario.Caption = ""
    End If
    
    With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
    End With
    
    If rc.fields("CODIGO") > 0 Then
        Call carga_grid(rc.fields("CODIGO"))
        'If rcdet.State = 1 Then rcdet.Close
        'rcdet.Open "SELECT * FROM DETPRESTA WHERE CODIGO = " & rc.fields("CODIGO") & " AND CODCAJA = " & CajaActual & " ORDER BY CODIGO", locCnn, adOpenStatic, adLockOptimistic
    End If
         
  Else
  
  lblCliente.Caption = ""
  
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
         
         lblUsuario.Caption = Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & TmpUsr, locCnn))
         DoEvents
  
  
  With rc
  
  
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    
    tmpcodigo = devuelve_campo("select max(codigo) + 1 from CABPRESTA where CODCAJA = " & CajaActual)
    
    If tmpcodigo = "@" Then tmpcodigo = 1
    
    'si ya se le asigna el cliente de manera externa ...
    If (mCODCLI > 0) And (mCAJACLI > 0) Then
    
        .fields("CODCLI") = mCODCLI
        .fields("CAJACLI ") = mCAJACLI
        lblCliente.Caption = devuelve_campo("SELECT RAZO FROM CLIENTES WHERE CODIGO = " & .fields("CODCLI") & " AND CODCAJA = " & .fields("CAJACLI"), locCnn)
            
    End If
        
    .fields("CODIGO") = tmpcodigo
    .fields("CODUSR") = TmpUsr
    .fields("CODCAJA") = CajaActual
   
    Set tmpcodigo = Nothing
    
    lblstatus.Caption = "Agregar registro"
    mbAddNewFlag = True
    SetButtons False
    cbActualizar.Visible = True
    ioCODBAR.SetFocus
    
    If (mCODCLI = 0) And (mCAJACLI = 0) Then Call Form_KeyDown(vbKeyC, 0)

    
    'ioNIF.SetFocus
  End With

    
  Exit Sub
AddErr:
  MsgBox Err.Description
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
  
  ioCODBAR.SetFocus
  'ioNIF.Text = ""

End Sub

Private Sub cbactualizar_Click()
 
   On Error GoTo cbactualizar_Click_Error

  'comprobar que el codigo de cliente no entre a cero
  If (rc.fields("CODCLI") = 0) Or (rc.fields("CAJACLI") = 0) Then
    lblstatus.Caption = "Debe asignarse un cliente (con -C-, -N- o escribiendo el NIF)"
    ioCODBAR.SetFocus
    'ioNIF.SetFocus
    Exit Sub
  End If
    
  rc.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    rc.MoveLast              'va al nuevo registro
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  
  'ioNIF.Text = ""
  
  lblstatus.Caption = ""

   On Error GoTo 0
   Exit Sub

cbactualizar_Click_Error:
    If Err.Number = -2147217887 Then Exit Sub
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbactualizar_Click de Formulario frmCabPrue"

End Sub

Private Sub cbcerrar_Click()
   On Error GoTo cbcerrar_Click_Error

  DoEvents
  Unload Me

   On Error GoTo 0
   Exit Sub

cbcerrar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbcerrar_Click de Formulario frmCabPrue"
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

Private Sub carga_grid(xCODPRES As Long)
Dim tmpart  As Variant
Dim tmpcodcolor  As Variant
  
   On Error GoTo carga_grid_Error

   With fg
        .Clear
        .Cols = 11
        .Rows = 2
        .ColHidden(1) = True
        .ColDataType(3) = flexDTString
        .ColAlignment(3) = flexAlignLeftCenter
        '.ColHidden(10) = True
        .Redraw = flexRDNone
        .TextMatrix(0, 2) = "Proveedor"
        .TextMatrix(0, 3) = "Referencia"
        .TextMatrix(0, 4) = "Modelo"
        .TextMatrix(0, 5) = "Temp."
        .TextMatrix(0, 6) = "Talla"
        .TextMatrix(0, 7) = "Color"
        .TextMatrix(0, 8) = "Uds."
        .TextMatrix(0, 9) = "PVP"
        .TextMatrix(0, 10) = "Linea"
   End With

    conta_lineas = 0
    totimpor = 0
    
      'abrir detalle
     If rcdet.State = 1 Then rcdet.Close
     rcdet.Open "SELECT * FROM DETPRESTA WHERE CODIGO = " & xCODPRES & " AND CODCAJA = " & CajaActual & " ORDER BY CODIGO", locCnn, adOpenStatic, adLockOptimistic

    Do Until rcdet.EOF

        'articulo no existe en la transferencia.
        tmpart = devuelve_matriz("SELECT CODPROV, PREVEN FROM MAARTIC WHERE CODIGO = " & rcdet.fields("CODART") & " AND TEMPOR = " & rcdet.fields("TEMPOR"), locCnn)
      
        'articulo OK, existe.
        fg.AddItem "", 1
        fg.TextMatrix(1, 1) = 0
        fg.TextMatrix(1, 2) = Trim(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & tmpart(0), locCnn))
        
        fg.TextMatrix(1, 3) = Trim(devuelve_campo("SELECT REF FROM MAARTIC WHERE CODIGO = " & rcdet.fields("CODART") & " AND TEMPOR = " & rcdet.fields("TEMPOR"), locCnn))
        
        fg.TextMatrix(1, 4) = Format(rcdet.fields("CODART"), "00000") & " " & Trim(devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & rcdet.fields("CODART") & " AND TEMPOR = " & rcdet.fields("TEMPOR"), locCnn))
        fg.TextMatrix(1, 5) = Trim(devuelve_campo("SELECT ABREVIA FROM TEMPOR WHERE IDTEM = " & rcdet.fields("TEMPOR"), locCnn))
        
        If rcdet.fields("CODTALLA") > 0 Then
            fg.TextMatrix(1, 6) = Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & rcdet.fields("CODTALLA"), locCnn))
        End If
        
        If rcdet.fields("CODCOL") > 0 Then
            fg.TextMatrix(1, 7) = Trim(devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & rcdet.fields("CODCOL"), locCnn))
            tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & rcdet.fields("CODCOL"), locCnn)
            fg.Row = 1
            fg.Col = 7
            fg.CellBackColor = tmpcodcolor
            fg.Col = 3
        End If
        
        totimpor = totimpor + (rcdet.fields("UNIDADES") * tmpart(1))
    
        fg.TextMatrix(1, 8) = rcdet.fields("UNIDADES")
        fg.TextMatrix(1, 9) = tmpart(1)
        
        fg.TextMatrix(1, 10) = rcdet.fields("LINEA")
        
        rcdet.MoveNext
    
    Loop
    
    With fg
        
        .subtotal flexSTSum, , 8, , vbBlue, vbWhite
        .subtotal flexSTSum, , 9, , vbBlue, vbWhite
        
        .AutoSize 1, .Cols - 1
        .Redraw = True
        
        '.TextMatrix(1, 5) = "Uds: "
       ' .TextMatrix(1, 1) = ""
       ' .TextMatrix(1, 9) = "Total: " & Format(totimpor, "Currency")
       
    End With

   On Error GoTo 0
   Exit Sub

carga_grid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_grid de Formulario frmCabPrue"

End Sub

Private Sub SetButtons(bVal As Boolean)
  cbAgregar.Visible = bVal
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
