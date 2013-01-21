VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmFlexArre 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar Arreglo Existente ..."
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11460
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
   ScaleHeight     =   6915
   ScaleWidth      =   11460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   5415
      Left            =   15
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1500
      Visible         =   0   'False
      Width           =   11445
      _cx             =   20188
      _cy             =   9551
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
      FormatString    =   $"frmFlexArre.frx":0000
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
      Height          =   315
      Left            =   9465
      TabIndex        =   8
      Top             =   1185
      Width           =   1260
      _extentx        =   2223
      _extenty        =   556
      caption         =   "&Consultar"
      font            =   "frmFlexArre.frx":00DE
      image           =   "frmFlexArre.frx":010A
   End
   Begin PCGestion.ucGrdBttn cbBorrar 
      Height          =   315
      Left            =   10710
      TabIndex        =   9
      Top             =   1185
      Width           =   750
      _extentx        =   1323
      _extenty        =   556
      caption         =   "&Borrar"
      font            =   "frmFlexArre.frx":0128
      image           =   "frmFlexArre.frx":0154
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel4 
      Height          =   300
      Left            =   5160
      Top             =   1185
      Width           =   4305
      _extentx        =   7594
      _extenty        =   529
      caption         =   "-F4- Consultar -F5- Ir a Rejilla  -F8- Salir"
      fount           =   "frmFlexArre.frx":0172
      captioncolour   =   0
      colour1         =   15640462
      colour2         =   7177785
      captionalignment=   1
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   1455
      Left            =   15
      TabIndex        =   10
      Top             =   30
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   2566
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Hoja 1"
      TabPicture(0)   =   "frmFlexArre.frx":01A0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label9"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ioFECHAFIN"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ioFECHA"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ioMODELO"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ioNOMBRE"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Hoja 2"
      TabPicture(1)   =   "frmFlexArre.frx":01BC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cbImprimir"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cbESTADO"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "ioIMPORTE"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cbCAJAS"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin PCGestion.miCombo cbCODTALLA 
         Height          =   495
         Left            =   -74220
         TabIndex        =   11
         Top             =   30
         Width           =   2625
         _extentx        =   4630
         _extenty        =   873
         font            =   "frmFlexArre.frx":01D8
      End
      Begin PCGestion.miCombo cbCODCOL 
         Height          =   465
         Left            =   -70905
         TabIndex        =   12
         Top             =   30
         Width           =   3405
         _extentx        =   6006
         _extenty        =   820
         font            =   "frmFlexArre.frx":0204
      End
      Begin PCGestion.miCombo cbCATTALL 
         Height          =   495
         Left            =   -68910
         TabIndex        =   13
         Top             =   525
         Width           =   4155
         _extentx        =   5821
         _extenty        =   873
         font            =   "frmFlexArre.frx":0230
      End
      Begin PCGestion.miCombo cbFAMILIA 
         Height          =   480
         Left            =   -73995
         TabIndex        =   14
         Top             =   45
         Width           =   4155
         _extentx        =   7329
         _extenty        =   847
         font            =   "frmFlexArre.frx":025C
      End
      Begin PCGestion.miCombo cbTIPO 
         Height          =   480
         Left            =   -69840
         TabIndex        =   15
         Top             =   165
         Width           =   3210
         _extentx        =   5662
         _extenty        =   847
         font            =   "frmFlexArre.frx":0288
      End
      Begin PCGestion.miText ioNOMBRE 
         Height          =   495
         Left            =   1395
         TabIndex        =   0
         Top             =   75
         Width           =   4080
         _extentx        =   7197
         _extenty        =   873
         font            =   "frmFlexArre.frx":02B4
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioMODELO 
         Height          =   495
         Left            =   6405
         TabIndex        =   1
         Top             =   75
         Width           =   3780
         _extentx        =   6668
         _extenty        =   873
         font            =   "frmFlexArre.frx":02E0
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioFECHA 
         Height          =   495
         Left            =   1395
         TabIndex        =   2
         Top             =   585
         Width           =   1410
         _extentx        =   2487
         _extenty        =   873
         font            =   "frmFlexArre.frx":030C
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioFECHAFIN 
         Height          =   495
         Left            =   4035
         TabIndex        =   3
         Top             =   585
         Width           =   1425
         _extentx        =   2514
         _extenty        =   873
         font            =   "frmFlexArre.frx":0338
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miCombo cbCAJAS 
         Height          =   495
         Left            =   -69870
         TabIndex        =   5
         Top             =   60
         Width           =   4140
         _extentx        =   7303
         _extenty        =   873
         font            =   "frmFlexArre.frx":0364
      End
      Begin PCGestion.miText ioIMPORTE 
         Height          =   495
         Left            =   -71865
         TabIndex        =   4
         Top             =   75
         Width           =   1260
         _extentx        =   2223
         _extenty        =   873
         font            =   "frmFlexArre.frx":0390
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miCombo cbESTADO 
         Height          =   495
         Left            =   -69555
         TabIndex        =   6
         Top             =   570
         Width           =   3825
         _extentx        =   6747
         _extenty        =   873
         font            =   "frmFlexArre.frx":03BC
      End
      Begin PCGestion.chameleonButton cbImprimir 
         Height          =   795
         Left            =   -74535
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   120
         Width           =   780
         _extentx        =   1376
         _extenty        =   1402
         btype           =   9
         tx              =   ""
         enab            =   -1  'True
         font            =   "frmFlexArre.frx":03E8
         coltype         =   1
         focusr          =   -1  'True
         bcol            =   11513775
         bcolo           =   11513775
         fcol            =   0
         fcolo           =   0
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmFlexArre.frx":0414
         picn            =   "frmFlexArre.frx":0432
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   2
         ngrey           =   0   'False
         fx              =   1
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CAJA"
         Height          =   330
         Left            =   -70470
         TabIndex        =   32
         Top             =   150
         Width           =   540
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IMPORTE"
         Height          =   330
         Left            =   -72810
         TabIndex        =   31
         Top             =   165
         Width           =   885
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO"
         Height          =   330
         Left            =   -70470
         TabIndex        =   30
         Top             =   660
         Width           =   810
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA FIN"
         Height          =   330
         Left            =   2880
         TabIndex        =   29
         Top             =   660
         Width           =   1110
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA INI"
         Height          =   330
         Left            =   270
         TabIndex        =   28
         Top             =   645
         Width           =   1065
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "COSTURERA"
         Height          =   330
         Left            =   150
         TabIndex        =   27
         Top             =   135
         Width           =   1200
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "MODELO"
         Height          =   330
         Left            =   5505
         TabIndex        =   26
         Top             =   150
         Width           =   870
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA FINAL"
         Height          =   330
         Left            =   -71835
         TabIndex        =   25
         Top             =   180
         Width           =   1395
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F.INICIAL"
         Height          =   285
         Left            =   -69960
         TabIndex        =   24
         Top             =   585
         Width           =   1035
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TALLA"
         Height          =   300
         Left            =   -74925
         TabIndex        =   23
         Top             =   120
         Width           =   690
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "COLOR"
         Height          =   285
         Left            =   -71610
         TabIndex        =   22
         Top             =   105
         Width           =   735
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO"
         Height          =   345
         Left            =   -67290
         TabIndex        =   21
         Top             =   120
         Width           =   810
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "EN"
         Height          =   300
         Left            =   -74670
         TabIndex        =   20
         Top             =   540
         Width           =   450
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CAT. TALLA"
         Height          =   330
         Left            =   -70245
         TabIndex        =   19
         Top             =   615
         Width           =   1305
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FAMILIA"
         Height          =   315
         Left            =   -74985
         TabIndex        =   18
         Top             =   120
         Width           =   960
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO"
         Height          =   330
         Left            =   -74730
         TabIndex        =   17
         Top             =   225
         Width           =   810
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO"
         Height          =   330
         Left            =   -70395
         TabIndex        =   16
         Top             =   255
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmFlexArre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim first As Boolean

Dim tmpstrcombo As String

Public miosql As New clsSmartSQL
Public masql As New clsSmartSQL
Public artsql As New clsSmartSQL

Public miRc As New ADODB.Recordset
Public seleccionado As Boolean

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Si entra desde ventas, asignar al arreglo seleccionado el código de venta
'actual
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public desde_ventas As Boolean
Public Venta_Actual As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Dim nif As New clsNIF

Private Sub cbBorrar_click()

'ioFECHA.Text =
ioNOMBRE.Text = ""
ioFECHA.Text = Date
ioMODELO.Text = ""
ioNOMBRE.Text = ""
cbCAJAS.Text = CajaActual
cbESTADO.Text = "1"

Call cbLista_click

End Sub

Private Sub cbCAJAS_Validate(Cancel As Boolean)
Call cbLista_click
End Sub



Private Sub cbESTADO_LostFocus()

ioNOMBRE.SetFocus

End Sub

Private Sub cbImprimir_Click()

Dim linea1 As String
Dim linea2 As String
         
   On Error GoTo cbImprimir_Click_Error

    DoEvents

    linea1 = "Informe de Arreglos. Costurera: " & ioNOMBRE.Text & ". F.Inicial: " & ioFECHA.Text & ". F.Final: " & ioFECHAFIN.Text
    linea2 = "Usuario: " & Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & UsuarioActual, locCnn)) & ". Fecha: " & Now
    

    Call PrintFlexGrid(fg, 1, 1, 2, linea1, linea2, 13, 2)
    'fg.SaveGrid "c:\prueba.txt", flexFileCommaText
    'fg.PrintGrid "Transferencia " & rc.Fields("CODIGO"), , 2, 600, 600
    
    
    'actualizar los colores del grid, volivendolo a cargar
    'rc.Move 0

   On Error GoTo 0
   Exit Sub

cbImprimir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbImprimir_Click de Formulario frmFlexArre"

End Sub

Private Sub cbLista_click()
Dim usa_where As Boolean
Dim nuefech As String

   On Error GoTo cbLista_click_Error

miosql.ClearWhereClause
masql.ClearWhereClause
artsql.ClearWhereClause

If (ioFECHA.Text <> "") And IsDate(ioFECHA.Text) Then

    'miosql.AddSimpleWhereClause "FALTA", ioFECHA.Text, , CLAUSE_GREATERTHANOREQUAL
    'miosql.AddSimpleWhereClause "FALTA", CStr(DateAdd("d", 1, ioFECHA.Text)), , CLAUSE_LESSTHAN, LOGIC_AND
    'miOsql.AddComplexWhereClause "Year(FALTA IN (" & masql.SQL & ")", LOGIC_AND
    
    '>= q el dia actual
    '< que el dia siguiente
    
    If Trim(ioFECHAFIN.Text <> "") And IsDate(ioFECHAFIN.Text) Then
        nuefech = ioFECHAFIN.Text
    Else
        nuefech = DateAdd("d", 1, ioFECHA.Text)
    End If
        
    miosql.AddComplexWhereClause "(FMODI >= '" & Format(Year((ioFECHA.Text)), "0000") & Format(Month((ioFECHA.Text)), "00") & Format(Day((ioFECHA.Text)), "00") & "' AND FMODI < '" & Format(Year((nuefech)), "0000") & Format(Month((nuefech)), "00") & Format(Day((nuefech)), "00") & "')", LOGIC_AND
    usa_where = True
         
End If

If ioIMPORTE.Text <> "" Then
    If CDbl(ioIMPORTE.Text) > 0 Then
        miosql.AddSimpleWhereClause "PVP", CDbl(ioIMPORTE.Text), , , LOGIC_AND
        usa_where = True
    End If
End If

If ioNOMBRE.Text <> "" Then
    masql.AddSimpleWhereClause "NOMBRE", ioNOMBRE.Text, , CLAUSE_LIKE
    usa_where = True
End If

If cbCAJAS.Text <> "" Then
    miosql.AddSimpleWhereClause "CODCAJ", CByte(cbCAJAS.Text), , , LOGIC_AND
    usa_where = True
End If

If cbESTADO.Text <> "" Then
    miosql.AddSimpleWhereClause "ESTADO", CByte(cbESTADO.Text), , , LOGIC_AND
    usa_where = True
End If

If ioMODELO.Text <> "" Then
    artsql.AddSimpleWhereClause "MODELO", ioMODELO.Text, , CLAUSE_LIKE, LOGIC_AND
    usa_where = True
End If

'si deja todo en blanco, no mostrar ningun registro
If Not usa_where Then
    fg.Clear
    Exit Sub
End If

If ioMODELO.Text <> "" Then miosql.AddComplexWhereClause "CODART IN (" & artsql.SQL & ")", LOGIC_AND
If ioNOMBRE.Text <> "" Then miosql.AddComplexWhereClause "CODCOST IN (" & masql.SQL & ")", LOGIC_AND

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
Dim conta_lineas As Long

   On Error GoTo carga_grid_Error

   With fg
   
    .Clear
    .Cols = 13
    .ColFormat(8) = "Currency"
    .ColHidden(0) = True
    .ColHidden(1) = True
    .Rows = 1
    
    .TextMatrix(0, 2) = "Fecha"
    .TextMatrix(0, 3) = "Modelo"
    .TextMatrix(0, 4) = "Talla"
    .TextMatrix(0, 5) = "Color"
    .TextMatrix(0, 6) = "Motivo"
    .TextMatrix(0, 7) = "Costurera"
    .TextMatrix(0, 8) = "PVP"
    .TextMatrix(0, 9) = "Ticket"
    .TextMatrix(0, 10) = "Usuario"
    .TextMatrix(0, 11) = "Caja"
    .TextMatrix(0, 12) = "Estado"
    
    
    If miRc.RecordCount <= 0 Then Exit Sub
        
    Do
             .Rows = .Rows + 1
    
        If Not miRc.EOF Then
     
            conta_lineas = conta_lineas + 1
         
            'ID
            .TextMatrix(.Rows - 1, 0) = conta_lineas
            'CAJA
            .TextMatrix(.Rows - 1, 1) = miRc.fields("CODCAJ")
            
            'FECHA
            .TextMatrix(.Rows - 1, 2) = miRc.fields("FALTA")
            
            'MODELO
            If Not IsNull(miRc.fields("CODART")) Then
            .TextMatrix(.Rows - 1, 3) = Format(miRc.fields("CODART"), "00000") & "-" & Trim(devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & miRc.fields("CODART") & " AND TEMPOR = " & miRc.fields("TEMPOR")))
            Else
            .TextMatrix(.Rows - 1, 3) = "Arreglo Varios"
            End If
            
            'TALLA
            If Not IsNull(miRc.fields("CODTALLA")) Then _
            .TextMatrix(.Rows - 1, 4) = Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & miRc.fields("CODTALLA")))
            
            
            'obtener el texto del color y su codigo de color (para colorear
            'la celda del grid)
            'COLOR
            If Not IsNull(miRc.fields("CODCOL")) Then
                
                tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & miRc.fields("CODCOL"))
                .TextMatrix(.Rows - 1, 5) = Trim(devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & miRc.fields("CODCOL")))
                .Col = 5
                .Row = .Rows - 1
                .CellBackColor = tmpcodcolor
                .Col = 2
            
            End If
            
            'MOTIVO
            .TextMatrix(.Rows - 1, 6) = Trim(miRc.fields("DESCRIPCION"))
            
            'COSTURERA
            If Not IsNull(miRc.fields("CODCOST")) Then .TextMatrix(.Rows - 1, 7) = devuelve_campo("SELECT NOMBRE FROM COSTURE WHERE CODIGO = " & miRc.fields("CODCOST"), locCnn)
            
            'precio de venta
            .TextMatrix(.Rows - 1, 8) = miRc.fields("PVP")
            
            
            .TextMatrix(.Rows - 1, 9) = miRc.fields("CODVEN") & Format(miRc.fields("CODCAJ"), "000")
            
            'usuario
            .TextMatrix(.Rows - 1, 10) = Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & miRc.fields("CODUSR"), locCnn))
            
             'caja
            .TextMatrix(.Rows - 1, 11) = Trim(devuelve_campo("SELECT DESCRIPCION FROM CAJAS WHERE CODIGO = " & miRc.fields("CODCAJ"), locCnn))
            
            '1=PENDIENTE, 2=SERVIDO, 3=CANCELADO
            Select Case miRc.fields("ESTADO")
            
            Case 1
                      .TextMatrix(.Rows - 1, 12) = "PENDIENTE"
            Case 2
                      .TextMatrix(.Rows - 1, 12) = "SERVIDO"
            Case 3
                      .TextMatrix(.Rows - 1, 12) = "CANCELADO"
                      
            End Select
                        
            'estado
            '.TextMatrix(.Rows - 1, 10) = miRc.Fields("ESTADO")
            
     
        End If
    
    If Not miRc.EOF Then miRc.MoveNext
    
    Loop Until miRc.EOF
          
        .SubtotalPosition = flexSTAbove
        .subtotal flexSTCount, -1, 4, , vbBlue, vbWhite
        .subtotal flexSTSum, -1, 8, , vbBlue, vbWhite
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 3) = "Nº Arreglos: (" & conta_lineas & ")"
        .TextMatrix(1, 8) = "Total: " & Format(.TextMatrix(1, 8), "Currency")
        .TextMatrix(1, 4) = ""
        
    .AutoSize 1, .Cols - 1
    .Redraw = True

  End With
  
   On Error GoTo 0
   Exit Sub

carga_grid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_grid de Formulario frmFlexArre"
 
 End Sub


Private Sub fg_dblClick()
Dim IArr As Long

    seleccionado = True

If miRc.State = 0 Then Exit Sub
If miRc.RecordCount <= 0 Then Exit Sub

    seleccionado = True
    
If fg.Rows <= 1 Then Exit Sub

If IsNumeric(fg.TextMatrix(fg.Row, 0)) Then
        
        IArr = fg.TextMatrix(fg.Row, 0)

        'posicionarse en el registro
        miRc.Move (fg.TextMatrix(fg.Row, 0) - 1), 1
        
        If desde_ventas Then
        
            'preguntar al usuario ...
            If MsgBox("¿Desea asignar el arreglo seleccionado? " & Chr(13) & _
            "Modelo: " & fg.TextMatrix(fg.Row, 3) & Chr(13) & _
            "Talla: " & fg.TextMatrix(fg.Row, 4) & Chr(13) & _
            "Color: " & fg.TextMatrix(fg.Row, 5) & Chr(13) & _
            "Importe: " & fg.TextMatrix(fg.Row, 8), vbQuestion + vbYesNo, titulo) = vbYes Then
        
                
                'si el arreglo ya ha sido previamente asignado a otra venta, salir
                If miRc.fields("CODVEN") > 0 Then
                
                    MsgBox "El arreglo seleccionado ya ha sido asignado a otra venta", vbExclamation, titulo
                    Exit Sub
                    
                End If
                
   '             loccnn.Execute "UPDATE ARREGLOS SET CODVEN = " & venta_actual & " WHERE CODCAJA
                miRc.fields("CODVEN") = Venta_Actual
                miRc.Update
                
                DoEvents
                
                miRc.Close
                
                DoEvents
                
                Unload Me
        
            End If
        
            Exit Sub
            
        End If
        
        DoEvents
                
End If

'si corresponde a algun ID
If fg.TextMatrix(fg.Row, 0) <> "" Then

With frmNuArr

    .Solo_Actualizar = True
    .Sel_Caja = fg.TextMatrix(fg.Row, 1)
    .Sel_ID = fg.TextMatrix(fg.Row, 0)
    
    .ioNOMBRE.Text = fg.TextMatrix(fg.Row, 7)
    .ioDESCRIPCION.Text = fg.TextMatrix(fg.Row, 6)
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
    
    Call fg_dblClick
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
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode

'Ir al grid, o regresar
Case vbKeyF5
    
    If fg.Rows > 1 Then
        If fg.TabStop Then
            fg.TabStop = False
            ioNOMBRE.SetFocus
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
    DoEvents
    .Text = CajaActual
  End With
  
  With ioFECHA
    .dspFormat = "dd/mm/yyyy"
   .LongMaxima = 10
   .Text = Date
  End With
  
    With ioFECHAFIN
    .dspFormat = "dd/mm/yyyy"
   .LongMaxima = 10
  End With
  
    With ioIMPORTE
    .dspFormat = "Currency"
   .LongMaxima = 10
   .Alineacion = 1
  End With
  
 With ioMODELO
   .LongMaxima = 30
 End With
  
 With ioNOMBRE
     .LongMaxima = 50
 End With
  
 With cbESTADO
    .añade_item "1   - PENDIENTE"
    .añade_item "2   - SERVIDO"
    .añade_item "3   - CANCELADO"
    .LenCodigo = 1
    .CodigoWidth = 300
    .Text = "2"
 End With
 
    
   Select Case TipoPermiso
   
   'usuario comun, ver solo los pedidos de su almacén
   Case 0
        cbCAJAS.Enabled = False
   'supervisor, ver todos los pedidos
  ' Case 1
   
   End Select
  
 artsql.AddTable "MAARTIC"
 artsql.AddField "CODIGO"
 masql.AddTable "COSTURE"
 miosql.AddTable "ARREGLOS"
 masql.AddField "CODIGO"
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    tmpstrcombo = ""
    'Set nif = Nothing
        
    'si hemos establecido un filtro que no devuelve ningun registro,
    'borrar filtro para que no de error al volver al formulario
    'If miRc.EOF Then Call cbBorrar_click
    
    'No descargar desde aqui, descargar desde el formulario desde donde
    'se llame
    desde_ventas = False
    Set frmFlexArre = Nothing
    
End Sub



Private Sub ioFECHA_Validate(Cancel As Boolean)

If ioFECHA.Text <> "" Then Call cbLista_click

End Sub



Private Sub ioIMPORTE_GotFocus()

If Tab1.Tab <> 1 Then Tab1.Tab = 1

End Sub

Private Sub ioNOMBRE_GotFocus()

If Tab1.Tab > 0 Then Tab1.Tab = 0

End Sub

Private Sub ioNOMBRE_Validate(Cancel As Boolean)

If ioNOMBRE.Text <> "" Then Call cbLista_click

End Sub

Private Sub miCombo1_Click()

End Sub

