VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmFlexPtrans 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Peticiones de Transferencia ..."
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11430
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
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   11430
   StartUpPosition =   2  'CenterScreen
   Begin PCGestion.bsGradientLabel bsGradientLabel4 
      Height          =   285
      Left            =   5115
      Top             =   1095
      Width           =   3495
      _extentx        =   6165
      _extenty        =   503
      caption         =   "-F4- Consultar -F5- Ir a Rejilla  -F8- Salir"
      fount           =   "frmFlexPtrans.frx":0000
      captioncolour   =   0
      colour1         =   15640462
      colour2         =   7177785
      captionalignment=   1
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   5460
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1395
      Width           =   11430
      _cx             =   20161
      _cy             =   9631
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
      FormatString    =   $"frmFlexPtrans.frx":002E
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
   Begin PCGestion.ucGrdBttn cbBorrar 
      Height          =   315
      Left            =   9870
      TabIndex        =   20
      Top             =   1080
      Width           =   750
      _extentx        =   1323
      _extenty        =   556
      caption         =   "&Borrar"
      font            =   "frmFlexPtrans.frx":010C
      image           =   "frmFlexPtrans.frx":0138
   End
   Begin PCGestion.ucGrdBttn cbLista 
      Height          =   315
      Left            =   8625
      TabIndex        =   21
      Top             =   1080
      Width           =   1260
      _extentx        =   2223
      _extenty        =   556
      caption         =   "&Consultar"
      font            =   "frmFlexPtrans.frx":0156
      image           =   "frmFlexPtrans.frx":0182
   End
   Begin PCGestion.ucGrdBttn cmCerrar 
      Height          =   315
      Left            =   10605
      TabIndex        =   22
      Top             =   1080
      Width           =   810
      _extentx        =   1429
      _extenty        =   556
      caption         =   "C&errar"
      font            =   "frmFlexPtrans.frx":01A0
      image           =   "frmFlexPtrans.frx":01CC
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1380
      Left            =   0
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   2434
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
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
      TabPicture(0)   =   "frmFlexPtrans.frx":01EA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ioNUMPED"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ioCODIGO"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ioFECHAFIN"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ioFECHAINI"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cbCODALMDEST"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cbCODALMORIG"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Hoja 2"
      TabPicture(1)   =   "frmFlexPtrans.frx":0206
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "ckDirectas"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "chameleonButton1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "ioCODBAR"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cbEstado"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "ioULTIMAS"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      Begin PCGestion.miCombo cbCODALMORIG 
         Height          =   495
         Left            =   1005
         TabIndex        =   0
         Top             =   30
         Width           =   4575
         _extentx        =   8070
         _extenty        =   873
         font            =   "frmFlexPtrans.frx":0222
      End
      Begin PCGestion.miCombo cbCODALMDEST 
         Height          =   495
         Left            =   6615
         TabIndex        =   1
         Top             =   45
         Width           =   4695
         _extentx        =   8281
         _extenty        =   873
         font            =   "frmFlexPtrans.frx":024E
      End
      Begin PCGestion.miText ioULTIMAS 
         Height          =   525
         Left            =   -67155
         TabIndex        =   8
         Top             =   60
         Width           =   945
         _extentx        =   1667
         _extenty        =   926
         font            =   "frmFlexPtrans.frx":027A
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miCombo cbEstado 
         Height          =   480
         Left            =   -70935
         TabIndex        =   7
         Top             =   45
         Width           =   3255
         _extentx        =   5741
         _extenty        =   847
         font            =   "frmFlexPtrans.frx":02A6
      End
      Begin PCGestion.miText ioCODBAR 
         Height          =   525
         Left            =   -74250
         TabIndex        =   6
         Top             =   45
         Width           =   2355
         _extentx        =   4498
         _extenty        =   926
         font            =   "frmFlexPtrans.frx":02D2
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioFECHAINI 
         Height          =   525
         Left            =   6630
         TabIndex        =   4
         Top             =   555
         Width           =   1635
         _extentx        =   4498
         _extenty        =   926
         font            =   "frmFlexPtrans.frx":02FE
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioFECHAFIN 
         Height          =   525
         Left            =   9690
         TabIndex        =   5
         Top             =   570
         Width           =   1635
         _extentx        =   4498
         _extenty        =   926
         font            =   "frmFlexPtrans.frx":032A
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioCODIGO 
         Height          =   525
         Left            =   1020
         TabIndex        =   2
         Top             =   525
         Width           =   1515
         _extentx        =   2672
         _extenty        =   926
         font            =   "frmFlexPtrans.frx":0356
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.miText ioNUMPED 
         Height          =   525
         Left            =   3270
         TabIndex        =   3
         Top             =   540
         Width           =   1515
         _extentx        =   2672
         _extenty        =   926
         font            =   "frmFlexPtrans.frx":0382
         dspformat       =   ""
         enabled         =   -1  'True
         espassword      =   -1  'True
      End
      Begin PCGestion.chameleonButton chameleonButton1 
         Height          =   555
         Left            =   -65925
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   735
         _extentx        =   1296
         _extenty        =   979
         btype           =   9
         tx              =   ""
         enab            =   -1  'True
         font            =   "frmFlexPtrans.frx":03AE
         coltype         =   1
         focusr          =   -1  'True
         bcol            =   11513775
         bcolo           =   11513775
         fcol            =   0
         fcolo           =   0
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmFlexPtrans.frx":03DA
         picn            =   "frmFlexPtrans.frx":03F8
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   0
         ngrey           =   0   'False
         fx              =   0
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
      End
      Begin MSForms.CheckBox ckDirectas 
         Height          =   405
         Left            =   -74265
         TabIndex        =   23
         Top             =   585
         Width           =   3975
         VariousPropertyBits=   746588179
         BackColor       =   12632256
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "7011;714"
         Value           =   "0"
         Caption         =   "Ver solo directas desde pedido"
         FontName        =   "Trebuchet MS"
         FontHeight      =   225
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PEDIDO"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2325
         TabIndex        =   19
         Top             =   630
         Width           =   930
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Left            =   75
         TabIndex        =   18
         Top             =   615
         Width           =   930
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA FINAL"
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
         Left            =   8370
         TabIndex        =   17
         Top             =   645
         Width           =   1275
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA INICIAL"
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
         Left            =   5100
         TabIndex        =   16
         Top             =   660
         Width           =   1485
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ALMACEN ORIGEN"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   990
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ALMACEN DESTINO"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   5625
         TabIndex        =   14
         Top             =   15
         Width           =   990
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CB."
         Height          =   360
         Left            =   -74850
         TabIndex        =   13
         Top             =   105
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO"
         Height          =   285
         Left            =   -71790
         TabIndex        =   12
         Top             =   105
         Width           =   840
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ver "
         Height          =   360
         Left            =   -67620
         TabIndex        =   11
         Top             =   120
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmFlexPtrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Modulo       :  frmFlexPtrans
' Fecha/Hora  : 06/12/2003 00:05
' Autor          : JCASTILLO
' Propósito    : Formulario de Consulta para mostrar las peticiones de
'                    transferencia junto con su estado y numero de mensajes
'                    asociados a cada petición de transferencia.
'---------------------------------------------------------------------------------------

Option Explicit

Dim first As Boolean

Public miosql As New clsSmartSQL
Public miRc As New ADODB.Recordset

Public selecciona_registro As Boolean

Dim miCod As MiCodBar
'Dim miref As String

Dim busca_por_codbar As Boolean

Private Sub cbCODALMORIG_GotFocus()

If SSTab1.Tab <> 0 Then SSTab1.Tab = 0

End Sub

Private Sub cbLista_click()

If (cbCODALMDEST.Text = cbCODALMORIG.Text) And (cbCODALMDEST.Text <> "") Then

 MsgBox "ORIGEN y DESTINO deben ser diferentes", vbInformation
 Exit Sub

End If


'restringir solo al almacen actual para TipoPermiso = 0 (dependiente comun)
'If TipoPermiso = 0 And cbCODALMORIG.Text <> AlmacenActual Then
    'MsgBox "Almacen de origen INCORRECTO", vbInformation
    'cbCODALMORIG.SetFocus
    'Exit Sub
'End If
  

With ioULTIMAS

If .Text <> "" Then
    Call carga_grid(.Text)
Else
    .Text = "150"
    MsgBox "Debe especificar un numero de Transferencias"
End If

End With

cbCODALMORIG.SetFocus

End Sub

Private Sub chameleonButton1_Click()
Dim linea1 As String
Dim linea2 As String
Dim tmpalm As String
         
   
   On Error GoTo chameleonButton1_Click_Error

       DoEvents

    If cbCODALMORIG.Text <> "" Then
        tmpalm = devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & cbCODALMORIG.Text, locCnn)
        If tmpalm = "@" Then tmpalm = ""
    End If
    
    linea1 = "Transferencias:  F.Inicial: " & ioFECHAINI.Text & ". F.Final: " & ioFECHAFIN.Text & ". Almacén: " & tmpalm
    linea2 = "Usuario: " & Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & UsuarioActual, locCnn)) & ". Fecha: " & Now
    
    Call PrintFlexGrid(fg, 1, 1, 2, linea1, linea2, 13, 2, 10)
    'fg.SaveGrid "c:\prueba.txt", flexFileCommaText
    'fg.PrintGrid "Transferencia " & rc.Fields("CODIGO"), , 2, 600, 600
    
    
    'actualizar los colores del grid, volivendolo a cargar
    'rc.Move 0

   On Error GoTo 0
   Exit Sub

chameleonButton1_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento chameleonButton1_Click de Formulario frmFlexPtrans"

End Sub

Private Sub cmCerrar_Click()

Unload Me

End Sub

Private Sub fg_KeyDown(KeyCode As Integer, Shift As Integer)

   On Error GoTo fg_KeyDown_Error

If KeyCode = 13 Then Call fg_dblClick

   On Error GoTo 0
   Exit Sub

fg_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento fg_KeyDown de Formulario frmFlexPtrans"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   On Error GoTo Form_KeyDown_Error

Select Case KeyCode

'Ir al grid, o regresar
Case vbKeyF5
    
    If fg.Rows > 1 Then
        If fg.TabStop Then
            fg.TabStop = False
            ioCODIGO.SetFocus
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
    Call cmCerrar_Click
    
Case vbKeyF4
    KeyCode = 0
    Call cbLista_click

End Select

   On Error GoTo 0
   Exit Sub

Form_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Form_KeyDown de Formulario frmFlexPtrans"

End Sub

Private Sub Form_Load()

With ioULTIMAS
    .Alineacion = 1
    .SoloNumeros = True
    .LongMaxima = 9
End With

With cbCODALMORIG
    .ConexionString = locCnn
    .LenCodigo = 3
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM ALMACENES WHERE MBAJA = 0 ORDER BY CODIGO"
    .CodigoWidth = 500
    .carga
End With

'If TipoPermiso = 0 Then
'    cbCODALMORIG.Text = AlmacenActual
'    cbCODALMORIG.Enabled = False
'End If

With cbCODALMDEST
    .ConexionString = locCnn
    .LenCodigo = 3
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM ALMACENES WHERE MBAJA = 0 ORDER BY CODIGO"
    .CodigoWidth = 500
    .carga
End With

With cbESTADO
    .añade_item "0  EN CREACION", 1
    .añade_item "1  PENDIENTE", 2
    .añade_item "2  ACEPTADA", 3
    .añade_item "3  CANCELADA", 4
    .LenCodigo = 1
    .CodigoWidth = 300
End With

With ioFECHAINI
    .dspFormat = "dd/mm/yyyy"
    .LongMaxima = 10
    .Alineacion = 1
End With

With ioFECHAFIN
    .dspFormat = "dd/mm/yyyy"
    .LongMaxima = 10
    .Alineacion = 1
End With
  
With ioCODIGO
    .SoloNumeros = True
    .Alineacion = 1
    .LongMaxima = 10
End With

With ioNUMPED
    .SoloNumeros = True
    .Alineacion = 1
    .LongMaxima = 10

End With
  
 ioULTIMAS.Text = "150"
 Call carga_grid(CLng(ioULTIMAS.Text))
 
 cbCODALMORIG.Text = AlmacenActual
 
End Sub



Private Sub ioCODBAR_GotFocus()

If SSTab1.Tab <> 1 Then SSTab1.Tab = 1

End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : ioCODBAR_Validate
' Fecha/Hora    : 29/01/2004 12:48
' Autor         : JCastillo
' Propósito     :  Validar codigo de Barras
'---------------------------------------------------------------------------------------
Private Sub ioCODBAR_Validate(Cancel As Boolean)
Dim t_articulo As Variant
 
   On Error GoTo ioCODBAR_Validate_Error

With ioCODBAR

If Trim(.Text) = "" Then Exit Sub

    'si es un codigo de barras con la longitud válidad
    If Len(Trim(.Text)) = LenCodBar Then
        
        
         'descomponer el codigo
        miCod = Descompone_CBAR(.Text)
        
        'comprobar si existe el artículo/temporada
        t_articulo = devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & miCod.CODIGO_ART & " AND TEMPOR = " & miCod.TEMPORADA_ART, locCnn)
        
        If t_articulo = "@" Then
        
                MsgBox "No existe el artículo para esa temporada!, Codigo de Barras no Válido", titulo, vbExclamation
                ioCODBAR.Text = ""
                .CancelarValidacion
                Cancel = True
                Exit Sub
                
        End If
        

        
             'codigo de artículo
        'ioCODART.Text = CLng(Left(.Text, 5))
        'temporada
        'cbTEMPOR.Text = CLng(Mid(.Text, 6, 3))
           
        'talla
        'cbCODTALLA.Text = CLng(Mid(.Text, 9, 2))
        'color
        'cbCODCOL.Text = CLng(Mid(.Text, 11, 3)
        'ioUNIDADES.Text = "1"
        
        'miref = Trim(T_articulo(0))
        
        
        busca_por_codbar = True
        Call cbLista_click

    Else

        MsgBox "Código de Barras no válido", vbExclamation, titulo
        ioCODBAR.Text = ""
        .CancelarValidacion
        Cancel = True
        busca_por_codbar = False
        Exit Sub

    End If
    
      '  DoEvents
       ' ioUNIDADES.SetFocus

End With

   On Error GoTo 0
   Exit Sub

ioCODBAR_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioCODBAR_Validate de Formulario frmFlexPtrans"

End Sub



Private Sub ioULTIMAS_Validate(Cancel As Boolean)

   On Error GoTo ioULTIMAS_Validate_Error

If ioULTIMAS.Text = "" Then ioULTIMAS.Text = "50"
    
cbCODALMORIG.SetFocus

DoEvents

carga_grid (CLng(ioULTIMAS.Text))




   On Error GoTo 0
   Exit Sub

ioULTIMAS_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioULTIMAS_Validate de Formulario frmPtrans"
    
End Sub

Private Sub fg_dblClick()
Dim tmpfila As Long
Dim tmpid As Long
Dim var As Integer

   On Error GoTo fg_dblClick_Error

    selecciona_registro = False
    
    'posicionarse en el registro
    If miRc.RecordCount > 0 Then
        
       frmPtrans.cargando_Grid = True
       
       miRc.MoveFirst
       
       ''ir al codigo seleccionado
       Do Until miRc.EOF
       
            If (miRc.fields("CODIGO") = CLng(fg.TextMatrix(fg.Row, 1))) And (miRc.fields("CODALMORIG") = CLng(Left(fg.TextMatrix(fg.Row, 3), 3))) Then
                   selecciona_registro = True
                   Exit Do
            End If
       
            miRc.MoveNext
       
       Loop
       
       frmPtrans.cargando_Grid = False
    
       
       'miRc.Find "CODIGO = " & fg.TextMatrix(fg.Row, 1) & " AND CODALMORIG = " & Left(fg.TextMatrix(fg.Row, 2), 3), , adSearchForward, 1
       'miRc.Filter = "CODIGO = " & fg.TextMatrix(fg.Row, 1) & " AND CODALMORIG = " & Left(fg.TextMatrix(fg.Row, 2), 3)
       
       DoEvents
       DoEvents
       
       'si busca por codigo de barras que se posicione en el artículo seleccionado
       If busca_por_codbar Then
          
          With frmPtrans
                        
           'devolver el id correspondiente
           tmpid = devuelve_campo("SELECT TOP 1 ID FROM DETTRANS WHERE CODIGO = " & fg.TextMatrix(fg.Row, 1) & " AND CODALM = " & Left(fg.TextMatrix(fg.Row, 2), 3) & " AND TEMPOR = " & miCod.TEMPORADA_ART & " AND CODART = " & miCod.CODIGO_ART & " AND CODTALLA = " & miCod.TALLA_ART & " AND CODCOL = " & miCod.COLOR_ART, locCnn)
           DoEvents
           
           'buscar el ID en el grid, para colorear la fila correspondiente
           tmpfila = .fg.FindRow(tmpid, , 1)
           .fg.Row = tmpfila
           
           'colorear de amarillo la fila entera
           For var = 0 To .fg.Cols - 1
            .fg.Col = var
            .fg.CellBackColor = vbYellow
           Next var
           
           .fg.Col = 3
           .fg.TopRow = tmpfila
           
            'seleccionar la fila ...
           '.fg.Select tmpfila, 1, tmpfila, .fg.Cols - 1
           
          End With
       
       End If
       
    End If
    
    Unload Me

   On Error GoTo 0
   Exit Sub

fg_dblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento fg_DblClick de Formulario frmFlexPtrans"
End Sub

Private Sub fg_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

Case 13, vbKeyEscape
    KeyAscii = 0
    Call fg_dblClick
    
    
End Select

End Sub

Private Sub Form_Activate()

    'Me.Refresh
    DoEvents
    If Not first Then
        first = True
        DoEvents
        fg.Visible = True
        fg.AutoSearch = flexSearchFromCursor
        fg.HighLight = flexHighlightWithFocus
    End If
    
    cbCODALMORIG.SetFocus
    
   
    
End Sub

Private Sub cbBorrar_click()

cbCODALMORIG.Text = AlmacenActual
cbCODALMDEST.Text = ""
cbESTADO.Text = ""
ioNUMPED.Text = ""
ioCODIGO.Text = ""
ioFECHAINI.Text = ""
ioFECHAFIN.Text = ""
ioCODBAR.Text = ""
ioULTIMAS.Text = "150"

If miRc.Filter <> "" Then miRc.Filter = ""

Call cbLista_click

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'si hemos establecido un filtro que no devuelve ningun registro,
    'borrar filtro para que no de error al volver al formulario
    If miRc.RecordCount = 0 Then Call cbBorrar_click
    
    'Set frmFlexPtrans = Nothing

End Sub


'---------------------------------------------------------------------------------------
' Subrutina     : carga_grid
' Fecha/Hora    : 24/11/2003 20:57
' Autor         : JCASTILLO
' Propósito     : cargar los primeros n registros de la tabla PTRANS
'---------------------------------------------------------------------------------------
Private Sub carga_grid(Numero_Registros As Long)
'Dim mirc As New ADODB.Recordset
Dim tmpsql As New clsSmartSQL
Dim tmplinea As Long
Dim tmpcodcolor As Long
Dim tmpprecom As Double

'1  codigo
'2  numped
'3  codalmorig
'4  codalmdest
'5  entrada
'6  estado
'7  fmodi

  
   On Error GoTo carga_grid_Error
   
   frmPtrans.cargando_Grid = True

If Numero_Registros = 0 Then Exit Sub
   
 miosql.ClearWhereClause
 tmpsql.AddTable "DETTRANS"
 
  Select Case TipoPermiso
  
  Case 0 'dependiente comun (restringir solo a las transferencias en las que vean las de estado 0 y 1 para su almacen (cuando sea su alm. el origen)
    
    miosql.AddComplexWhereClause "CODALMORIG = " & AlmacenActual & " AND (ESTADO = 0 OR ESTADO = 1)"
    'que vean las de estado 1 para su almacen como destino
    miosql.AddComplexWhereClause "CODALMDEST = " & AlmacenActual & " AND ESTADO = 1", LOGIC_OR
        
  Case 1 'supervisor
  
   ' miOsql.AddSimpleWhereClause "ESTADO", 0
   ' miOsql.AddSimpleWhereClause "ESTADO", 1, , , LOGIC_OR
  
  End Select

If ioCODBAR.Text <> "" Then

    tmpsql.AddField "cast(codigo AS char(15)) + cast(codalm AS char(4))"
    tmpsql.AddSimpleWhereClause "CODART", miCod.CODIGO_ART
    tmpsql.AddSimpleWhereClause "TEMPOR", miCod.TEMPORADA_ART
    tmpsql.AddSimpleWhereClause "CODTALLA", miCod.TALLA_ART
    tmpsql.AddSimpleWhereClause "CODCOL", miCod.COLOR_ART
    
    miosql.AddComplexWhereClause "cast(codigo AS char(15)) + cast(codalmorig AS char(4)) in (" & tmpsql.SQL & ")"
    
Else
    busca_por_codbar = False
End If

'para buscar por código y almacen origen de la transferencia ...
If Trim(ioCODIGO.Text) <> "" And (cbCODALMORIG.Text <> "") Then
    miosql.AddSimpleWhereClause "CODIGO", CLng(ioCODIGO.Text)
ElseIf Trim(ioCODIGO.Text) <> "" And (cbCODALMORIG.Text = "") Then
    MsgBox "Si quiere buscar por codigo, debe seleccionar almacen de origen", vbInformation, titulo
    cbCODALMORIG.SetFocus
End If

If ioNUMPED.Text <> "" Then
    miosql.AddSimpleWhereClause "NUMPED", CLng(ioNUMPED.Text)
End If

If cbCODALMORIG.Text <> "" Then
    miosql.AddSimpleWhereClause "CODALMORIG", CLng(cbCODALMORIG.Text)
End If

'mostrar solo las introducidas directamente desde almacén
'>=90000000
If ckDirectas.Value = True Then
    miosql.AddSimpleWhereClause "CODIGO", 90000000, , CLAUSE_GREATERTHANOREQUAL
End If

If cbCODALMDEST.Text <> "" Then
    miosql.AddSimpleWhereClause "CODALMDEST", CLng(cbCODALMDEST.Text)
End If

'filtrar por fechas
If (ioFECHAINI.Text <> "" And ioFECHAFIN.Text <> "") Then
    miosql.AddComplexWhereClause "FMODI >= '" & Format(Year((ioFECHAINI.Text)), "0000") & Format(Month((ioFECHAINI.Text)), "00") & Format(Day((ioFECHAINI.Text)), "00") & "' AND FMODI <= '" & Format(Year((ioFECHAFIN.Text)), "0000") & Format(Month((ioFECHAFIN.Text)), "00") & Format(Day((ioFECHAFIN.Text)), "00") & "'", LOGIC_AND
    'usa_where = True
End If

' estados de la transferencia
' 0 -> en creacion
' 1 -> pendientes
' 2 -> aceptadas
' 3 -> canceladas
If cbESTADO.Text <> "" Then
    
    miosql.AddSimpleWhereClause "Estado", CLng(cbESTADO.Text)

'para q solo salgan las canceladas si se piden
Else
    
    miosql.AddSimpleWhereClause "Estado", 3, , CLAUSE_DOESNOTEQUAL
    
End If



miRc.Close
miRc.Open "SELECT TOP " & Numero_Registros & " " & Right(miosql.SQL, Len(miosql.SQL) - 7), locCnn, adOpenStatic, adLockOptimistic

'miRc.Open "SELECT TOP " & Numero_Registros & " IDTRANS, codart, tempor, CODTALLA, CODCOL, codalmorig, codalmdest, entrada, unidades, estado, fmodi FROM PTRANS ORDER BY IDTRANS DESC", locCnn, adOpenDynamic, adLockReadOnly
    
    With fg
        
        .Redraw = flexRDNone
        .Clear
        .Cols = 10
        
        .ColFormat(6) = "Currency"
        .ColFormat(8) = "Currency"
        
        'poner títulos
        .TextMatrix(0, 1) = "CODIGO"    '
        .TextMatrix(0, 2) = "PEDIDO"
        .TextMatrix(0, 3) = "ORIGEN"    '
        .TextMatrix(0, 4) = "DESTINO"   '
       ' .TextMatrix(0, 4) = "ENTRADA"   '
        .TextMatrix(0, 5) = "ESTADO"   '
        .TextMatrix(0, 6) = "SUBTOT."
        .TextMatrix(0, 7) = "DCTO"
        .TextMatrix(0, 8) = "TOTAL"
        .TextMatrix(0, 9) = "FECHA"    '
      
        
        .Rows = 1
    
    Do Until miRc.EOF

        .Rows = .Rows + 1
        
        'numero de linea
        .TextMatrix(.Rows - 1, 1) = miRc.fields("CODIGO").Value
        
        'pedido
        .TextMatrix(.Rows - 1, 2) = miRc.fields("NUMPED").Value
                             
        If miRc.fields("CODALMORIG").Value > 0 Then _
        .TextMatrix(.Rows - 1, 3) = Format(miRc.fields("CODALMORIG").Value, "000") & " " & devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & miRc.fields("CODALMORIG").Value)
        
        If miRc.fields("CODALMDEST").Value > 0 Then _
        .TextMatrix(.Rows - 1, 4) = Format(miRc.fields("CODALMDEST").Value, "000") & " " & devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & miRc.fields("CODALMDEST").Value)
        
        Select Case miRc.fields("ESTADO").Value
        
        Case 0
            .TextMatrix(.Rows - 1, 5) = "EN CREACION"
        Case 1
            .TextMatrix(.Rows - 1, 5) = "PENDIENTE"
        Case 2
            .TextMatrix(.Rows - 1, 5) = "ACEPTADA"
        Case 3
            .TextMatrix(.Rows - 1, 5) = "CANCELADA"
        
        End Select
        
        'si busca por codigo de barras, sacar solo el importe correspondiente a esas filas ...
        If busca_por_codbar Then
        
            tmpprecom = devuelve_campo("SELECT PRECOM FROM MAARTIC WHERE CODIGO = " & miCod.CODIGO_ART & " AND TEMPOR = " & miCod.TEMPORADA_ART, locCnn)
            'tmpprecom = Obtiene_Precom_Pedido(CLng(miCod.CODIGO_ART), CLng(miCod.TEMPORADA_ART), CLng(miCod.TALLA_ART), CLng(miCod.COLOR_ART), locCnn)
            .TextMatrix(.Rows - 1, 6) = tmpprecom * devuelve_campo("SELECT SUM(UNIDADES) FROM DETTRANS WHERE CODIGO = " & miRc.fields("CODIGO").Value & " AND CODALM = " & miRc.fields("CODALMORIG").Value & " AND CODART = " & miCod.CODIGO_ART & " AND TEMPOR = " & miCod.TEMPORADA_ART & " AND CODTALLA = " & miCod.TALLA_ART & " AND CODCOL = " & miCod.COLOR_ART, locCnn)
        Else
            'importe
            .TextMatrix(.Rows - 1, 6) = miRc.fields("TOTAL").Value
        
        End If
        
        'dcto
        .TextMatrix(.Rows - 1, 7) = miRc.fields("DCTO").Value & "%"
        
        'total
        .TextMatrix(.Rows - 1, 8) = .TextMatrix(.Rows - 1, 6) - ((.TextMatrix(.Rows - 1, 6) * miRc.fields("DCTO").Value) / 100)
        
        .TextMatrix(.Rows - 1, 9) = miRc.fields("FMODI").Value
        
        
       miRc.MoveNext
        
    Loop
         
     If miRc.RecordCount > 0 Then
        .SubtotalPosition = flexSTAbove
        .subtotal flexSTCount, , 4, , vbBlue, vbWhite
        .subtotal flexSTSum, , 6, , vbBlue, vbWhite
        .subtotal flexSTSum, , 8, , vbBlue, vbWhite
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 3) = "Total: (" & .TextMatrix(1, 4) & ")" & " Transferencias"
        .TextMatrix(1, 4) = ""
        .Row = 1
        .Col = 3
        .CellAlignment = flexAlignRightCenter
    End If
    
    
    .AutoSize 1, .Cols - 1
    
    
    
    .Redraw = True
    
    
    '.Enabled = True
    End With
    
   
   frmPtrans.cargando_Grid = False

   On Error GoTo 0
   Exit Sub
   
carga_grid_Error:

   frmPtrans.cargando_Grid = False
   busca_por_codbar = False

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_grid de Formulario frmFlexPtrans"
 
End Sub



