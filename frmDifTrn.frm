VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmDifTrn 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diferencias con Pedido (artículos que FALTAN)"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9930
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
   ScaleHeight     =   6540
   ScaleWidth      =   9930
   Begin PCGestion.chameleonButton cbCancelar 
      Height          =   795
      Left            =   4905
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5715
      Width           =   930
      _extentx        =   1640
      _extenty        =   1402
      btype           =   9
      tx              =   "&Cancelar"
      enab            =   -1  'True
      font            =   "frmDifTrn.frx":0000
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmDifTrn.frx":002C
      picn            =   "frmDifTrn.frx":004A
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   5685
      Left            =   30
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   15
      Width           =   9885
      _cx             =   17436
      _cy             =   10028
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
      FormatString    =   $"frmDifTrn.frx":0926
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
   Begin PCGestion.chameleonButton cbImprimir 
      Height          =   795
      Left            =   4095
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5715
      Width           =   780
      _extentx        =   1376
      _extenty        =   1402
      btype           =   9
      tx              =   ""
      enab            =   -1  'True
      font            =   "frmDifTrn.frx":0A04
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmDifTrn.frx":0A30
      picn            =   "frmDifTrn.frx":0A4E
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   1
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
End
Attribute VB_Name = "frmDifTrn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NUMERO_PEDIDO As Long
Public CODIGO_TRANSF As Long

'---------------------------------------------------------------------------------------
' Procedimiento : cbImprimir_Click
' Fecha/Hora    : 04/02/2004 11:12
' Autor         : JCastillo
' Propósito     :  Imprimir la transferencia actual.
'---------------------------------------------------------------------------------------
'
Private Sub cbImprimir_Click()
Dim linea1 As String
Dim linea2 As String

    On Error GoTo cbImprimir_Click_Error
        
    DoEvents


    linea1 = "Diferencias con Pedido. Número de Pedido: " & NUMERO_PEDIDO & ". Número de Transferencia: " & CODIGO_TRANSF
    linea2 = "Usuario: " & Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & UsuarioActual, locCnn)) & ". Fecha: " & Now
    Call PrintFlexGrid(fg, 1, 1, 2, linea1, linea2, 13, 2)
    'fg.SaveGrid "c:\prueba.txt", flexFileCommaText
    'fg.PrintGrid "Transferencia " & rc.Fields("CODIGO"), , 2, 600, 600
    
    
    'actualizar los colores del grid, volivendolo a cargar
    'rc.Move 0

   On Error GoTo 0
   Exit Sub

cbImprimir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbImprimir_Click de Formulario frmPtrans"
End Sub


Private Sub cbCancelar_Click()

Unload Me

End Sub

'---------------------------------------------------------------------------------------
' Módulo     : frmDifTrn
' Fecha/Hora : 10/03/2004 12:29
' Autor      : JCastillo
' Propósito  : Mostrar las diferencias de las transferencias respecto al pedido
'                 es decir, los articulos que faltan, pues no deja meter de sobra.
'---------------------------------------------------------------------------------------
Private Sub Form_Load()

Move (Screen.Width - Width) \ 2, Separacion_MDIForm '(Screen.Height - Height) \ 2

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Set frmDifTrn = Nothing

End Sub
