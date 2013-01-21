VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmComPtrans 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comprobar Transferencia"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11910
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin PCGestion.miText ioCODBAR 
      Height          =   495
      Left            =   1275
      TabIndex        =   0
      Top             =   60
      Width           =   2595
      _extentx        =   4577
      _extenty        =   873
      font            =   "frmComPtrans.frx":0000
      dspformat       =   ""
      enabled         =   -1  'True
      espassword      =   -1  'True
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   6300
      Left            =   30
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   975
      Width           =   5610
      _cx             =   9895
      _cy             =   11112
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
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
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmComPtrans.frx":002C
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
   Begin VSFlex8Ctl.VSFlexGrid fgDif 
      Height          =   6300
      Left            =   5640
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   975
      Width           =   6270
      _cx             =   11060
      _cy             =   11112
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
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
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmComPtrans.frx":00CF
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
   Begin PCGestion.bsGradientLabel lblstatus 
      Height          =   315
      Left            =   45
      Top             =   7305
      Width           =   12405
      _extentx        =   10081
      _extenty        =   556
      caption         =   ""
      fount           =   "frmComPtrans.frx":0172
      captioncolour   =   0
      colour1         =   15640462
      colour2         =   7177785
      captionalignment=   1
   End
   Begin PCGestion.chameleonButton cmdAceptarTrans 
      Height          =   615
      Left            =   10245
      TabIndex        =   4
      Top             =   15
      Width           =   1650
      _extentx        =   2910
      _extenty        =   1085
      btype           =   3
      tx              =   "Introducir código de seguridad"
      enab            =   -1  'True
      font            =   "frmComPtrans.frx":01A0
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   16776960
      fcolo           =   16776960
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmComPtrans.frx":01CC
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel1 
      Height          =   315
      Left            =   3885
      Top             =   0
      Width           =   2685
      _extentx        =   4736
      _extenty        =   556
      caption         =   "- B - Borrar ultimo artículo"
      fount           =   "frmComPtrans.frx":01EA
      captioncolour   =   0
      colour1         =   16761024
      colour2         =   16744576
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel2 
      Height          =   315
      Left            =   30
      Top             =   645
      Width           =   5610
      _extentx        =   9895
      _extenty        =   556
      caption         =   "Introducido"
      fount           =   "frmComPtrans.frx":0218
      captioncolour   =   0
      colour1         =   16761024
      colour2         =   16744576
      captionalignment=   1
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel3 
      Height          =   315
      Left            =   3885
      Top             =   315
      Width           =   2685
      _extentx        =   4736
      _extenty        =   556
      caption         =   "- A - Aceptar Transferencia"
      fount           =   "frmComPtrans.frx":0246
      captioncolour   =   0
      colour1         =   16761024
      colour2         =   16744576
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel4 
      Height          =   330
      Left            =   5670
      Top             =   645
      Width           =   6225
      _extentx        =   10980
      _extenty        =   582
      caption         =   "Diferencias"
      fount           =   "frmComPtrans.frx":0274
      captioncolour   =   0
      colour1         =   16761024
      colour2         =   16744576
      captionalignment=   1
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel5 
      Height          =   315
      Left            =   6600
      Top             =   315
      Width           =   1050
      _extentx        =   1852
      _extenty        =   556
      caption         =   "- S - Salir"
      fount           =   "frmComPtrans.frx":02A2
      captioncolour   =   0
      colour1         =   16761024
      colour2         =   16744576
   End
   Begin PCGestion.bsGradientLabel lblSobran 
      Height          =   315
      Left            =   8490
      Top             =   -15
      Width           =   1725
      _extentx        =   3043
      _extenty        =   556
      caption         =   ""
      fount           =   "frmComPtrans.frx":02D0
      captioncolour   =   0
      colour1         =   11311500
      colour2         =   16558731
   End
   Begin PCGestion.bsGradientLabel lblFaltan 
      Height          =   315
      Left            =   8490
      Top             =   315
      Width           =   1725
      _extentx        =   3043
      _extenty        =   556
      caption         =   ""
      fount           =   "frmComPtrans.frx":02FE
      captioncolour   =   0
      colour1         =   11311500
      colour2         =   16558731
   End
   Begin PCGestion.chameleonButton cbImprimir 
      Height          =   615
      Left            =   7680
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   15
      Width           =   780
      _extentx        =   1376
      _extenty        =   1085
      btype           =   9
      tx              =   ""
      enab            =   -1  'True
      font            =   "frmComPtrans.frx":032C
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmComPtrans.frx":0358
      picn            =   "frmComPtrans.frx":0376
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   1
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin PCGestion.chameleonButton cbRepetir 
      Height          =   300
      Left            =   6570
      TabIndex        =   6
      Top             =   15
      Width           =   1095
      _extentx        =   1931
      _extenty        =   529
      btype           =   3
      tx              =   "Repetir"
      enab            =   -1  'True
      font            =   "frmComPtrans.frx":1052
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   11513775
      bcolo           =   11513775
      fcol            =   16776960
      fcolo           =   16776960
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmComPtrans.frx":107E
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO DE BARRAS"
      Height          =   585
      Left            =   -15
      TabIndex        =   1
      Top             =   15
      Width           =   1245
   End
End
Attribute VB_Name = "frmComPtrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Modulo      : frmComPtrans
' Fecha/Hora  : 11/01/2004 17:28
' Autor       : JCASTILLO
' Propósito   : Comprobar transferencia. Al aceptar la transferencia se debe comprobar
'               con la mercancía. Para ello se introducen todos los codigos de barras de la
'               mercancía
'---------------------------------------------------------------------------------------
Option Explicit

'para localizar la transferencia
Public ID_TRANSF As Long
Public ALM_TRANSF As Byte

Dim tmpCnn As New ADODB.Connection
Dim rc As New ADODB.Recordset
Dim rcDif As New ADODB.Recordset
Dim rcRec As New ADODB.Recordset

'para controlar que no se compruebe mas de una vez para la misma transferencia
Dim Comprobada As Boolean

'si ultimo_ok = true es que el ultimo codigo pertenece correctamente a la
'transferencia, Para cuando el usuario pulse B, saber en que tabla se ha
'Estados de la comprobación
'0 -> no se ha pasado
'1 -> correcta, todos los articulos OK.
'2 -> incorrecta (faltan o sobran)
Public estado As Byte

Dim filas As Long
Dim filasdif As Long

Const path_tmp = "C:\TRANSFERENCIAS\COMPTRN\"
Dim fichero As String
'Const fichero = path_tmp & "C:\TmpCmpPtrn.mdb"

Dim sCnn As String

Dim repetir As Integer




'---------------------------------------------------------------------------------------
' Procedimiento : cbImprimir_Click
' Fecha/Hora    : 04/02/2004 11:12
' Autor         : JCastillo
' Propósito     :  Imprimir las diferencias con la transferencia actual.
'---------------------------------------------------------------------------------------
'
Private Sub cbImprimir_Click()
Dim linea1 As String
Dim linea2 As String
        
   On Error GoTo cbImprimir_Click_Error

    DoEvents

    linea1 = "Diferencias de con Transferencia. Codigo: " & Me.ID_TRANSF & ". Origen: " & Trim(devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & Me.ALM_TRANSF, locCnn)) & ". Destino: " & Trim(devuelve_campo("select descripcion from almacenes where codigo = " & AlmacenActual, locCnn))
    linea2 = lblFaltan.Caption & " . " & lblSobran.Caption & ". Usuario: " & Trim(devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & UsuarioActual, locCnn)) & ". Fecha: " & Now
    Call PrintFlexGrid(fgDif, 1, 1, 1, linea1, linea2, 10, 2)

   On Error GoTo 0
   Exit Sub

cbImprimir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbImprimir_Click de Formulario frmComPtrans"

End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : comprobar_codigo_trn
' Fecha/Hora  : 11/01/2004 17:31
' Autor       : JCASTILLO
' Propósito   : Comprueba que el codigo exista en la transferencia
'               TRUE -> CORRECTO
'               FALSE-> INCORRECTO
'---------------------------------------------------------------------------------------
Private Function comprobar_codigo_trn(codigo As String) As Boolean
Dim Codigo_B As MiCodBar
Dim tmpcodcolor As Variant
Dim tmpcodprov As Variant
Dim i As Integer

 On Error GoTo comprobar_codigo_trn_Error
 
  Codigo_B = Descompone_CBAR(codigo)
    
  'primero comprobar que exista el artículo en la base de datos ...
  If devuelve_campo("Select CODIGO FROM MAARTIC WHERE CODIGO = " & Codigo_B.CODIGO_ART & " AND TEMPOR = " & Codigo_B.TEMPORADA_ART, locCnn) = "@" Then
    lblstatus.Caption = "Código no válido. No existe en la base de datos."
    comprobar_codigo_trn = False
    Exit Function
  End If
   
  'luego si existe el artículo en esa transferencia ...
  If devuelve_campo("Select CODART from tmp_DETTRANS where CODART = " & CLng(Codigo_B.CODIGO_ART) & _
    " AND TEMPOR = " & CLng(Codigo_B.TEMPORADA_ART) & " AND CODTALLA = " & CLng(Codigo_B.TALLA_ART) & " AND CODCOL = " & CLng(Codigo_B.COLOR_ART), tmpCnn) <> "@" Then
    
     
    'articulo no existe en la transferencia.
    tmpcodprov = devuelve_campo("SELECT CODPROV FROM MAARTIC WHERE CODIGO = " & Codigo_B.CODIGO_ART & " AND TEMPOR = " & Codigo_B.TEMPORADA_ART, locCnn)
    
    If repetir <= 0 Then
    
    'articulo OK, existe.
    fg.AddItem "", 2
    fg.TextMatrix(2, 1) = Trim(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & tmpcodprov, locCnn))
    fg.TextMatrix(2, 2) = Trim(devuelve_campo("SELECT REF FROM MAARTIC WHERE CODIGO = " & Codigo_B.CODIGO_ART & " AND TEMPOR = " & Codigo_B.TEMPORADA_ART, locCnn))
        
    fg.TextMatrix(2, 3) = Codigo_B.CODIGO_ART & " " & devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & Codigo_B.CODIGO_ART & " AND TEMPOR = " & Codigo_B.TEMPORADA_ART, locCnn)
    fg.TextMatrix(2, 4) = Codigo_B.TEMPORADA_ART & " " & devuelve_campo("SELECT ABREVIA FROM TEMPOR WHERE IDTEM = " & Codigo_B.TEMPORADA_ART, locCnn)
    fg.TextMatrix(2, 5) = devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & Codigo_B.TALLA_ART, locCnn)
    fg.TextMatrix(2, 6) = devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & Codigo_B.COLOR_ART, locCnn)
    tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & Codigo_B.COLOR_ART, locCnn)
    fg.Row = 2
    fg.Col = 6
    fg.CellBackColor = tmpcodcolor
    fg.Col = 3
    fg.AutoSize 1, fg.Cols - 1
    
    fg.subtotal flexSTCount, , 6, , vbBlue, vbWhite
    fg.TextMatrix(1, 5) = "Total"
    fg.TextMatrix(1, 1) = ""
    
   ' filas = filas + 1
    
    'añadir registro de recibido
    
    
    
    With rcRec
        .AddNew
        .fields("CODART") = CDbl(Codigo_B.CODIGO_ART)
        .fields("CODCOL") = CLng(Codigo_B.COLOR_ART)
        .fields("CODTALLA") = CLng(Codigo_B.TALLA_ART)
        .fields("TEMPOR") = CLng(Codigo_B.TEMPORADA_ART)
        .Update
    End With
    
    
    Else
    
        For i = 1 To repetir
        
        
        'articulo OK, existe.
    fg.AddItem "", 2
    fg.TextMatrix(2, 1) = Trim(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & tmpcodprov, locCnn))
    fg.TextMatrix(2, 2) = Trim(devuelve_campo("SELECT REF FROM MAARTIC WHERE CODIGO = " & Codigo_B.CODIGO_ART & " AND TEMPOR = " & Codigo_B.TEMPORADA_ART, locCnn))
        
    fg.TextMatrix(2, 3) = Codigo_B.CODIGO_ART & " " & devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & Codigo_B.CODIGO_ART & " AND TEMPOR = " & Codigo_B.TEMPORADA_ART, locCnn)
    fg.TextMatrix(2, 4) = Codigo_B.TEMPORADA_ART & " " & devuelve_campo("SELECT ABREVIA FROM TEMPOR WHERE IDTEM = " & Codigo_B.TEMPORADA_ART, locCnn)
    fg.TextMatrix(2, 5) = devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & Codigo_B.TALLA_ART, locCnn)
    fg.TextMatrix(2, 6) = devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & Codigo_B.COLOR_ART, locCnn)
    tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & Codigo_B.COLOR_ART, locCnn)
    fg.Row = 2
    fg.Col = 6
    fg.CellBackColor = tmpcodcolor
    fg.Col = 3
    fg.AutoSize 1, fg.Cols - 1
    
    fg.subtotal flexSTCount, , 6, , vbBlue, vbWhite
    fg.TextMatrix(1, 5) = "Total"
    fg.TextMatrix(1, 1) = ""
        
        
        
        With rcRec
            .AddNew
            .fields("CODART") = CDbl(Codigo_B.CODIGO_ART)
            .fields("CODCOL") = CLng(Codigo_B.COLOR_ART)
            .fields("CODTALLA") = CLng(Codigo_B.TALLA_ART)
            .fields("TEMPOR") = CLng(Codigo_B.TEMPORADA_ART)
            .Update
        End With
    
        Next i
        
        repetir = 0
    
    End If
    
    DoEvents
    
    
    
    'fg.Rows = fg.Rows + 1
    comprobar_codigo_trn = True
       
        
    
  Else
  
      'añade registro de diferencias y en el grid
      Call añade_diferencia(CLng(Codigo_B.CODIGO_ART), CByte(Codigo_B.TEMPORADA_ART), CInt(Codigo_B.TALLA_ART), CInt(Codigo_B.COLOR_ART), 1, 1)

    comprobar_codigo_trn = False
    
  End If

  Exit Function

comprobar_codigo_trn_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento comprobar_codigo_trn de Formulario frmComPtrans"
    
End Function


Private Sub añade_diferencia(codart As Long, tempor As Byte, codtalla As Integer, codcol As Integer, uds As Single, causa As Byte)
Dim tmpcodcolor As Variant
Dim tmpcodprov As Variant

   On Error GoTo añade_diferencia_Error

    'articulo no existe en la transferencia.
    tmpcodprov = devuelve_campo("SELECT CODPROV FROM MAARTIC WHERE CODIGO = " & codart & " AND TEMPOR = " & tempor, locCnn)
    
    With fgDif
        .AddItem "", 2
        .TextMatrix(2, 1) = Trim(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & tmpcodprov, locCnn))
        .TextMatrix(2, 2) = Trim(devuelve_campo("SELECT REF FROM MAARTIC WHERE CODIGO = " & codart & " AND TEMPOR = " & tempor, locCnn))
        .TextMatrix(2, 3) = Format(codart, "00000") & " " & devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & codart & " AND TEMPOR = " & tempor, locCnn)
        .TextMatrix(2, 4) = Format(tempor, "000") & " " & devuelve_campo("SELECT ABREVIA FROM TEMPOR WHERE IDTEM = " & tempor, locCnn)
        .TextMatrix(2, 5) = devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & codtalla, locCnn)
        .TextMatrix(2, 6) = devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & codcol, locCnn)
        .TextMatrix(2, 7) = uds
        
        If causa = 1 Then
            .TextMatrix(2, 8) = "SOBRA"
        Else
            .TextMatrix(2, 8) = "FALTA"
        End If
        
       tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & codcol, locCnn)
    
    If tmpcodcolor <> "@" Then
        .Row = 2
        .Col = 6
        .CellBackColor = tmpcodcolor
        .Col = 3
    End If

    '.SubtotalPosition = flexSTBelow
        .subtotal flexSTCount, , 5, , vbBlue, vbWhite
        .subtotal flexSTSum, , 7, , vbBlue, vbWhite
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 4) = "Total"
        .TextMatrix(1, 6) = "Uds:"
        .AutoSize 1, .Cols - 1
    
    End With
        
    With rcDif
        .AddNew
        .fields("CODART") = codart
        .fields("CODCOL") = codcol
        .fields("CODTALLA") = codtalla
        .fields("TEMPOR") = tempor
        .fields("UNIDADES") = uds
        .fields("CAUSA") = causa
        .Update
    End With

   On Error GoTo 0
   Exit Sub

añade_diferencia_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento añade_diferencia de Formulario frmComPtrans"

End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : cbRepetir_Click
' Fecha/Hora  : 13/10/2004 19:10
' Autor       : JCASTILLO
' Propósito   :
'
'---------------------------------------------------------------------------------------
Private Sub cbRepetir_Click()
Dim srepetir As String

   On Error GoTo cbRepetir_Click_Error

    Do
        srepetir = InputBox("Repetir la siguiente prenda Nº veces", "Número de veces a repetir", 1)
    Loop Until Trim(srepetir) <> "" And IsNumeric(Trim(srepetir))
    
    repetir = CLng(srepetir)
    

   On Error GoTo 0
   Exit Sub

cbRepetir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbRepetir_Click de Formulario frmComPtrans"
End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : cmdAceptarTrans_Click
' Fecha/Hora  : 14/01/2004 21:20
' Autor       : JCASTILLO
' Propósito   : Intenta aceptar transferencia desde el codigo de seguridad
'---------------------------------------------------------------------------------------
Private Sub cmdAceptarTrans_Click()
Dim m As Double
Dim codusr As String
Dim tmpid As Variant
Dim rcmsg As New ADODB.Recordset

   'On Error GoTo cmdAceptarTrans_Click_Error
   
codusr = InputBox("Introduzca Código de Seguridad para la Transferencia", "Introduzca Codigo")

If codusr = "" Or Not IsNumeric(codusr) Then Exit Sub

m = CodigoSeguridad_TRN(Format(ID_TRANSF, "000000000") & Format(ALM_TRANSF, "000"))

If codusr <> m Then
    MsgBox "¡Código de Seguridad No Válido!", vbExclamation, titulo
    'estado = 2
Else
        
    'que escriba todas las diferencias al fichero si no lo ha echo ya
    If Not Comprobada Then Call comprueba_unidades
    
    tmpid = devuelve_campo("SELECT MAX(ID) + 1 FROM PTRANSMSG WHERE CODALM = " & AlmacenActual, locCnn)
    
    If tmpid = "@" Then tmpid = 1
    
    With rcmsg
        
        .Open "SELECT TOP 1 * FROM PTRANSMSG", locCnn, adOpenDynamic, adLockOptimistic
        
        .AddNew
        .fields("ID") = tmpid
        .fields("CODIGO") = ID_TRANSF
        .fields("CODALMORIG") = ALM_TRANSF
        .fields("CODUSR") = UsuarioActual
        .fields("CODALM") = AlmacenActual
        .fields("MSG") = Generar_Informe
        .Update
                
        .Close
                
    End With
    
    'poner a true para q no deje meter mas artículos despues de validar
    Set rcmsg = Nothing
    
    Comprobada = True
    MsgBox "El código se ha validado correctamente. Transferencia Aceptada", vbInformation, titulo
    estado = 1  'poner a correcta
    Unload Me   'descargar el formulario para continuar el proceso
    
End If

   On Error GoTo 0
   Exit Sub

cmdAceptarTrans_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cmdAceptarTrans_Click de Formulario frmComPtrans"

End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : Generar_Informe
' Fecha/Hora  : 14/01/2004 21:32
' Autor       : JCASTILLO
' Propósito   : Genera un informe con las diferencias que sera insertado en un mensaje
'               relacionado con esta transferencia
'---------------------------------------------------------------------------------------
Private Function Generar_Informe() As String
Dim tmpinf As String
Dim tmpcabe As String

   On Error GoTo Generar_Informe_Error

    With rcDif
    'si no hay nada, salir
   ' If .RecordCount <= 0 Then
   '     Generar_Informe = ""
   '     Exit Function
   ' End If
    
    .MoveFirst
   ' .Sort = "CAUSA"
    
   Do
    
            tmpinf = tmpinf & "CB: [" & Format(.fields("CODART"), "00000") & Format(.fields("TEMPOR"), "000") & Format(.fields("CODTALLA"), "00") & Format(.fields("CODCOL"), "000") & "] "
            tmpinf = tmpinf & Format(.fields("CODART"), "00000") & " " & devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & .fields("CODART") & " AND TEMPOR = " & .fields("TEMPOR"), locCnn)
            tmpinf = tmpinf & "  " & Format(.fields("TEMPOR"), "000") & " " & devuelve_campo("SELECT ABREVIA FROM TEMPOR WHERE IDTEM = " & .fields("TEMPOR"), locCnn)
            tmpinf = tmpinf & "-" & devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & .fields("CODTALLA"), locCnn)
            tmpinf = tmpinf & "-" & devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & .fields("CODCOL"), locCnn)
            tmpinf = tmpinf & "-Unidades: " & .fields("UNIDADES") & "   "
            
            
        Select Case .fields("CAUSA")
        
        Case 1 'no existe en la transferencia (sobra)
        
           tmpinf = tmpinf & "  CAUSA: SOBRA"
        Case 2 'falta articulo
        
           tmpinf = tmpinf & "  CAUSA: FALTA"
                
        End Select
        
           'nueva linea
           tmpinf = tmpinf & vbCrLf
    
       If Not .EOF Then .MoveNext
    
   Loop Until .EOF
    
   End With
      
   tmpcabe = Format(Now, "dd/mm/yyyy") & " INFORME DE DIFERENCIAS: " & vbCrLf & vbCrLf & "Total Sobran: " & devuelve_campo("SELECT COUNT(CODART) from DIFERENCIAS where CAUSA = 1", tmpCnn) & vbCrLf & "Total Faltan: " & devuelve_campo("SELECT COUNT(CODART) from DIFERENCIAS where CAUSA = 2", tmpCnn)
   Generar_Informe = tmpcabe & vbCrLf & tmpinf
   
   tmpinf = ""

   On Error GoTo 0
   Exit Function

Generar_Informe_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Generar_Informe de Formulario frmComPtrans"
End Function



'Private Sub cbAceptar_Click()
'Unload Me
'End Sub

'Private Sub cbCancelar_Click()

'Unload Me
'Set frmComPtrans = Nothing

'End Sub

Private Sub Form_Load()

    With fg
    .Cols = 7
    .Rows = 2
    .TextMatrix(filas, 1) = "PROV."
    .TextMatrix(filas, 2) = "REF."
    .TextMatrix(filas, 3) = "MODELO"
    .TextMatrix(filas, 4) = "TEMP."
    .TextMatrix(filas, 5) = "TALLA"
    .TextMatrix(filas, 6) = "COLOR"
    .subtotal flexSTCount, , 6, , vbBlue, vbWhite
    .TextMatrix(1, 5) = "Total"
    .TextMatrix(1, 1) = ""
    .AutoSize 1, .Cols - 1
    
    End With
    
    With fgDif
    .Cols = 9  'una columna mas, el motivo de la diferencia
    .Rows = 2
    
    .TextMatrix(filas, 1) = "PROV."
    .TextMatrix(filas, 2) = "REF."
    .TextMatrix(filas, 3) = "MODELO"
    .TextMatrix(filas, 4) = "TEMP."
    .TextMatrix(filas, 5) = "TALLA"
    .TextMatrix(filas, 6) = "COLOR"
    .TextMatrix(filas, 7) = "UDS."
    .TextMatrix(filas, 8) = "CAUSA"
    .subtotal flexSTCount, , 4, , vbBlue, vbWhite
    .TextMatrix(1, 3) = "Total"
    .TextMatrix(1, 1) = ""
    .AutoSize 1, .Cols - 1
    End With

    'filas = 1
    'filasdif = 1
    'fg.Rows = fg.Rows + 1
    
    With ioCODBAR
        .LongMaxima = LenCodBar
        .SoloNumeros = True
        .PermitirBlanco = True
    End With
    
    fichero = "ComTRN-" & ID_TRANSF & "-" & ALM_TRANSF & ".ctran"
        
    sCnn = strCnnMdb & path_tmp & fichero
    'sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Engine Type=5;Data Source=" & fichero
    
    'crea la db temporal para guardar los datos
    Call CreateDatabaseTemporal
    
    
    
    
    'carga transferencia al fichero de trabajo
    Call cargar_transferencia

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'descargar objetos
If rc.State = 1 Then
    rc.Close
    Set rc = Nothing
End If

If rcDif.State = 1 Then
    rcDif.Close
    Set rcDif = Nothing
End If

If rcRec.State = 1 Then
    rcRec.Close
    Set rcRec = Nothing
End If

If tmpCnn.State = 1 Then
    tmpCnn.Close
    Set tmpCnn = Nothing
End If

End Sub

Private Sub ioCODBAR_LostFocus()

   On Error GoTo ioCODBAR_LostFocus_Error

With ioCODBAR

'si ya ha sido comprobada no dejar meter mas prendas
If Comprobada Then
    .SetFocus
    lblstatus.Caption = "La transferencia ya ha sido comprobada"
    Exit Sub
End If

If Len(.Text) <> 13 Then

    .SetFocus
    .CancelarValidacion
    lblstatus.Caption = "Código No válido"
    Beep
    Call Espera(1)
    Beep
    Call Espera(1)
    Beep

Else
       
    If comprobar_codigo_trn(.Text) Then
        .Text = ""
        .SetFocus 'si es correcta, simplemente posicionar otra vez
       
    Else
        .Text = ""
        .SetFocus 'si es incorrecta, activar cancelar valdación
        .CancelarValidacion
        lblstatus.Caption = "No existe en transferencia"
        
         Beep
         Call Espera(1)
         Beep
         Call Espera(1)
         Beep
        
    End If
    
End If

End With

   On Error GoTo 0
   Exit Sub

ioCODBAR_LostFocus_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioCODBAR_LostFocus de Formulario frmComPtrans"

End Sub



'---------------------------------------------------------------------------------------
' Subrutina   : comprueba_unidades
' Fecha/Hora  : 13/01/2004 22:11
' Autor       : JCASTILLO
' Propósito   : Comprueba si concuerdan las unidades del articulo, articulo por articulo
'               buscando en tmp_DETTRANS
'---------------------------------------------------------------------------------------
Private Sub comprueba_unidades()
Dim rcbus As New ADODB.Recordset
Dim rctrn As New ADODB.Recordset
Dim tmpcodcolor As Variant

   On Error GoTo comprueba_unidades_Error
    
    'sacar totales
    rcbus.Open "SELECT Recibidos.CODART, Recibidos.CODCOL, Recibidos.CODTALLA, Recibidos.TEMPOR, Count([codart]) AS UNIDADES " & _
               "From Recibidos GROUP BY Recibidos.CODART, Recibidos.CODCOL, Recibidos.CODTALLA, Recibidos.TEMPOR;", tmpCnn, adOpenStatic, adLockReadOnly
    
    
   ' rctrn.Open "SELECT tmp_DETTRANS.CODART, tmp_DETTRANS.CODCOL, tmp_DETTRANS.CODTALLA, tmp_DETTRANS.TEMPOR, tmp_DETTRANS.UNIDADES " & _
   '            "From tmp_DETTRANS GROUP BY tmp_DETTRANS.CODART, tmp_DETTRANS.CODCOL, tmp_DETTRANS.CODTALLA, tmp_DETTRANS.TEMPOR, tmp_DETTRANS.UNIDADES;", tmpCnn, adOpenDynamic, adLockReadOnly

    
    
    Do Until rcbus.EOF
        'buscar ...
        
        If rctrn.State = 1 Then rctrn.Close
        
        rctrn.Open "SELECT tmp_DETTRANS.CODART, tmp_DETTRANS.CODCOL, tmp_DETTRANS.CODTALLA, tmp_DETTRANS.TEMPOR, sum(tmp_DETTRANS.UNIDADES) as UNIDADES " & _
               "From tmp_DETTRANS WHERE CODART =" & rcbus.fields("CODART") & " AND TEMPOR = " & rcbus.fields("TEMPOR") & " AND CODTALLA = " & rcbus.fields("CODTALLA") & " AND CODCOL = " & rcbus.fields("CODCOL") & " GROUP BY tmp_DETTRANS.CODART, tmp_DETTRANS.CODCOL, tmp_DETTRANS.CODTALLA, tmp_DETTRANS.TEMPOR", tmpCnn, adOpenStatic, adLockReadOnly

                 
        'no se encuentran registros, añadir uno en diferencias, la
        'mercancia no existe  (SOBRA EN LA TRANSFERENCIA)...
        If (rctrn.EOF) Or (rctrn.RecordCount <= 0) Then
                     
            Call añade_diferencia(rcbus.fields("CODART"), rcbus.fields("TEMPOR"), rcbus.fields("CODTALLA"), rcbus.fields("CODCOL"), 1, 1)
            
        'si se encuentra registros, comprobar que las unidades coinciden ...
        'de lo contrario añadir un registro en Diferencias, con la diferencia
        'de unidades
        Else
                
              Select Case rcbus.fields("UNIDADES")
              
              'si es igual OK
              Case Is = rctrn.fields("UNIDADES")
              
              'si es mayor (han mandado de mas)
              Case Is > rctrn.fields("UNIDADES")
                             
                    Call añade_diferencia(rcbus.fields("CODART"), rcbus.fields("TEMPOR"), rcbus.fields("CODTALLA"), rcbus.fields("CODCOL"), rcbus.fields("UNIDADES") - rctrn.fields("UNIDADES"), 1)
             
              End Select
        
        End If
    
    If Not rcbus.EOF Then rcbus.MoveNext
    Loop
                  
    '----------------------------------------------------------------------------------------------------------------
    
    'Ahora comprobar que no FALTEN articulos (que se hayan pasado todos por el lector, o
    'que no se haya perdido ninguno. (proceso inverso al anterior)
    
       If rctrn.State = 1 Then rctrn.Close
       rctrn.Open "SELECT tmp_DETTRANS.CODART, tmp_DETTRANS.CODCOL, tmp_DETTRANS.CODTALLA, tmp_DETTRANS.TEMPOR, sum(tmp_DETTRANS.UNIDADES) as UNIDADES " & _
               "From tmp_DETTRANS GROUP BY tmp_DETTRANS.CODART, tmp_DETTRANS.CODCOL, tmp_DETTRANS.CODTALLA, tmp_DETTRANS.TEMPOR", tmpCnn, adOpenStatic, adLockReadOnly

      Do Until rctrn.EOF
      
       If rcbus.State = 1 Then rcbus.Close
       rcbus.Open "SELECT Recibidos.CODART, Recibidos.CODCOL, Recibidos.CODTALLA, Recibidos.TEMPOR, Count([codart]) AS UNIDADES " & _
               "From Recibidos  WHERE CODART =" & rctrn.fields("CODART") & " AND TEMPOR = " & rctrn.fields("TEMPOR") & " AND CODTALLA = " & rctrn.fields("CODTALLA") & " AND CODCOL = " & rctrn.fields("CODCOL") & " GROUP BY Recibidos.CODART, Recibidos.CODCOL, Recibidos.CODTALLA, Recibidos.TEMPOR;", tmpCnn, adOpenStatic, adLockReadOnly
              
        If rcbus.EOF Or (rcbus.RecordCount <= 0) Then
                     
            Call añade_diferencia(rctrn.fields("CODART"), rctrn.fields("TEMPOR"), rctrn.fields("CODTALLA"), rctrn.fields("CODCOL"), rctrn.fields("UNIDADES"), 2)
                     
        Else
        
        
            Select Case rcbus.fields("UNIDADES")
            
            'si es igual, esta OK, no hacer nada
            Case Is = rctrn.fields("UNIDADES")
            
            'faltan unidades
            Case Is < rctrn.fields("UNIDADES")
        
                   Call añade_diferencia(rctrn.fields("CODART"), rctrn.fields("TEMPOR"), rctrn.fields("CODTALLA"), rctrn.fields("CODCOL"), rctrn.fields("UNIDADES") - rcbus.fields("UNIDADES"), 2)
                   
      
            End Select
      
        
        End If
            
      
      If Not rctrn.EOF Then rctrn.MoveNext
      Loop
      
    '----------------------------------------------------------------------------------------------------------------

   rcbus.Close
   Set rcbus = Nothing
   rctrn.Close
   Set rctrn = Nothing
   On Error GoTo 0
   Exit Sub

comprueba_unidades_Error:

    Set rcbus = Nothing
    Set rctrn = Nothing
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento comprueba_unidades de Formulario frmComPtrans"
End Sub


'---------------------------------------------------------------------------------------
' Subrutina   : CreateDatabaseTemporal
' Fecha/Hora  : 13/01/2004 21:03
' Autor       : JCASTILLO
' Propósito   : Crea la base de datos temporal para trabajar con datos temporales
'---------------------------------------------------------------------------------------
Private Sub CreateDatabaseTemporal()

Dim Cat     As New ADOX.Catalog
Dim Tbl(8) As ADOX.Table
Dim Idx()   As ADOX.Index
Dim msgErrR As Integer

   On Error GoTo CreateDatabaseTemporal_Error

If Dir(path_tmp & fichero) = "" Then

DoEvents

Cat.Create sCnn

  '----------* Table Definition of Diferencias *----------
  Set Tbl(0) = New ADOX.Table
  Tbl(0).ParentCatalog = Cat
  With Tbl(0)
    .Name = "Diferencias"
    .Columns.Append "CAUSA", adUnsignedTinyInt
    .Columns("CAUSA").Properties("Description").Value = "1=NO CONSTA EN TRNSF, 2=FALTA DE LA TRNSF"
    .Columns("CAUSA").Properties("Default").Value = "0"
    .Columns.Append "CODART", adSmallInt
    .Columns("CODART").Properties("Nullable").Value = False
    .Columns.Append "TEMPOR", adUnsignedTinyInt
    .Columns("TEMPOR").Properties("Nullable").Value = False
    .Columns.Append "CODCOL", adSmallInt
    .Columns("CODCOL").Properties("Nullable").Value = False
    .Columns.Append "CODTALLA", adSmallInt
    .Columns("CODTALLA").Properties("Nullable").Value = False
    .Columns.Append "UNIDADES", adSingle
    .Columns("UNIDADES").Properties("Nullable").Value = False
   ' .Columns.Append "TEMPOR", adUnsignedTinyInt
   ' .Columns("TEMPOR").Properties("Nullable").Value = False
  End With
  Cat.tables.Append Tbl(0)
  
  '----------* Table Definition of Recibidos *----------
  Set Tbl(7) = New ADOX.Table
  Tbl(7).ParentCatalog = Cat
  With Tbl(7)
    .Name = "Recibidos"
    .Columns.Append "ID", adInteger
    .Columns("ID").Properties("AutoIncrement").Value = True
    .Columns.Append "CODART", adSmallInt
    .Columns("CODART").Properties("Nullable").Value = False
    .Columns.Append "CODCOL", adSmallInt
    .Columns("CODCOL").Properties("Nullable").Value = False
    .Columns.Append "CODTALLA", adSmallInt
    .Columns("CODTALLA").Properties("Nullable").Value = False
    .Columns.Append "TEMPOR", adUnsignedTinyInt
    .Columns("TEMPOR").Properties("Nullable").Value = False
  End With
  Cat.tables.Append Tbl(7)

  '----------* Table Definition of tmp_DETTRANS *----------
  Set Tbl(8) = New ADOX.Table
  Tbl(8).ParentCatalog = Cat
  With Tbl(8)
    .Name = "tmp_DETTRANS"
   ' .Columns.Append "CODALM", adUnsignedTinyInt
   '   .Columns("CODALM").Properties("Nullable").Value = False
    .Columns.Append "CODART", adSmallInt
      .Columns("CODART").Properties("Nullable").Value = False
    .Columns.Append "CODCOL", adSmallInt
      .Columns("CODCOL").Properties("Nullable").Value = False
   ' .Columns.Append "CODIGO", adInteger
   '   .Columns("CODIGO").Properties("Nullable").Value = False
    .Columns.Append "CODTALLA", adSmallInt
    .Columns("CODTALLA").Properties("Nullable").Value = False
    '.Columns.Append "FMODI", adDate
    '  .Columns("FMODI").Properties("Nullable").Value = False
  '  .Columns.Append "ID", adInteger
  '    .Columns("ID").Properties("Nullable").Value = False
    .Columns.Append "TEMPOR", adUnsignedTinyInt
    .Columns("TEMPOR").Properties("Nullable").Value = False
    .Columns.Append "UNIDADES", adSingle
    .Columns("UNIDADES").Properties("Nullable").Value = False
  End With
  Cat.tables.Append Tbl(8)

  Set Cat = Nothing
  
  End If
  
tmpCnn.Open sCnn
  
'borrar las que faltan, por si se vuelve a pasar otra vez
tmpCnn.Execute "DELETE FROM DIFERENCIAS WHERE CAUSA = 2"

'borrar las filas de transferencia (se volveran a cargar)
tmpCnn.Execute "DELETE FROM tmp_DETTRANS"

DoEvents
  
'abrir un recordset a a las tablas con las que vamos a trabajar
rc.Open "SELECT * from tmp_DETTRANS", tmpCnn, adOpenStatic, adLockOptimistic
rcRec.Open "SELECT * from Recibidos", tmpCnn, adOpenStatic, adLockOptimistic

If rcRec.RecordCount > 0 Then
    Call carga_grid_int
End If

rcDif.Open "SELECT * from Diferencias", tmpCnn, adOpenStatic, adLockOptimistic

If rcDif.RecordCount > 0 Then
    Call carga_grid_dif
End If

   On Error GoTo 0
   Exit Sub

CreateDatabaseTemporal_Error:
    
    Set Cat = Nothing
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento CreateDatabaseTemporal de Formulario frmComPtrans"

End Sub


'---------------------------------------------------------------------------------------
' Subrutina   : cargar_transferencia
' Fecha/Hora  : 13/01/2004 21:30
' Autor       : JCASTILLO
' Propósito   : Carga la transferencia actual al fichero de trabajo
'---------------------------------------------------------------------------------------
Private Sub cargar_transferencia()
Dim tmprc As New ADODB.Recordset

   On Error GoTo cargar_transferencia_Error
    
    'buscar todos los registros de esta transferencia
    tmprc.Open "select CODART,TEMPOR,CODTALLA,CODCOL,UNIDADES from DETTRANS WHERE CODIGO =" & ID_TRANSF & " AND CODALM = " & ALM_TRANSF, locCnn
   
    With rc
        
        Do Until tmprc.EOF
        
            .AddNew
            
            .fields("CODART").Value = tmprc.fields("CODART").Value
            .fields("TEMPOR").Value = tmprc.fields("TEMPOR").Value
            .fields("CODTALLA").Value = tmprc.fields("CODTALLA").Value
            .fields("CODCOL").Value = tmprc.fields("CODCOL").Value
            .fields("UNIDADES").Value = tmprc.fields("UNIDADES").Value
                        
            .Update
            
            If Not tmprc.EOF Then tmprc.MoveNext
        
        Loop
    
    End With
    
    tmprc.Close
    Set tmprc = Nothing

   On Error GoTo 0
   Exit Sub

cargar_transferencia_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cargar_transferencia de Formulario frmComPtrans"
End Sub


'---------------------------------------------------------------------------------------
' Subrutina   : carga_etiquetas
' Fecha/Hora  : 19/03/2004 16:07
' Autor       : JCASTILLO
' Propósito   : Mostrar el número de unidades que sobran y faltan de esta transferencia
'---------------------------------------------------------------------------------------
Private Sub carga_etiquetas()
Dim tmpsobran As Variant
Dim tmpfaltan As Variant

   On Error GoTo carga_etiquetas_Error

tmpsobran = devuelve_campo("SELECT SUM(UNIDADES) FROM DIFERENCIAS WHERE CAUSA = 1", tmpCnn)
tmpfaltan = devuelve_campo("SELECT SUM(UNIDADES) FROM DIFERENCIAS WHERE CAUSA = 2", tmpCnn)

If tmpsobran <> "@" Then
    lblSobran.Caption = "Sobran: " & tmpsobran
Else
    lblSobran.Caption = "Sobran: 0"
End If

If tmpfaltan <> "@" Then
    lblFaltan.Caption = "Faltan: " & tmpfaltan
Else
    lblFaltan.Caption = "Faltan: 0"
End If

   On Error GoTo 0
   Exit Sub

carga_etiquetas_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_etiquetas de Formulario frmComPtrans"

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ultimo As Long

   On Error GoTo Form_KeyDown_Error

  Select Case KeyCode
    
    'borrar ultimo articul0o
    Case vbKeyB
      
      If rcRec.EOF Then Exit Sub
      ultimo = devuelve_campo("select max(id) from Recibidos", tmpCnn)
      
      tmpCnn.Execute "DELETE FROM RECIBIDOS WHERE ID = " & ultimo
      fg.RemoveItem 2
      DoEvents
      rcRec.Requery
      
      fg.subtotal flexSTCount, , 4, , vbBlue, vbWhite
      fg.TextMatrix(1, 3) = "Total"
      fg.TextMatrix(1, 1) = ""
    
      ioCODBAR.Text = ""
      
      lblstatus.Caption = "Se ha borrado un artículo"
     
    'intentar terminar transferencia ...
    Case vbKeyA
           
      DoEvents
      
      If Not Comprobada Then
        Comprobada = True
      Else
        lblstatus.Caption = "La transferencia ya ha sido comprobada"
        ioCODBAR.Text = ""
        ioCODBAR.Valor = ""
        Exit Sub
      End If
      
      Call comprueba_unidades
      
      ioCODBAR.Text = ""
      ioCODBAR.Valor = ""
      
      rcDif.Requery
      rc.Requery
      
      'mostrar los nºs de diferencias
      Call carga_etiquetas
      
      'si no hay diferencias, permitir aceptar la transferencia normalmente
      If (rcDif.RecordCount > 0) Or (fg.Rows <= 1) Then
        Me.estado = 2
        'Exit Sub
       'si hay diferencias, devolver estado = 2 (imposible aceptar transferencia)
      Else
        Me.estado = 1
        'Exit Sub
      End If
            
      'si no se han introducido diferencias
      If rcRec.RecordCount <= 0 Then
            Me.estado = 2
        Exit Sub
      End If
      
      
    
    'salir
    Case vbKeyS
    
      DoEvents
      ioCODBAR.Text = ""
      ioCODBAR.Valor = ""
      
      Call Form_KeyDown(vbKeyA, 0)
      
      Unload Me

  End Select


   On Error GoTo 0
   Exit Sub

Form_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Form_KeyDown de Formulario frmComPtrans"

End Sub



'---------------------------------------------------------------------------------------
' Procedimiento : carga_grid_int
' Fecha/Hora    : 22/03/2004 10:41
' Autor         : JCastillo
' Propósito     :  Carga el grid con los datos previamente introducidos por el usuario
'                      para esta comprobación de transferencia
'---------------------------------------------------------------------------------------
'
Private Sub carga_grid_int()
Dim tmpcodprov  As Variant
Dim tmpcodcolor  As Variant
  
   On Error GoTo carga_grid_int_Error

    With fg
        .Clear
        .Cols = 7
        .Rows = 2
        .TextMatrix(0, 1) = "Prov."
        .TextMatrix(0, 2) = "Ref."
        .TextMatrix(0, 3) = "Modelo"
        .TextMatrix(0, 4) = "Temp."
        .TextMatrix(0, 5) = "Talla"
        .TextMatrix(0, 6) = "Color"
        
    End With

    Do Until rcRec.EOF

        'articulo no existe en la transferencia.
        tmpcodprov = devuelve_campo("SELECT CODPROV FROM MAARTIC WHERE CODIGO = " & rcRec.fields("CODART") & " AND TEMPOR = " & rcRec.fields("TEMPOR"), locCnn)
      
        'articulo OK, existe.
        fg.AddItem "", 1
        fg.TextMatrix(1, 1) = Trim(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & tmpcodprov, locCnn))
        fg.TextMatrix(1, 2) = Trim(devuelve_campo("SELECT REF FROM MAARTIC WHERE CODIGO = " & rcRec.fields("CODART") & " AND TEMPOR = " & rcRec.fields("TEMPOR"), locCnn))
        
        fg.TextMatrix(1, 3) = rcRec.fields("CODART") & " " & Trim(devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & rcRec.fields("CODART") & " AND TEMPOR = " & rcRec.fields("TEMPOR"), locCnn))
        fg.TextMatrix(1, 4) = rcRec.fields("TEMPOR") & " " & Trim(devuelve_campo("SELECT ABREVIA FROM TEMPOR WHERE IDTEM = " & rcRec.fields("TEMPOR"), locCnn))
        fg.TextMatrix(1, 5) = Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & rcRec.fields("CODTALLA"), locCnn))
        fg.TextMatrix(1, 6) = Trim(devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & rcRec.fields("CODCOL"), locCnn))
        tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & rcRec.fields("CODCOL"), locCnn)
        fg.Row = 1
        fg.Col = 6
        fg.CellBackColor = tmpcodcolor
        fg.Col = 3
    
        rcRec.MoveNext
    
    Loop
    
    fg.AutoSize 1, fg.Cols - 1
    
    fg.subtotal flexSTCount, , 6, , vbBlue, vbWhite
    fg.TextMatrix(1, 5) = "Total"
    fg.TextMatrix(1, 1) = ""
    

   On Error GoTo 0
   Exit Sub

carga_grid_int_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_grid_int de Formulario frmComPtrans"

End Sub

Private Sub carga_grid_dif()
Dim tmpcodcolor As Variant
Dim tmpcodprov As Variant

 
   On Error GoTo carga_grid_dif_Error

    'articulo no existe en la transferencia.
    tmpcodprov = devuelve_campo("SELECT CODPROV FROM MAARTIC WHERE CODIGO = " & rcDif.fields("CODART") & " AND TEMPOR = " & rcDif.fields("TEMPOR"), locCnn)
    
    With fgDif
    
        .Clear
        .Cols = 9
        .Rows = 1
        .TextMatrix(0, 1) = "Prov."
        .TextMatrix(0, 2) = "Ref."
        .TextMatrix(0, 3) = "Modelo"
        .TextMatrix(0, 4) = "Temp."
        .TextMatrix(0, 5) = "Talla"
        .TextMatrix(0, 6) = "Color"
        .TextMatrix(0, 7) = "Uds."
        .TextMatrix(0, 8) = "Causa"
                
    Do Until rcDif.EOF
        
        .AddItem "", 1
        .TextMatrix(1, 1) = Trim(devuelve_campo("SELECT NOMBRE FROM MAPROV WHERE CODIGO = " & tmpcodprov, locCnn))
        .TextMatrix(1, 2) = Trim(devuelve_campo("SELECT REF FROM MAARTIC WHERE CODIGO = " & rcDif.fields("CODART") & " AND TEMPOR = " & rcDif.fields("TEMPOR"), locCnn))
        .TextMatrix(1, 3) = Format(rcDif.fields("CODART"), "00000") & " " & devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & rcDif.fields("CODART") & " AND TEMPOR = " & rcDif.fields("TEMPOR"), locCnn)
        .TextMatrix(1, 4) = Format(rcDif.fields("TEMPOR"), "000") & " " & devuelve_campo("SELECT ABREVIA FROM TEMPOR WHERE IDTEM = " & rcDif.fields("TEMPOR"), locCnn)
        .TextMatrix(1, 5) = devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & rcDif.fields("CODTALLA"), locCnn)
        .TextMatrix(1, 6) = devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & rcDif.fields("CODCOL"), locCnn)
        .TextMatrix(1, 7) = rcDif.fields("UNIDADES")
        
        If rcDif.fields("CAUSA") = 1 Then
            .TextMatrix(1, 8) = "SOBRA"
        Else
            .TextMatrix(1, 8) = "FALTA"
        End If
        
       tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & rcDif.fields("CODCOL"), locCnn)
    
    If tmpcodcolor <> "@" Then
        .Row = 1
        .Col = 6
        .CellBackColor = tmpcodcolor
        .Col = 3
    End If

            
        rcDif.MoveNext
  
   Loop
  
  
        .SubtotalPosition = flexSTAbove
        .subtotal flexSTCount, , 5, , vbBlue, vbWhite
        .subtotal flexSTSum, , 7, , vbBlue, vbWhite
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 4) = "Total"
        .TextMatrix(1, 6) = "Uds:"
        .AutoSize 1, .Cols - 1
        
  End With

   Exit Sub

carga_grid_dif_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_grid_dif de Formulario frmComPtrans"

End Sub

  'grabar la diferencia como no se encuentra
                'With rcDif
               '     .AddNew
              ''      .Fields("CODART") = rcbus.Fields("CODART")
               '     .Fields("TEMPOR") = rcbus.Fields("TEMPOR")
               '     .Fields("CODTALLA") = rcbus.Fields("CODTALLA")
               '     .Fields("CODCOL") = rcbus.Fields("CODCOL")
               '     'grabar las unidades que sobran
               '     .Fields("UNIDADES") = rcbus.Fields("UNIDADES") - rctrn.Fields("UNIDADES") 'una unidad
               '     .Fields("CAUSA") = 1 'no se encuentra en la transferencia (sobran)
               '     .Update
               ' End With
                
               ' With fgDif
            '
            '    .AddItem "", 2
            '    .TextMatrix(2, 1) = Format(rcbus.Fields("CODART"), "00000") & " " & devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & rcbus.Fields("CODART") & " AND TEMPOR = " & rcbus.Fields("TEMPOR"), locCnn)
            '    .TextMatrix(2, 2) = rcbus.Fields("TEMPOR") & " " & devuelve_campo("SELECT ABREVIA FROM TEMPOR WHERE IDTEM = " & rcbus.Fields("TEMPOR"), locCnn)
            '    .TextMatrix(2, 3) = devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & rcbus.Fields("CODTALLA"), locCnn)
            '    .TextMatrix(2, 4) = devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & rcbus.Fields("CODCOL"), locCnn)
            '
            '    tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & rcbus.Fields("CODCOL"), locCnn)
            '    If tmpcodcolor <> "@" Then
            '        .Row = 2
            '        .Col = 4
            '        .CellBackColor = tmpcodcolor
            '        .Col = 3
            '    End If
            '
            '    .subtotal flexSTCount, , 4, , vbBlue, vbWhite
        '        .TextMatrix(1, 1) = ""
        '        .TextMatrix(1, 3) = "Total"
        '
            '    .AutoSize 1, fgDif.Cols - 1
            ' End With
              
                              '''''''''''''''''''''''''''''''''''''''''''''''''
              'AQUI QUITE
              'si es menor (han mandado de menos)
            '  Case Is < rctrn.Fields("UNIDADES")
              
                                
                'Call añade_diferencia(rcbus.Fields("CODART"), rcbus.Fields("TEMPOR"), rcbus.Fields("CODTALLA"), rcbus.Fields("CODCOL"), rctrn.Fields("UNIDADES") - rcbus.Fields("UNIDADES"), 2)
                
                '''''''''''''''''''''''''''''''''''''''''''''''''
                
                'grabar la diferencia como que falta
                'With rcDif
                '    .AddNew
                '    .Fields("CODART") = rcbus.Fields("CODART")
                '    .Fields("TEMPOR") = rcbus.Fields("TEMPOR")
                '    .Fields("CODTALLA") = rcbus.Fields("CODTALLA")
                '    .Fields("CODCOL") = rcbus.Fields("CODCOL")
                '    'grabar las unidades que faltan
                '    .Fields("UNIDADES") = rctrn.Fields("UNIDADES") - rcbus.Fields("UNIDADES") 'una unidad
                '    .Fields("CAUSA") = 2 'falta de la transferencia (faltan)
                '    .Update
                'End With
                
        '        With fgDif
            
        '        .AddItem "", 2
        '        .TextMatrix(2, 1) = Format(rcbus.Fields("CODART"), "00000") & " " & devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & rcbus.Fields("CODART") & " AND TEMPOR = " & rcbus.Fields("TEMPOR"), locCnn)
        '        .TextMatrix(2, 2) = rcbus.Fields("TEMPOR") & " " & devuelve_campo("SELECT ABREVIA FROM TEMPOR WHERE IDTEM = " & rcbus.Fields("TEMPOR"), locCnn)
        '        .TextMatrix(2, 3) = devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & rcbus.Fields("CODTALLA"), locCnn)
        '        .TextMatrix(2, 4) = devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & rcbus.Fields("CODCOL"), locCnn)
        '
        '        tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & rcbus.Fields("CODCOL"), locCnn)
        '        If tmpcodcolor <> "@" Then
        '            .Row = 2
        '            .Col = 4
        '            .CellBackColor = tmpcodcolor
        '            .Col = 3
        '        End If
        '
        '        .subtotal flexSTCount, , 4, , vbBlue, vbWhite
        '        .TextMatrix(1, 1) = ""
        '        .TextMatrix(1, 3) = "Total"
        '
                '.AutoSize 1, fgDif.Cols - 1
                'End With
              


            'With rcDif
            '    .AddNew
            '    .Fields("CODART") = rcbus.Fields("CODART")
            '    .Fields("TEMPOR") = rcbus.Fields("TEMPOR")
            '    .Fields("CODTALLA") = rcbus.Fields("CODTALLA")
            '    .Fields("CODCOL") = rcbus.Fields("CODCOL")
            '    .Fields("UNIDADES") = 1 'graban las unidades que sobran
            '    .Fields("CAUSA") = 1 'no se encuentra en la transferencia
            '    .Update
            'End With
            
            'articulo no existe en la transferencia.
            'With fgDif
            '
            '    .AddItem "", 2
            '    .TextMatrix(2, 1) = Format(rcbus.Fields("CODART"), "00000") & " " & devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & rcbus.Fields("CODART") & " AND TEMPOR = " & rcbus.Fields("TEMPOR"), locCnn)
            '    .TextMatrix(2, 2) = rcbus.Fields("TEMPOR") & " " & devuelve_campo("SELECT ABREVIA FROM TEMPOR WHERE IDTEM = " & rcbus.Fields("TEMPOR"), locCnn)
            '    .TextMatrix(2, 3) = devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & rcbus.Fields("CODTALLA"), locCnn)
            '    .TextMatrix(2, 4) = devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & rcbus.Fields("CODCOL"), locCnn)
            '
            '    tmpcodcolor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & rcbus.Fields("CODCOL"), locCnn)
            '    If tmpcodcolor <> "@" Then
            '        .Row = 2
            '        .Col = 4
            '        .CellBackColor = tmpcodcolor
            '        .Col = 3
            '    End If
        '    '
        '        .subtotal flexSTCount, , 4, , vbBlue, vbWhite
        '        .TextMatrix(1, 1) = ""
        '        .TextMatrix(1, 3) = "Total"
        '
            '    .AutoSize 1, fgDif.Cols - 1
            'End With

