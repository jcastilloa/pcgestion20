VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmDeudCli 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Deudas de Clientes ..."
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11415
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
   ScaleHeight     =   6555
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin PCGestion.bsGradientLabel lblStatus 
      Height          =   345
      Left            =   45
      Top             =   5355
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   609
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   12632256
      Colour2         =   16558731
      CaptionAlignment=   1
   End
   Begin VSFlex8Ctl.VSFlexGrid fgArt 
      Height          =   3435
      Left            =   3000
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1890
      Width           =   4620
      _cx             =   8149
      _cy             =   6059
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
      FormatString    =   $"frmDeudCli.frx":0000
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
   Begin PCGestion.chameleonButton cbCancelar 
      Height          =   795
      Left            =   6180
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5730
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
      MICON           =   "frmDeudCli.frx":00DE
      PICN            =   "frmDeudCli.frx":00FA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbImprimir 
      Height          =   795
      Left            =   4305
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5730
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
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
      MICON           =   "frmDeudCli.frx":09D4
      PICN            =   "frmDeudCli.frx":09F0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VSFlex8Ctl.VSFlexGrid fgPag 
      Height          =   3435
      Left            =   7635
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1890
      Width           =   3765
      _cx             =   6641
      _cy             =   6059
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
      FormatString    =   $"frmDeudCli.frx":16CA
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
   Begin PCGestion.bsGradientLabel lblCliente 
      Height          =   345
      Left            =   1080
      Top             =   75
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   609
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   12632256
      Colour2         =   16558731
      CaptionAlignment=   1
   End
   Begin PCGestion.chameleonButton cbAceptar 
      Height          =   795
      Left            =   5190
      TabIndex        =   5
      Top             =   5730
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "&Aceptar"
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
      MICON           =   "frmDeudCli.frx":17A8
      PICN            =   "frmDeudCli.frx":17C4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.bsGradientLabel bsGradientLabel4 
      Height          =   315
      Left            =   6705
      Top             =   90
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      Caption         =   "-C- Asignar Cliente"
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   12632256
      Colour2         =   16761024
      CaptionAlignment=   1
   End
   Begin PCGestion.bsGradientLabel lblTotal 
      Height          =   375
      Left            =   1080
      Top             =   570
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   661
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   15640462
      Colour2         =   12632256
      CaptionAlignment=   1
   End
   Begin PCGestion.bsGradientLabel lblPagado 
      Height          =   375
      Left            =   4080
      Top             =   570
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   661
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   15640462
      Colour2         =   12632256
      CaptionAlignment=   1
   End
   Begin PCGestion.bsGradientLabel lblPendiente 
      Height          =   375
      Left            =   7380
      Top             =   570
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   661
      Caption         =   ""
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   15640462
      Colour2         =   12632256
      CaptionAlignment=   1
   End
   Begin VSFlex8Ctl.VSFlexGrid fgVen 
      Height          =   3435
      Left            =   30
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1890
      Width           =   2955
      _cx             =   5212
      _cy             =   6059
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
      FormatString    =   $"frmDeudCli.frx":249E
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PENDIENTE"
      Height          =   270
      Left            =   6180
      TabIndex        =   8
      Top             =   615
      Width           =   1125
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PAGADO"
      Height          =   270
      Left            =   3075
      TabIndex        =   7
      Top             =   615
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      Height          =   270
      Left            =   255
      TabIndex        =   6
      Top             =   615
      Width           =   765
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENTE"
      Height          =   270
      Left            =   60
      TabIndex        =   4
      Top             =   90
      Width           =   915
   End
End
Attribute VB_Name = "frmDeudCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Módulo       : frmDeudCli
' Fecha/Hora : 14/05/2004 11:28
' Autor         : JCastillo
' Propósito   :  Cobrar deudas de clientes
'---------------------------------------------------------------------------------------
Option Explicit

Public Codigo_Cliente As Long
Public Caja_Cliente As Long

Dim prime As Boolean

Private Sub Form_Activate()

  If prime Then Exit Sub

  'si estan a 0 mostrar el grid de clientes
  If Caja_Cliente = 0 And Codigo_Cliente = 0 Then
    Call Abre_Grid_Clientes
  'si vienen con datos de algun otro formulario, mostrar el nombre directamente
  ElseIf Caja_Cliente > 0 And Codigo_Cliente > 0 Then
   lblCliente.Caption = devuelve_campo("SELECT RAZO FROM CLIENTES WHERE CODIGO = " & Codigo_Cliente & " AND CODCAJA = " & Caja_Cliente, locCnn)
    'cargar ventas para el cliente
    Call carga_totales_ventas_pendientes
  End If
  
  prime = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
      
   Select Case KeyCode
   
      'Asignar Cliente ...
      Case vbKeyC

       'abre el grid de los clientes
       Call Abre_Grid_Clientes
        KeyCode = 0
    
      Case vbKeyEscape
    
   End Select
   
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
            
    .desde_pagos = True
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

'Asigna el cliente seleccionado en el flexgrid, para llamar desde el flexclientes
Public Sub Asignar_cliente_flex(CodigoCliente As Long, CodCaja As Byte)

With frmFlexCli
    
    If .seleccionado Then
    
        'asignar valores ...
        Codigo_Cliente = CodigoCliente 'rccli.Fields("CODIGO")
        Caja_Cliente = CodCaja 'rccli.Fields("CODCAJA")
        lblCliente.Caption = devuelve_campo("SELECT RAZO FROM CLIENTES WHERE CODIGO = " & Codigo_Cliente & " AND CODCAJA = " & Caja_Cliente, locCnn)
        
        'cargar ventas para el cliente
        Call carga_totales_ventas_pendientes
    
    'dejar como estaba
    'Else
    
      '  rc.Fields("CODCLI") = Null
      '  rc.Fields("CAJACLI") = Null
      '  lblCliente.Caption = ""
        
    End If
    
End With
    
     '   rc.Update
        
End Sub



'---------------------------------------------------------------------------------------
' Procedimiento : carga_totales_ventas_pendientes
' Fecha/Hora     : 14/05/2004 11:05
' Autor             : JCastillo
' Propósito       : Carga una lista-resumen de las ventas pendientes en el listbox
'---------------------------------------------------------------------------------------
Private Sub carga_totales_ventas_pendientes()
Dim rc As New ADODB.Recordset

   On Error GoTo carga_totales_ventas_pendientes_Error

    With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
    End With
  
    rc.Open "SELECT FMODI, CODVEN, IMPORTE, PAGADO FROM CABDEUDCLI WHERE CODCLI = " & Codigo_Cliente & " AND CAJACLI = " & Caja_Cliente & " AND ESTADO IN (0, 1)", locCnn, adOpenStatic, adLockReadOnly
    
    'Poner titulos antes de nada
    With fgVen
        .Clear
        .Rows = 1
        .Cols = 6
        
        .ColHidden(5) = True
        .ColFormat(2) = "Currency"
        .ColFormat(4) = "Currency"
              
        .TextMatrix(0, 1) = "FECHA"
        .TextMatrix(0, 2) = "IMPORTE"
        .TextMatrix(0, 3) = "PAGADO"
        .TextMatrix(0, 4) = "TICKET"
     
    
    
    If rc.RecordCount < 0 Then Exit Sub
    
    Do Until rc.EOF
    
        .TextMatrix(.Rows - 1, 1) = Format(rc.fields("FMODI"), "dd/mm/yyyy")
        .TextMatrix(.Rows - 1, 2) = rc.fields("IMPORTE")
        .TextMatrix(.Rows - 1, 3) = rc.fields("PAGADO")
        .TextMatrix(.Rows - 1, 4) = CStr(rc.fields("CODVEN")) & Format(Caja_Cliente, "000")
        .TextMatrix(.Rows - 1, 5) = rc.fields("CODVEN")
        
        fgVen.Rows = fgVen.Rows + 1
        rc.MoveNext
        
        'lstVentasPen.AddItem Format(rc.fields("FMODI"), "dd/mm/yyyy") & " - " & ". Ticket: " & CStr(rc.fields("CODVEN")) & Format(Caja_Cliente, "000") & ". " & Format(rc.fields("IMPORTE"), "Currency")
    Loop
                
        If .Rows > 1 Then
            .SubtotalPosition = flexSTAbove
            .subtotal flexSTCount, , 4, , vbBlue, vbWhite
            .subtotal flexSTSum, , 3, , vbBlue, vbWhite
            .subtotal flexSTSum, , 2, , vbBlue, vbWhite
            .TextMatrix(1, 4) = "Nº (" & .TextMatrix(1, 4) & ")"
        End If
        
    End With
    
    
    rc.Close
    Set rc = Nothing

   On Error GoTo 0
   Exit Sub

carga_totales_ventas_pendientes_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento carga_totales_ventas_pendientes de Formulario frmDeudCli"

End Sub


Private Sub Form_Load()

  With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
  End With
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Set frmDeudCli = Nothing

End Sub
