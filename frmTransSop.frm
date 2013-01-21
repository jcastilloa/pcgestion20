VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTransSop 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preparar Transferencia en Soporte Magnético"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11415
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
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   11415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox opIncluirArt 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Incluir Artículos"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   2145
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4635
      Width           =   1125
   End
   Begin VB.DirListBox Dir1 
      Height          =   3060
      Left            =   7950
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1965
      Width           =   3435
   End
   Begin VB.DriveListBox Drive1 
      Height          =   420
      Left            =   7965
      TabIndex        =   6
      Top             =   1500
      Width           =   3405
   End
   Begin PCGestion.chameleonButton cbCancelar 
      Height          =   795
      Left            =   6975
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4620
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "&Terminar"
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
      MICON           =   "frmTransSop.frx":0000
      PICN            =   "frmTransSop.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblDestino 
      Height          =   345
      Left            =   1155
      Top             =   30
      Width           =   8655
      _ExtentX        =   15266
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
   Begin PCGestion.chameleonButton cbGenerar 
      Height          =   390
      Left            =   9840
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   15
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   688
      BTYPE           =   3
      TX              =   "&Generar"
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
      MICON           =   "frmTransSop.frx":08F6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VSFlex8LCtl.VSFlexGrid fg 
      Height          =   3105
      Left            =   30
      TabIndex        =   5
      Top             =   1125
      Width           =   7905
      _cx             =   13944
      _cy             =   5477
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
      BackColorAlternate=   13613178
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
      Rows            =   1
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmTransSop.frx":0912
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
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   4035
      Top             =   2340
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin PCGestion.bsGradientLabel lblStatus 
      Height          =   345
      Left            =   60
      Top             =   4245
      Width           =   7875
      _ExtentX        =   13891
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
   Begin PCGestion.chameleonButton cmEstablecer 
      Height          =   345
      Left            =   7935
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5055
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   609
      BTYPE           =   9
      TX              =   "&Establecer"
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
      MICON           =   "frmTransSop.frx":09C2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.miCombo cbCODALMDEST 
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   495
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PCGestion.miCombo cbTRANSFER 
      Height          =   495
      Left            =   7170
      TabIndex        =   1
      Top             =   510
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PCGestion.chameleonButton cmQuitarTRN 
      Height          =   390
      Left            =   30
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4635
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   688
      BTYPE           =   3
      TX              =   "&Quitar Transferencia"
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
      MICON           =   "frmTransSop.frx":09DE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.miText ioFECHA 
      Height          =   525
      Left            =   4035
      TabIndex        =   2
      Top             =   4785
      Width           =   1305
      _ExtentX        =   2302
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
   Begin PCGestion.chameleonButton cbCargarDesdeFecha 
      Height          =   795
      Left            =   5400
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4620
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "&Introducir todas desde Fecha"
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
      MICON           =   "frmTransSop.frx":09FA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbQuitarTodas 
      Height          =   390
      Left            =   30
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5040
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   688
      BTYPE           =   3
      TX              =   "Quitar &Todas"
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
      MICON           =   "frmTransSop.frx":0A16
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA"
      Height          =   285
      Left            =   3315
      TabIndex        =   13
      Top             =   4860
      Width           =   675
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSF."
      Height          =   285
      Left            =   6345
      TabIndex        =   11
      Top             =   585
      Width           =   840
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ALMACEN DESTINO"
      Height          =   570
      Left            =   1215
      TabIndex        =   10
      Top             =   420
      Width           =   990
   End
   Begin VB.Label ioFMODI 
      Alignment       =   2  'Center
      BackColor       =   &H00EEA78E&
      Caption         =   "Guardar en"
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   7965
      TabIndex        =   8
      Top             =   1125
      Width           =   3405
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   15
      Picture         =   "frmTransSop.frx":0A32
      Top             =   0
      Width           =   1125
   End
End
Attribute VB_Name = "frmTransSop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Módulo       : frmTransSop
' Fecha/Hora   : 19/02/2004 13:13
' Autor        : JCastillo
' Propósito    : Prepara transferencias en soporte magnetico
'---------------------------------------------------------------------------------------
Option Explicit

Dim destino As String
Dim fichero As String
Dim f_path As String
Dim generada As Boolean

Private Sub cbCancelar_Click()
    Unload Me
End Sub

Private Sub cbDestino_Click()

With Dialogo
    .DialogTitle = "Seleccione destino para los datos ..."
    .filename = destino
    .ShowSave
    
    If (.CancelError = False) And Trim(.filename) <> "" Then
    
        'si el fichero ya existe ...
        If Dir(.filename) <> "" Then
            
            'preguntar al usuario ...
            If MsgBox("El fichero: " & Chr(13) & .filename & Chr(13) & "ya existe. ¿Desea sobreescribirlo?", vbQuestion + vbYesNo, titulo) = vbYes Then
            
            End If
        
        End If
        
        destino = .filename
        lblDestino.Caption = destino
    
    End If
    
End With

End Sub

'---------------------------------------------------------------------------------------
' Subrutina   : cbCargarDesdeFecha_Click
' Fecha/Hora  : 28/03/2004 19:13
' Autor       : JCASTILLO
' Propósito   : Cargar las transferencias a partir de una fecha
'---------------------------------------------------------------------------------------
Private Sub cbCargarDesdeFecha_Click()
Dim mrc As New ADODB.Recordset

   On Error GoTo cbCargarDesdeFecha_Click_Error

If ioFECHA.Text = "" Then
    lblstatus.Caption = "Debe introducir una FECHA"
    ioFECHA.SetFocus
    Exit Sub
End If

If cbCODALMDEST.Text = "" Then
    lblstatus.Caption = "Debe introducir ALMACEN de destino"
    cbCODALMDEST.SetFocus
    Exit Sub
End If

If fg.Rows > 1 Then Call cbQuitarTodas_Click

'introducir las transferencias por fecha
mrc.Open "SELECT CODIGO FROM PTRANS WHERE (ESTADO = 1) AND (FMODI >= '" & Format(ioFECHA.Text, "yyyymmdd") & "') AND (CODALMORIG = " & AlmacenActual & ") AND (CODALMDEST = " & cbCODALMDEST.Text & ") ORDER BY CODIGO DESC", locCnn, adOpenStatic, adLockReadOnly

If mrc.RecordCount <= 0 Then

    lblstatus.Caption = "No se encuentran transferencias desde la fecha dada"
    
    mrc.Close
    Set mrc = Nothing
    
    Exit Sub
    
End If

Do Until mrc.EOF
    Call añade_linea_grid(mrc.fields("CODIGO"))
    mrc.MoveNext
Loop

lblstatus.Caption = "Se han introducido: " & mrc.RecordCount & " transferencias"

mrc.Close
Set mrc = Nothing

   On Error GoTo 0
   Exit Sub

cbCargarDesdeFecha_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbCargarDesdeFecha_Click de Formulario frmTransSop"

End Sub

Private Sub cbcodalmdest_Validate(Cancel As Boolean)
    
    If cbCODALMDEST.Text <> "" Then
    
    With cbTRANSFER
        .ConexionString = locCnn
        .SQLString = "SELECT CODIGO, CAST(TOTAL AS VARCHAR) + ' " & SimboloMoneda & "' + ' Dcto:' + CAST(DCTO AS VARCHAR) + ' %' FROM PTRANS WHERE ESTADO = 1 AND CODALMORIG = " & AlmacenActual & " AND CODALMDEST = " & cbCODALMDEST.Text & " ORDER BY CODIGO DESC"
        .LenCodigo = 8
        .CodigoWidth = 1000
        .carga
    End With
    
    Else
    
        cbTRANSFER.borra_combo
       
    End If
    
End Sub

Private Sub cbGenerar_Click()
Dim var As Long
Dim incluir As Boolean

On Error GoTo cbGenerar_Click_Error

With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
End With

lblstatus.Caption = "Generando fichero ..."
DoEvents

If opIncluirArt.Value = 1 Then incluir = True

Call CreateDatabaseTRN(destino)
'Exit Sub

Select Case fg.Rows

'si hay solo una transferencia (titlo + 1 transferencia)
'hacer todo en una linea
Case Is = 3
    Call Crea_TRN_Datos(fg.TextMatrix(2, 1), AlmacenActual, cbCODALMDEST.Text, False, True, destino, locCnn, incluir)

'solo crear fichero
Case Is > 3

    Call Crea_TRN_Datos(fg.TextMatrix(2, 1), AlmacenActual, cbCODALMDEST.Text, False, False, destino, locCnn, incluir)
    
    For var = 3 To fg.Rows - 1
    
       'si es la ultima fila, comprimir
       If var = fg.Rows - 1 Then
        Call Crea_TRN_Datos(fg.TextMatrix(var, 1), AlmacenActual, cbCODALMDEST.Text, False, True, destino, locCnn, incluir)
       'de lo contrario simplemente añadir registros
       Else
        Call Crea_TRN_Datos(fg.TextMatrix(var, 1), AlmacenActual, cbCODALMDEST.Text, False, False, destino, locCnn, incluir)
       End If
    
    Next var
    
End Select
generada = True

lblstatus.Caption = "Se ha generado correctamente"

'si quiere copiar la transferencia al disco de A:
If MsgBox("¿Desea guardar la transferencia en el Disco A:?" & Chr(13) & "(Introduzca un disco formateado en la unidad, y pulse SI para guardar en el disco)", vbQuestion + vbYesNo, titulo) = vbNo Then Exit Sub

If MsgBox("Desea formatear un DISQUETE (introduzca un disquete). ATENCION: al formatear un diskette se borraran todos los datos que este contenia previamente.", vbQuestion + vbYesNo, titulo) = vbYes Then
    lblstatus.Caption = "Formateando Diskette en A:\ ..."
    DoEvents
    Call FormatDrive(Me, "A:\")
    DoEvents
    lblstatus.Caption = "Formato completado"
End If

'si existe, quitar posibles atributos, y borrar
If Dir("A:\" & fichero & ".trnz") <> "" Then
    SetAttr "A:\" & fichero & ".trnz", vbNormal
    Kill "A:\" & fichero & ".trnz"
End If

'copiar al disco de a la transferencia ...
FileCopy f_path & fichero & ".trnz", "A:\" & fichero & ".trnz"
DoEvents

lblstatus.Caption = "Se ha generado y copiado al disco correctamente"
DoEvents

   On Error GoTo 0
   Exit Sub

cbGenerar_Click_Error:

    lblstatus.Caption = ""
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbGenerar_Click de Formulario frmTransSop"

End Sub




Private Sub cbQuitarTodas_Click()

'si no la quiere quitar, salir
If MsgBox("¿Desea quitar la lista de transferencias actual?", vbQuestion + vbYesNo, titulo) = vbNo Then Exit Sub

With fg
    .Clear
    .Rows = 1
    .Cols = 6
    .TextMatrix(0, 1) = "CODIGO"
    .TextMatrix(0, 2) = "ORIGEN"
    .TextMatrix(0, 3) = "DESTINO"
    .TextMatrix(0, 4) = "DCTO"
    .TextMatrix(0, 5) = "TOTAL"
End With

End Sub

Private Sub cbTRANSFER_Validate(Cancel As Boolean)
Dim var As Long


If cbTRANSFER.Text <> "" Then

    If fg.Rows > 1 Then
    For var = 1 To fg.Rows - 1
    
        If fg.TextMatrix(var, 1) = cbTRANSFER.Text Then
            
            MsgBox "La transferencia seleccionada ya ha sido incluida en este envío", vbInformation, titulo
            Exit Sub
        
        End If
    
    Next var
    End If

    Call añade_linea_grid(cbTRANSFER.Text)
    cbCODALMDEST.Locked = True

End If

End Sub

Private Sub añade_linea_grid(codigo As Long)
Dim t_trans As Variant

With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
End With

fg.Rows = fg.Rows + 1

t_trans = devuelve_matriz("Select TOTAL, DCTO FROM PTRANS WHERE CODIGO = " & codigo & " AND CODALMORIG = " & AlmacenActual, locCnn)

If Not IsArray(t_trans) Then
    lblstatus.Caption = "Error al introducir la transferencia"
    Exit Sub
End If

'codigo
fg.TextMatrix(fg.Rows - 1, 1) = Format(codigo, "00000000")

'origen
fg.TextMatrix(fg.Rows - 1, 2) = Trim(devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & AlmacenActual, locCnn))

'destino
fg.TextMatrix(fg.Rows - 1, 3) = Trim(devuelve_campo("SELECT DESCRIPCION FROM ALMACENES WHERE CODIGO = " & cbCODALMDEST.Text, locCnn))

'dcto
fg.TextMatrix(fg.Rows - 1, 4) = t_trans(1) & " %"

'importe
fg.TextMatrix(fg.Rows - 1, 5) = t_trans(0) - ((t_trans(0) * t_trans(1)) / 100)

fg.SubtotalPosition = flexSTAbove
fg.subtotal flexSTSum, , 5, "Currency", vbBlue, vbWhite, True
fg.subtotal flexSTCount, , 3, , vbBlue, vbWhite, True
'fg.Subtotal , , , ,

DoEvents
fg.AutoSize 1, fg.Cols - 1

If cbCODALMDEST.Locked = False Then cbCODALMDEST.Locked = True

End Sub


Private Sub cmEstablecer_Click()

    If MsgBox("¿Desea guardar la transferencia en: " & Dir1.Path & "?", vbQuestion + vbYesNo, titulo) = vbNo Then Exit Sub
    
    f_path = Dir1.Path
    
    'si ya tiene la barra no ponersela otra vez
    If Right(f_path, 1) = "\" Then
        destino = f_path & fichero & ".trz"
        lblDestino.Caption = f_path & fichero & ".trnz"
        
    Else
        destino = f_path & "\" & fichero & ".trz"
        lblDestino.Caption = f_path & "\" & fichero & ".trnz"
    End If


End Sub

Private Sub Command1_Click()
fg.Rows = fg.Rows - 1
End Sub

Private Sub cmQuitarTRN_Click()

If fg.Row > 1 Then fg.RemoveItem fg.Row

DoEvents
fg.AutoSize 1, fg.Cols - 1

End Sub



Private Sub Drive1_change()

   On Error GoTo Drive1_change_Error

    Dir1.Path = Drive1.Drive

   On Error GoTo 0
   Exit Sub

Drive1_change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Drive1_change de Formulario frmTransSop"
    
End Sub

Private Sub Form_Load()
Dim direc As String

    On Error GoTo Form_Load_Error
   
    With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
    End With
    
    If Dir("C:\TRANSFERENCIAS\", vbDirectory) = "" Then
        MkDir "C:\TRANSFERENCIAS"
    End If

    direc = "TRANSFERENCIAS"
    f_path = "c:\TRANSFERENCIAS\"
    
    ChDrive (Left(f_path, 1))
    
    'If Dir(f_path) = "" Then MkDir direc
    
    ChDir (f_path)
    
    'poner ruta por defecto
    Dir1.Path = f_path
    Drive1.Drive = f_path
        
    fichero = "TRN-" & CajaActual & "-" & Format(Now, "dd-mm-yyyy-hh-mm-ss")
    
    destino = f_path & fichero & ".trz"
    lblDestino.Caption = f_path & fichero & ".trnz"
    
    With cbCODALMDEST
        .ConexionString = locCnn
        .SQLString = "SELECT CODIGO, DESCRIPCION FROM ALMACENES WHERE MBAJA = 0 ORDER BY DESCRIPCION"
        .LenCodigo = 3
        .CodigoWidth = 500
        .carga
    End With
    
    With ioFECHA
        .dspFormat = "dd/mm/yyyy"
        .LongMaxima = 10
    End With
    
    fg.SelectionMode = flexSelectionByRow
    fg.HighLight = flexHighlightWithFocus
        
    fg.ColFormat(5) = "Currency"
    fg.ColAlignment(4) = flexAlignRightCenter

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Form_Load de Formulario frmTransSop"
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If Not generada Then
    If MsgBox("ATENCION, va a salir SIN generar ninguna transferencia. ¿Esta seguro?.", vbQuestion + vbYesNo, titulo) = vbNo Then
        Cancel = True
        Exit Sub
    End If
End If

Set frmTransSop = Nothing

End Sub
