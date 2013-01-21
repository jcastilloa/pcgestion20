VERSION 5.00
Begin VB.Form frmImpRes 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Imprimir Ticket Resumen Ventas"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5820
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
   ScaleHeight     =   2655
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin PCGestion.miText ioFINI 
      Height          =   525
      Left            =   1545
      TabIndex        =   0
      Top             =   165
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
   Begin PCGestion.chameleonButton cbCancelar 
      Height          =   795
      Left            =   2925
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1800
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
      MICON           =   "frmImpRes.frx":0000
      PICN            =   "frmImpRes.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.miText ioFFIN 
      Height          =   525
      Left            =   4245
      TabIndex        =   1
      Top             =   165
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
   Begin PCGestion.chameleonButton cbImprimir 
      Height          =   795
      Left            =   1980
      TabIndex        =   3
      Top             =   1800
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "&Imprimir"
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
      MICON           =   "frmImpRes.frx":08F6
      PICN            =   "frmImpRes.frx":0912
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblstatus 
      Height          =   375
      Left            =   0
      Top             =   1365
      Width           =   5805
      _ExtentX        =   10239
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
   Begin PCGestion.miCombo cbCODCAJA 
      Height          =   480
      Left            =   615
      TabIndex        =   2
      Top             =   825
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CAJA"
      Height          =   285
      Left            =   90
      TabIndex        =   7
      Top             =   885
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Final"
      Height          =   315
      Left            =   2880
      TabIndex        =   6
      Top             =   225
      Width           =   1275
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Inicial"
      Height          =   315
      Left            =   195
      TabIndex        =   4
      Top             =   225
      Width           =   1275
   End
End
Attribute VB_Name = "frmImpRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Modulo      : frmImpRes
' Fecha/Hora  : 26/04/2004 21:16
' Autor       : JCASTILLO
' Propósito   : Imprime ticket de resumen de ventas ...
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------------------------
' Subrutina   : imprime_ticket_resumen_caja
' Fecha/Hora  : 26/04/2004 21:13
' Autor       : JCASTILLO
' Propósito   : ticket resumen con las ventas en formato dd/mm/yyyy  xxx.xx €
'
'---------------------------------------------------------------------------------------
Private Sub imprime_ticket_resumen_caja(fdesde As Date, fhasta As Date, codigo_caja As Byte, conexion As ADODB.Connection)
 
    Dim rc As ADODB.Recordset
    Dim suma As Currency
    Dim t_articulo As Variant
    Dim s_articulo As String
    Dim EmpresaActual As Integer
 
    On Error GoTo imprime_ticket_resumen_caja_Error
    
    Set rc = New ADODB.Recordset
    
    If EmpCnn.State = 1 Then EmpCnn.Close
    EmpCnn.Open strEmpCnn
    
    rc.Open "SELECT CODEMP FROM PUESTCNF", EmpCnn, adOpenStatic, adLockReadOnly
    
    EmpresaActual = rc.fields(0)
    
    rc.Close
    
    rc.Open "SELECT T_CAJAA, FECIERRE FROM CIERREDIA WHERE (FECIERRE >= '" & Format(fdesde, "yyyymmdd") & "')" & " AND (FECIERRE <= '" & Format(fhasta, "yyyymmdd") & "') AND CODCAJA =" & codigo_caja & " ORDER BY FECIERRE", conexion, adOpenStatic, adLockOptimistic
    
    If rc.RecordCount <= 0 Then
        MsgBox "No se encuentran datos entre " & Format(fdesde, "dd/mm/yyyy") & " y " & Format(fhasta, "dd/mm/yyyy"), vbExclamation, titulo
        rc.Close
        Set rc = Nothing
        Exit Sub
    End If
        
    t_articulo = devuelve_matriz("SELECT CODPOS, LOCALI, PROVIN FROM EMPRESAS WHERE ID = " & EmpresaActual, EmpCnn)
    
    If Not IsNull(t_articulo(0)) Then s_articulo = Trim(t_articulo(0))
    If Not IsNull(t_articulo(1)) Then s_articulo = s_articulo & "  " & Trim(t_articulo(1))
    If Not IsNull(t_articulo(2)) Then s_articulo = s_articulo & "  " & Trim(t_articulo(2))
    
    Printer.Font.Name = "Courier New"
    Printer.Print ""
    Printer.Print " " & Trim(devuelve_campo("SELECT RAZO FROM EMPRESAS WHERE ID = " & EmpresaActual, EmpCnn))
    Printer.Print " " & Trim(devuelve_campo("SELECT CIF FROM EMPRESAS WHERE ID = " & EmpresaActual, EmpCnn))
    Printer.Print " " & Trim(devuelve_campo("SELECT DIRECC FROM EMPRESAS WHERE ID = " & EmpresaActual, EmpCnn))
    Printer.Print " "
    Printer.Print " " & Trim(devuelve_campo("SELECT TELEF FROM EMPRESAS WHERE ID = " & EmpresaActual, EmpCnn))
    Printer.Print " Resumen de ventas"
    Printer.Print " Desde " & Format(fdesde, "dd/mm/yyyy")
    Printer.Print " Hasta " & Format(fhasta, "dd/mm/yyyy")
    
    Printer.Print ""
    Printer.Print ""
    
        
    Do Until rc.EOF
    
        Printer.Print " " & Format(rc.fields(1), "dd/mm/yyyy") & "    " & Format(rc.fields(0), "000.00") & " e"
        suma = suma + rc.fields(0)
    
        rc.MoveNext
    
    Loop
    
    rc.Close
    Set rc = Nothing
    
    Printer.Print " ----------------------"
    Printer.Print " Total:        " & Format(suma, "000.00") & " e"
    
    Printer.Print ""
    Printer.Print ""
    
    Printer.Print " Fecha: " & Now
    Printer.Print " " & Trim(devuelve_campo("SELECT DESCRIPCION FROM CAJAS WHERE CODIGO = " & codigo_caja, conexion))
    
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    
    s_articulo = ""
    Set t_articulo = Nothing
    
    Printer.EndDoc

   On Error GoTo 0
   Exit Sub

imprime_ticket_resumen_caja_Error:

    Printer.EndDoc
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento imprime_ticket_resumen_caja de Formulario frmImpRes"
    
End Sub


Private Sub cbCancelar_Click()

Unload Me

End Sub

Private Sub cbImprimir_Click()

   On Error GoTo cbImprimir_Click_Error

    If ioFINI.Text = "" Then
        lblstatus.Caption = "Fecha Inicial Inválida"
        ioFINI.SetFocus
        ioFINI.CancelarValidacion
        Exit Sub
    End If
    
    If ioFFIN.Text = "" Then
        lblstatus.Caption = "Fecha Final Inválida"
        ioFFIN.SetFocus
        ioFFIN.CancelarValidacion
        Exit Sub
    End If
    
    If cbCODCAJA.Text = "" Then
        lblstatus.Caption = "Caja Inválida"
        cbCODCAJA.SetFocus
        Exit Sub
    End If

  With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
  End With

    Call imprime_ticket_resumen_caja(CDate(ioFINI.Text), CDate(ioFFIN.Text), cbCODCAJA.Text, locCnn)
    
   On Error GoTo 0
   Exit Sub

cbImprimir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbImprimir_Click de Formulario frmImpRes"
End Sub

Private Sub Form_Load()

With ioFINI
    .LongMaxima = 10
    .dspFormat = "dd/mm/yyyy"
End With

With ioFFIN
    .LongMaxima = 10
    .dspFormat = "dd/mm/yyyy"
End With

 
  'Cargar el micombo cajas
  With cbCODCAJA
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM CAJAS WHERE MBAJA = 0 ORDER BY CODIGO"
    .LenCodigo = 1
    .CodigoWidth = 500
    .carga
    .Refresh
  End With


Select Case TipoPermiso

'normal
Case 0
        cbCODCAJA.Locked = True

End Select

'super usuario
'Case 1

        'cbCODCAJA.Locked = False
        cbCODCAJA.Text = CajaActual



End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Set frmImpRes = Nothing

End Sub
