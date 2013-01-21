VERSION 5.00
Begin VB.Form frmNuArr 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agregar Arreglo ..."
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9150
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
   ScaleHeight     =   6105
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstCostureras 
      BackColor       =   &H00EEA78E&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   30
      TabIndex        =   2
      Top             =   1005
      Width           =   9090
   End
   Begin PCGestion.miText ioNOMBRE 
      Height          =   525
      Left            =   5415
      TabIndex        =   1
      Top             =   465
      Width           =   3735
      _ExtentX        =   6588
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
   Begin PCGestion.chameleonButton cbAceptar 
      Height          =   795
      Left            =   3630
      TabIndex        =   7
      Top             =   5280
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
      MICON           =   "frmNuArr.frx":0000
      PICN            =   "frmNuArr.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbCancelar 
      Height          =   795
      Left            =   4605
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5280
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
      MICON           =   "frmNuArr.frx":0CF6
      PICN            =   "frmNuArr.frx":0D12
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblStatus 
      Height          =   390
      Left            =   30
      Top             =   4845
      Width           =   9090
      _ExtentX        =   16007
      _ExtentY        =   688
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
   Begin PCGestion.miText ioPVP 
      Height          =   525
      Left            =   3060
      TabIndex        =   5
      Top             =   4335
      Width           =   1635
      _ExtentX        =   2884
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
   Begin PCGestion.miText ioCOSTO 
      Height          =   525
      Left            =   960
      TabIndex        =   4
      Top             =   4335
      Width           =   1515
      _ExtentX        =   2672
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
   Begin PCGestion.bsGradientLabel lblArticulo 
      Height          =   390
      Left            =   45
      Top             =   15
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   688
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
   Begin PCGestion.miCombo cbESTADO 
      Height          =   480
      Left            =   5985
      TabIndex        =   6
      Top             =   4335
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   847
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
   Begin PCGestion.miText ioDESCRIPCION 
      Height          =   525
      Left            =   2385
      TabIndex        =   3
      Top             =   3825
      Width           =   5745
      _ExtentX        =   10134
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
   Begin PCGestion.miText ioCODBAR 
      Height          =   525
      Left            =   1020
      TabIndex        =   0
      Top             =   450
      Width           =   3150
      _ExtentX        =   5556
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
   Begin PCGestion.bsGradientLabel bsGradientLabel1 
      Height          =   390
      Left            =   1515
      Top             =   5280
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   688
      Caption         =   "-F8- Aceptar  -F9- Cancelar "
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
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
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CBARRAS"
      Height          =   300
      Left            =   60
      TabIndex        =   14
      Top             =   540
      Width           =   915
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MOTIVO"
      Height          =   300
      Left            =   1035
      TabIndex        =   13
      Top             =   3885
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ESTADO"
      Height          =   300
      Left            =   5040
      TabIndex        =   12
      Top             =   4410
      Width           =   885
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "COSTO"
      Height          =   300
      Left            =   60
      TabIndex        =   11
      Top             =   4395
      Width           =   810
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PVP"
      Height          =   300
      Left            =   2595
      TabIndex        =   10
      Top             =   4395
      Width           =   405
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "COSTURERA"
      Height          =   300
      Left            =   4095
      TabIndex        =   8
      Top             =   540
      Width           =   1275
   End
End
Attribute VB_Name = "frmNuArr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Módulo     : frmNuArr
' Fecha/Hora : 21/01/2004 12:25
' Autor      : JCastillo
' Propósito  :  Nuevo arreglo
'---------------------------------------------------------------------------------------
Option Explicit

'para guardar con el arreglo
Public CODIGO_VENTA As Long

Public mi_Codart As Long
Public mi_Tempor As Byte
Public mi_talla As Integer
Public mi_Color As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'para seleccionar un registro existente ...
'Solo_actualizar = true
Public Solo_Actualizar As Boolean
Public Sel_ID As Long
Public Sel_Caja As Byte
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Codigo_Cost As Long

Private Sub cbAceptar_Click()
    
If ioCOSTO.Text = "" Then ioCOSTO.Text = "0"
If ioPVP.Text = "" Then ioPVP.Text = "0"

If cbESTADO.Text = "" Then
    lblstatus.Caption = "Estado no puede estar en blanco"
    cbESTADO.SetFocus
    Exit Sub
End If

Call añade_Arreglo
DoEvents
Unload Me

End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : añade_Arreglo
' Fecha/Hora    : 21/01/2004 13:26
' Autor         : JCastillo
' Propósito     :  Añade el arreglo a la base de datos
'---------------------------------------------------------------------------------------
Private Sub añade_Arreglo()
Dim tmpid As Variant
Dim rc As New ADODB.Recordset

'ID
'CODCOST
'CODART
'TEMPOR
'CODTALLA
'CODCOL
'CODVEN
'CODCAJ
'DESCRIPCION
'COSTE
'PVP
'ESTADO

   On Error GoTo añade_Arreglo_Error

 With rc
  
 If Solo_Actualizar Then
 
    .Open "SELECT * FROM ARREGLOS WHERE ID = " & Sel_ID & " AND CODCAJ = " & Sel_Caja, locCnn, adOpenDynamic, adLockOptimistic
 
 Else
 
    tmpid = devuelve_campo("SELECT MAX(ID) + 1 FROM ARREGLOS WHERE CODCAJ = " & CajaActual)

    If tmpid = "@" Then tmpid = 1

    .Open "SELECT TOP 1 * FROM ARREGLOS", locCnn, adOpenDynamic, adLockOptimistic
    .AddNew
    .fields("ID") = tmpid
    
    If mi_Codart > 0 Then .fields("CODART") = mi_Codart
    If mi_Tempor > 0 Then .fields("TEMPOR") = mi_Tempor
    If mi_talla > 0 Then .fields("CODTALLA") = mi_talla
    If mi_Color > 0 Then .fields("CODCOL") = mi_Color
    .fields("CODUSR") = UsuarioActual
    .fields("CODVEN") = CODIGO_VENTA
    .fields("CODCAJ") = CajaActual
    
 End If

    If Codigo_Cost > 0 Then .fields("CODCOST") = Codigo_Cost

    If IsNull(.fields("CODART")) Then .fields("CODART") = mi_Codart
    If IsNull(.fields("TEMPOR")) Then .fields("TEMPOR") = mi_Tempor
    If IsNull(.fields("CODTALLA")) Then .fields("CODTALLA") = mi_talla
    If IsNull(.fields("CODCOL")) Then .fields("CODCOL") = mi_Color
    
    .fields("DESCRIPCION") = ioDescripcion.Text
    .fields("COSTE") = ioCOSTO.Text
    .fields("PVP") = ioPVP.Text
    .fields("ESTADO") = cbESTADO.Text
    .Update
    .Close
    
End With

Set rc = Nothing

   On Error GoTo 0
   Exit Sub

añade_Arreglo_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento añade_Arreglo de Formulario frmNuArr"
End Sub

Private Sub cbCancelar_Click()

Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   On Error GoTo Form_KeyDown_Error

Select Case KeyCode

Case vbKeyF8

    Call cbAceptar_Click
    KeyCode = 0
    
Case vbKeyF9

    Call cbCancelar_Click
    KeyCode = 0

End Select

   On Error GoTo 0
   Exit Sub

Form_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento Form_KeyDown de Formulario frmNuArr"

End Sub

Private Sub Form_Load()

Move (Screen.Width - Width) \ 2, Separacion_MDIForm

  With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
  End With
  
With cbESTADO
    .añade_item "1   - PENDIENTE"
    .añade_item "2   - SERVIDO"
    .añade_item "3   - CANCELADO"
    .LenCodigo = 1
    .CodigoWidth = 300
    .Text = "1"
End With

With ioCODBAR
    .LongMaxima = 13
    .SoloNumeros = True
    
    If ((mi_Codart = 0) And (mi_Tempor = 0)) Or ((mi_Codart = 0) And Me.Sel_ID = 0) Then
        .Visible = True
    Else
        .Visible = False
    End If
End With


With ioNOMBRE
    .LongMaxima = 50
End With

With ioCOSTO
    .dspFormat = "Currency"
    .Alineacion = 1
    .LongMaxima = 10
End With

With ioPVP
    .dspFormat = "Currency"
    .Alineacion = 1
    .LongMaxima = 10
End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Solo_Actualizar = False
End Sub



Private Sub ioCODBAR_Validate(Cancel As Boolean)
Dim mic As MiCodBar
Dim campos As String
Dim cadena As String
Dim tart As Variant

        If Len(ioCODBAR.Text) = LenCodBar Then
        
        mic = Descompone_CBAR(ioCODBAR.Text)

        mi_Codart = mic.CODIGO_ART
        mi_Tempor = mic.TEMPORADA_ART
        mi_talla = mic.TALLA_ART
        mi_Color = mic.COLOR_ART

        campos = devuelve_campo("SELECT MODELO FROM MAARTIC WHERE CODIGO = " & mic.CODIGO_ART & " AND TEMPOR = " & mic.TEMPORADA_ART, locCnn)
        
        If campos = "@" Then
            Cancel = True
            lblstatus.Caption = "No se encuentra el artículo en la base de datos"
            Exit Sub
        End If
        
                
        
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
                   
            mi_Codart = tart(1)
            mi_Tempor = TemporadaActual
            mi_talla = "0"
            mi_Color = "0"
            
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


        
             'codigo de artículo
        lblArticulo.Caption = Format(mic.CODIGO_ART, "00000") & " " & campos & " " & Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & mic.TALLA_ART, locCnn)) & "  " & Trim(devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & mic.COLOR_ART, locCnn))
        'obtener el % de iva
        'ioIVA.Text = devuelve_campo("SELECT IVA FROM IVA WHERE CODIGO = " & campos(2), locCnn)
        'la talla
        'lblTalla.Caption = Trim(devuelve_campo("SELECT DESCRIPCION FROM TALLAS WHERE CODIGO = " & Codigo_B.TALLA_ART, locCnn))
        'el color
       ' lblColor.BackColor = devuelve_campo("SELECT CODCOL FROM COLORES WHERE CODIGO = " & Codigo_B.COLOR_ART, locCnn)
        'lblColorDesc.Caption = Trim(devuelve_campo("SELECT DESCRIPCION FROM COLORES WHERE CODIGO = " & Codigo_B.COLOR_ART, locCnn))

        'lblModelo.Caption = Trim(campos(0))
        'ioPREVEN.Text = campos(1)
        
        lblstatus.Caption = ""

End Sub

'---------------------------------------------------------------------------------------
' Procedimiento : ioNOMBRE_Validate
' Fecha/Hora    : 21/01/2004 12:50
' Autor         : JCastillo
' Propósito     :   Rellenar el List con las costureras
'---------------------------------------------------------------------------------------
Private Sub ioNOMBRE_Validate(Cancel As Boolean)

Dim rc_Cost As New ADODB.Recordset

   On Error GoTo ioNOMBRE_Validate_Error

lstCostureras.Clear

If ioNOMBRE.Text = "" Then
    Codigo_Cost = 0
    ioDescripcion.SetFocus
    Exit Sub
End If

With rc_Cost

      .Open "SELECT CODIGO, NOMBRE, DIRECCION from COSTURE WHERE MBAJA = 0 AND NOMBRE LIKE '%" & ioNOMBRE.Text & "%' ORDER BY NOMBRE", locCnn, adOpenDynamic, adLockReadOnly

'lstCostureras.Clear

Do Until .EOF

lstCostureras.AddItem Format(.fields("CODIGO"), "0000000") & "   " & .fields("NOMBRE") & "   " & .fields("DIRECCION")
 If Not .EOF Then .MoveNext
     
Loop

.Close
Set rc_Cost = Nothing

If lstCostureras.ListCount > 0 Then
    lstCostureras.ListIndex = 0
    lstCostureras.SetFocus
End If

DoEvents

End With

   On Error GoTo 0
   Exit Sub

ioNOMBRE_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento ioNOMBRE_Validate de Formulario frmNuArr"

End Sub


Private Sub lstCostureras_DblClick()

    If lstCostureras.Text <> "" Then
        Codigo_Cost = Left(lstCostureras.Text, 7)
    Else
        Codigo_Cost = 0
    End If
    
    ioDescripcion.SetFocus

End Sub

Private Sub lstCostureras_GotFocus()

'si no hay registros pasar al siguiente control
If lstCostureras.ListCount = 0 Then
     ioDescripcion.SetFocus
End If

End Sub

Private Sub lstCostureras_KeyPress(KeyAscii As Integer)

'si pulsa enter pasar al siguiente control
If KeyAscii = 13 Then
    If lstCostureras.Text <> "" Then
        Codigo_Cost = Left(lstCostureras.Text, 7)
    Else
        Codigo_Cost = 0
    End If
    
     ioDescripcion.SetFocus
End If

End Sub
