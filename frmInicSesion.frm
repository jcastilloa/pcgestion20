VERSION 5.00
Begin VB.Form frmInicSesion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inicio de Sesión"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   6510
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin PCGestion.miCombo cbUSUARIO 
      Height          =   480
      Left            =   1335
      TabIndex        =   0
      Top             =   600
      Width           =   5010
      _ExtentX        =   8837
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
   Begin PCGestion.chameleonButton cbCancelar 
      Height          =   825
      Left            =   3300
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2130
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1455
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
      MICON           =   "frmInicSesion.frx":0000
      PICN            =   "frmInicSesion.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbAceptar 
      Height          =   825
      Left            =   2010
      TabIndex        =   2
      Top             =   2130
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1455
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
      MICON           =   "frmInicSesion.frx":08F6
      PICN            =   "frmInicSesion.frx":0912
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.bsGradientLabel lblCentro 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   714
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
      Colour1         =   16744576
      Colour2         =   16761024
      CaptionAlignment=   1
   End
   Begin PCGestion.bsGradientLabel lblstatus 
      Height          =   375
      Left            =   0
      Top             =   1650
      Width           =   6525
      _ExtentX        =   11509
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
   Begin PCGestion.miText ioCLAVE 
      Height          =   480
      Left            =   1350
      TabIndex        =   1
      Top             =   1110
      Width           =   1680
      _ExtentX        =   2963
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
      dspFormat       =   ""
      Enabled         =   -1  'True
      EsPassword      =   -1  'True
   End
   Begin PCGestion.chameleonButton cbCambioPWD 
      Height          =   450
      Left            =   4305
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1125
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   794
      BTYPE           =   9
      TX              =   "F1 Cambiar Clave"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
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
      MICON           =   "frmInicSesion.frx":11EC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CLAVE"
      Height          =   285
      Left            =   375
      TabIndex        =   4
      Top             =   1170
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "USUARIO"
      Height          =   285
      Left            =   375
      TabIndex        =   3
      Top             =   675
      Width           =   930
   End
End
Attribute VB_Name = "frmInicSesion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Modulo      : frmInicSesion
' Fecha/Hora  : 19/11/2003 21:59
' Autor       : JCASTILLO
' Propósito   : Inicio de sesión en la aplicación. Autentifica un usuario
'               de la empresa seleccionada y obtiene PERMISO, ALMACEN y CAJA (si es un
'               dependiente)
'---------------------------------------------------------------------------------------
Option Explicit

Private Sub cbAceptar_Click()
Dim t_centro As Variant
'
'Pre validaciones
If cbUSUARIO.Text <> "" And Trim(ioCLAVE.Text) <> "" Then
    
    'comprobar la clave de usuario
    If Trim(devuelve_campo("SELECT CLAVE FROM PERSONAL WHERE CODIGO = " & cbUSUARIO.Text, locCnn)) = Trim(ioCLAVE.Text) Then
    
        'cargar valores por defecto de esta sesión
        UsuarioActual = cbUSUARIO.Text
        'CajaActual = devuelve_campo("SELECT CODCAJA FROM PERSONAL WHERE CODIGO = " & cbUSUARIO.Text, locCnn)
        'AlmacenActual = devuelve_campo("SELECT CODALM FROM CAJAS WHERE CODIGO = " & CajaActual, locCnn)
        TipoPermiso = devuelve_campo("SELECT TIPPERM FROM PERSONAL WHERE CODIGO = " & cbUSUARIO.Text, locCnn)
       ' CentroActual = devuelve_campo("Select CODCEN from PERSONAL WHERE CODIGO = " & CLng(cbUSUARIO.Text), locCnn)
        
        lblstatus.Caption = ""
        DoEvents
        
        Call escribe_txt_config(False)
        
        FrmInicio.Show
                
        
        'Solo para MS SQL
        If TipoServer = 1 Then
        
        t_centro = devuelve_matriz("SELECT ACTIP, INTERVALO FROM CENTROS WHERE CODIGO = " & CentroActual, locCnn)
        'comprobar si se debe activar el timer

        'If t_centro(0) <> 0 Then
            'activar el timer si es necesario
            'Intervalo: 1000 * 1800 (una vez cada media hora)
         '   UsarTimerIP = SetTimer(&H0, &H0, 1000 * t_centro(1), AddressOf TimerProc)
        'End If
        
        End If
        
        'FrmTMPInicio.Show
        Unload Me
        Exit Sub
    
    Else
    
       lblstatus.Caption = "¡CLAVE INCORRECTA!"
       ioCLAVE.SetFocus
       Exit Sub
        
    End If
    
ElseIf cbUSUARIO.Text = "" Then
    lblstatus.Caption = "USUARIO no puede estar en blanco"
    cbUSUARIO.SetFocus
ElseIf Trim(ioCLAVE.Text) = "" Then
    lblstatus.Caption = "CLAVE no puede estar en blanco"
    ioCLAVE.SetFocus
End If


 

End Sub

Private Sub cbCancelar_Click()

Call cerrar_conexiones
End

End Sub

Private Sub cbCambioPWD_Click()

If Trim(cbUSUARIO.Text) = "" Then
    lblstatus.Caption = "Seleccione usuario"
    cbUSUARIO.SetFocus
    Exit Sub
End If

    lblstatus.Caption = ""
    
With frmCambioPwd
    .otro_usuario = True
    .ID_Otro_usuario = cbUSUARIO.Text
    .Show 1
    ioCLAVE.SetFocus
End With

End Sub



'---------------------------------------------------------------------------------------
' Subrutina   : cbUSUARIO_Validate
' Fecha/Hora  : 11/01/2004 14:09
' Autor       : JCASTILLO
' Propósito   :
'
'---------------------------------------------------------------------------------------
Private Sub cbUSUARIO_Validate(Cancel As Boolean)
Dim tmpcen As Byte

   On Error GoTo cbUSUARIO_Validate_Error

If cbUSUARIO.Text = "" Then
    tmpcen = 0
    Exit Sub
End If

tmpcen = devuelve_campo("Select CODCEN from PERSONAL WHERE CODIGO = " & CLng(cbUSUARIO.Text), locCnn)
lblCentro.Caption = "Centro: " & Trim(devuelve_campo("Select DESCRIPCION from CENTROS WHERE CODIGO = " & tmpcen, locCnn))

   On Error GoTo 0
   Exit Sub

cbUSUARIO_Validate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en procedimiento cbUSUARIO_Validate de Formulario frmInicSesion"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode

    'si pulsa F1, cambiar clave
    Case vbKeyF1
        Call cbCambioPWD_Click

End Select

End Sub


Private Sub Form_Load()


   Move (Screen.Width - Width) \ 2, Separacion_MDIForm

   With locCnn
    If .State = 0 Then
        .CursorLocation = adUseClient
        .Open strLocCnn
    End If
   End With

  'Usuarios
  With cbUSUARIO
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, NOMBRE FROM PERSONAL WHERE ((CODCEN = " & CentroActual & ") OR (TIPPERM = 1)) AND (MBAJA = 0) ORDER BY CODIGO"
    .LenCodigo = 5
    .CodigoWidth = 700
    .carga
  End With
  
  With ioCLAVE
    .LongMaxima = 10
    .EsPassword = True
  End With
  
  lblCentro.Caption = "Centro: " & Trim(devuelve_campo("Select DESCRIPCION from CENTROS WHERE CODIGO = " & CentroActual, locCnn))

End Sub

'Descargar de memoria
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   With locCnn
    If .State <> 0 Then .Close
   End With
  
Set frmInicSesion = Nothing

'If UsarTimerIP > 0 Then KillTimer &H0, UsarTimerIP

End Sub

Private Sub ioCLAVE_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then SendKeys "{tab}"

End Sub
