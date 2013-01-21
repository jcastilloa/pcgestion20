VERSION 5.00
Begin VB.Form frmCambioPwd 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambio de contraseña de usuario"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6165
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin PCGestion.miText ioCNUEVA 
      Height          =   465
      Left            =   2715
      TabIndex        =   1
      Top             =   1065
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   820
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
      Left            =   2130
      TabIndex        =   3
      Top             =   2640
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
      MICON           =   "frmCambioPwd.frx":0000
      PICN            =   "frmCambioPwd.frx":001C
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
      Left            =   3105
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2640
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
      MICON           =   "frmCambioPwd.frx":0CF6
      PICN            =   "frmCambioPwd.frx":0D12
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.miText ioCANTERIOR 
      Height          =   525
      Left            =   2715
      TabIndex        =   0
      Top             =   480
      Width           =   2130
      _ExtentX        =   3757
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
   Begin PCGestion.miText ioREPITAN 
      Height          =   480
      Left            =   2715
      TabIndex        =   2
      Top             =   1635
      Width           =   2130
      _ExtentX        =   3757
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
      dspFormat       =   ""
      Enabled         =   -1  'True
      EsPassword      =   -1  'True
   End
   Begin PCGestion.bsGradientLabel lblUsuario 
      Height          =   375
      Left            =   0
      Top             =   15
      Width           =   6165
      _ExtentX        =   10874
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
   Begin PCGestion.bsGradientLabel lblStatus 
      Height          =   375
      Left            =   0
      Top             =   2190
      Width           =   6165
      _ExtentX        =   10874
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Repita Nueva"
      Height          =   345
      Left            =   1320
      TabIndex        =   7
      Top             =   1695
      Width           =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña anterior"
      Height          =   585
      Left            =   1410
      TabIndex        =   6
      Top             =   375
      Width           =   1260
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nueva Contraseña"
      Height          =   615
      Left            =   1320
      TabIndex        =   5
      Top             =   1005
      Width           =   1350
   End
End
Attribute VB_Name = "frmCambioPwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Modulo      : frmCambioPwd
' Fecha/Hora  : 29/12/2003 20:48
' Autor       : JCASTILLO
' Propósito   : Cambio de contraseña de usuario
'---------------------------------------------------------------------------------------
Option Explicit

'---------------------------------------------------------------------------------------
'Para cambiar la contraseña de otro usuario distinto al usuario actual
'(variable UsuarioActual). Poner otro_usuario = TRUE y darle el codigo
'del usuario al que se desea cambiar la clave
Public otro_usuario As Boolean
Public ID_Otro_usuario As Single
'---------------------------------------------------------------------------------------

Dim tmp_ID As Single

Private Sub cbAceptar_Click()
 
'validaciones ...
If Trim(ioCANTERIOR.Text) = "" Then
    lblStatus.Caption = "La contraseña ANTERIOR no puede estar en blanco"
    Exit Sub
End If

If Trim(ioCNUEVA.Text) = "" Then
    lblStatus.Caption = "La contraseña NUEVA no puede estar en blanco"
    Exit Sub
End If
 
'si la contraseña anterior no es igual
If Trim(ioCANTERIOR.Text) <> Trim(devuelve_campo("SELECT CLAVE FROM PERSONAL WHERE CODIGO = " & tmp_ID, locCnn)) Then
 
    lblStatus.Caption = "¡La Contraseña ANTERIOR es incorrecta!"
    
    ioCANTERIOR.SetFocus
    ioCANTERIOR.CancelarValidacion
    Exit Sub
 
End If

'si no coinciden las 2 contraseñas nuevas ...
If (ioCNUEVA.Text) <> (ioREPITAN.Text) Then

    lblStatus.Caption = "¡Las Contraseñas NO coinciden!"
        
    ioREPITAN.SetFocus
    ioREPITAN.CancelarValidacion
    Exit Sub
    
End If

'si esta todo OK, cambiar la contraseña ...
locCnn.Execute "UPDATE PERSONAL SET CLAVE = '" & ioCNUEVA.Text & "' WHERE CODIGO = " & tmp_ID
DoEvents

lblStatus.Caption = ""

Unload Me

End Sub

Private Sub cbCancelar_Click()

Unload Me

End Sub

Private Sub Form_Load()

Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2

If otro_usuario Then
    tmp_ID = ID_Otro_usuario
Else
    tmp_ID = UsuarioActual
End If
    
lblUsuario.Caption = devuelve_campo("SELECT NOMBRE FROM PERSONAL WHERE CODIGO = " & tmp_ID, locCnn)

With ioCNUEVA
    .PermitirBlanco = False
    .EsPassword = True
End With

With ioCANTERIOR
    .PermitirBlanco = False
    .EsPassword = True
End With

With ioREPITAN
    .PermitirBlanco = False
    .EsPassword = True
End With

End Sub
