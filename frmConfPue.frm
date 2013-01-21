VERSION 5.00
Begin VB.Form frmConfPue 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración del Puesto"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   7200
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
   ScaleHeight     =   4020
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin PCGestion.bsGradientLabel lblTot_sin_IVA 
      Height          =   465
      Left            =   525
      Top             =   60
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   820
      Caption         =   "Configuración de Valores por Defecto"
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   4210752
      Colour1         =   14737632
      Colour2         =   12632256
      CaptionAlignment=   1
   End
   Begin PCGestion.chameleonButton cbAceptar 
      Height          =   795
      Left            =   2655
      TabIndex        =   4
      Top             =   3165
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
      MICON           =   "frmConfPue.frx":0000
      PICN            =   "frmConfPue.frx":001C
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
      Left            =   3630
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3165
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
      MICON           =   "frmConfPue.frx":0CF6
      PICN            =   "frmConfPue.frx":0D12
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.miCombo cbEMPRESAS 
      Height          =   510
      Left            =   1605
      TabIndex        =   0
      Top             =   630
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   900
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
   Begin PCGestion.miCombo cbALMACENES 
      Height          =   510
      Left            =   1605
      TabIndex        =   2
      Top             =   1635
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   900
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
   Begin PCGestion.miCombo cbCAJAS 
      Height          =   510
      Left            =   1605
      TabIndex        =   3
      Top             =   2145
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   900
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
   Begin PCGestion.miCombo cbCENTROS 
      Height          =   510
      Left            =   1605
      TabIndex        =   1
      Top             =   1140
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   900
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
   Begin PCGestion.bsGradientLabel lblstatus 
      Height          =   375
      Left            =   15
      Top             =   2700
      Width           =   7170
      _ExtentX        =   12647
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
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CENTRO"
      Height          =   330
      Left            =   570
      TabIndex        =   9
      Top             =   1230
      Width           =   990
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CAJA"
      Height          =   330
      Left            =   555
      TabIndex        =   8
      Top             =   2220
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ALMACEN"
      Height          =   330
      Left            =   570
      TabIndex        =   7
      Top             =   1710
      Width           =   990
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "EMPRESA"
      Height          =   330
      Left            =   540
      TabIndex        =   6
      Top             =   720
      Width           =   1020
   End
End
Attribute VB_Name = "frmConfPue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Módulo       : frmConfPue
' Fecha/Hora : 16/02/2004 11:49
' Autor         : JCastillo
' Propósito    : Formulario de Configuración del Puesto
'---------------------------------------------------------------------------------------
Option Explicit

Dim tmpconn As ADODB.Connection

Private Sub cbAceptar_Click()

If cbEMPRESAS.Text = "" Then
    lblstatus.Caption = "EMPRESA no puede estar en Blanco"
    Exit Sub
End If

If cbCENTROS.Text = "" Then
    lblstatus.Caption = "CENTRO no puede estar en Blanco"
    Exit Sub
End If

If cbALMACENES.Text = "" Then
    lblstatus.Caption = "ALMACEN no puede estar en Blanco"
    Exit Sub
End If

If cbCAJAS.Text = "" Then
    lblstatus.Caption = "CAJA no puede estar en Blanco"
    Exit Sub
End If

If MsgBox("Atención: Va a cambiar la configuración del puesto. ¿ESTA USTED SEGURO?", vbQuestion + vbYesNo, titulo) = vbNo Then Exit Sub

'hacer el update ...
tmpconn.Execute "UPDATE PUESTCNF SET CODEMP = " & cbEMPRESAS.Text & ", CODALM = " & cbALMACENES.Text & ", CODCAJA= " & cbCAJAS.Text & ", CODCEN =" & cbCENTROS.Text


DoEvents

tmpconn.Close
Set tmpconn = Nothing

Call leer_configuracion

Call escribe_txt_config(True)

Unload Me

End Sub

Private Sub cbCancelar_Click()

'salir sin cambiar ...
Unload Me

End Sub


Private Sub cbCENTROS_Validate(Cancel As Boolean)

If cbCENTROS.Text <> "" Then

    With cbALMACENES
        .borra_combo
        .ConexionString = locCnn
        .SQLString = "SELECT CODIGO, DESCRIPCION FROM ALMACENES WHERE MBAJA = 0 AND CODCEN = " & cbCENTROS.Text & " ORDER BY CODIGO"
        .LenCodigo = 2
        .CodigoWidth = 500
        .carga
    End With

End If

End Sub

Private Sub cbalmacenes_Validate(Cancel As Boolean)

If cbALMACENES.Text <> "" Then

    With cbCAJAS
        .borra_combo
        .ConexionString = locCnn
        .SQLString = "SELECT CODIGO, DESCRIPCION FROM CAJAS WHERE MBAJA = 0 AND CODALM = " & cbALMACENES.Text & " ORDER BY CODIGO"
        .LenCodigo = 2
        .CodigoWidth = 500
        .carga
    End With

End If

End Sub

Private Sub Form_Load()

Dim T_config As Variant

Move (Screen.Width - Width) \ 2, Separacion_MDIForm
'presentar datos

With cbALMACENES
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM ALMACENES WHERE MBAJA = 0 ORDER BY CODIGO"
    .LenCodigo = 2
    .CodigoWidth = 500
    .carga
End With

With cbCAJAS
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM CAJAS WHERE MBAJA = 0 ORDER BY CODIGO"
    .LenCodigo = 2
    .CodigoWidth = 500
    .carga
End With

With cbCENTROS
    .ConexionString = locCnn
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM CENTROS WHERE MBAJA = 0 ORDER BY CODIGO"
    .LenCodigo = 2
    .CodigoWidth = 500
    .carga
End With

Set tmpconn = New ADODB.Connection

tmpconn.Open strEmpCnn

With cbEMPRESAS
    .ConexionString = tmpconn
    .SQLString = "SELECT ID, RAZO FROM EMPRESAS ORDER BY ID"
    .LenCodigo = 1
    .CodigoWidth = 300
    .carga
End With

T_config = devuelve_matriz("SELECT CODEMP, CODALM, CODCAJA, CODCEN FROM PUESTCNF", tmpconn)

'poner valores por defecto
If IsArray(T_config) Then

    cbEMPRESAS.Text = T_config(0)
    cbCENTROS.Text = T_config(3)
    cbALMACENES.Text = T_config(1)
    cbCAJAS.Text = T_config(2)

End If

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'tmpconn.Close
'Set tmpconn = Nothing

Set frmConfPue = Nothing

End Sub
