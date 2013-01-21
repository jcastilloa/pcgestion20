VERSION 5.00
Begin VB.Form frmBusTrn 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Buscar Transferencia"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6375
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
   ScaleHeight     =   2505
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   Begin PCGestion.miText ioCODTRN 
      Height          =   525
      Left            =   1260
      TabIndex        =   1
      Top             =   615
      Width           =   1710
      _ExtentX        =   3016
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
      Left            =   2235
      TabIndex        =   2
      Top             =   1680
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
      MICON           =   "frmBusTrn.frx":0000
      PICN            =   "frmBusTrn.frx":001C
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
      Left            =   3210
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1680
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
      MICON           =   "frmBusTrn.frx":0CF6
      PICN            =   "frmBusTrn.frx":0D12
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.miCombo cbCODALMORIG 
      Height          =   510
      Left            =   1245
      TabIndex        =   0
      Top             =   105
      Width           =   4740
      _ExtentX        =   8361
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
   Begin PCGestion.bsGradientLabel lblStatus 
      Height          =   375
      Left            =   30
      Top             =   1245
      Width           =   6300
      _ExtentX        =   11113
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
      Colour1         =   12632256
      Colour2         =   16558731
      CaptionAlignment=   1
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   435
      TabIndex        =   5
      Top             =   690
      Width           =   765
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Origen"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   480
      TabIndex        =   4
      Top             =   165
      Width           =   750
   End
End
Attribute VB_Name = "frmBusTrn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Módulo       : frmBusTrn
' Fecha/Hora : 22/03/2004 12:33
' Autor         : JCastillo
' Propósito   :
'---------------------------------------------------------------------------------------

Option Explicit

Public miCODALM_ORIG As Byte
Public miCODTRN As Long

Private Sub cbAceptar_Click()

If cbCODALMORIG.Text = "" Then
    lblStatus.Caption = "Almacén de origen no puede estar en blanco"
    Exit Sub
End If

If ioCODTRN.Text = "" Then
    lblStatus.Caption = "Codigo de transferencia no puede estar en blanco"
    Exit Sub
End If

miCODALM_ORIG = cbCODALMORIG.Text
miCODTRN = ioCODTRN.Text

Unload Me

End Sub

Private Sub cbCancelar_Click()

miCODALM_ORIG = 0
miCODTRN = 0

Unload Me

End Sub

Private Sub Form_Load()

Move (Screen.Width - Width) \ 2, Separacion_MDIForm

With ioCODTRN
    .SoloNumeros = True
    .Alineacion = 1
End With

With cbCODALMORIG
    .ConexionString = locCnn
    .LenCodigo = 3
    .SQLString = "SELECT CODIGO, DESCRIPCION FROM ALMACENES WHERE MBAJA = 0 ORDER BY CODIGO"
    .CodigoWidth = 500
    .carga
    .Text = "1"
End With

End Sub

