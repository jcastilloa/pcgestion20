VERSION 5.00
Object = "*\ABWord\miBwordOCX.vbp"
Begin VB.Form frmTMPInicio 
   BackColor       =   &H00FCAA8B&
   Caption         =   "Inicio Temporal"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin miBwordOCX.miBword Editor 
      Left            =   5325
      Top             =   2670
      _ExtentX        =   2593
      _ExtentY        =   1508
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
   Begin PCGestion.ucGrdBttn ucGrdBttn9 
      Height          =   525
      Left            =   60
      TabIndex        =   15
      Top             =   3480
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   926
      Caption         =   "Centros"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmTMPInicio.frx":0000
   End
   Begin PCGestion.ucGrdBttn ucGrdBttn2 
      Height          =   645
      Left            =   3165
      TabIndex        =   1
      Top             =   2250
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   1138
      Caption         =   "Pedidos Proveedores"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmTMPInicio.frx":001C
   End
   Begin PCGestion.ucGrdBttn ucGrdBttn1 
      Height          =   1590
      Left            =   6630
      TabIndex        =   0
      Top             =   330
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   2805
      Caption         =   "GrdBttn1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmTMPInicio.frx":0038
   End
   Begin PCGestion.ucGrdBttn FWCoolButton1 
      Height          =   510
      Left            =   60
      TabIndex        =   2
      Top             =   2400
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   900
      Caption         =   "Formas Pago"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmTMPInicio.frx":75BA
   End
   Begin PCGestion.ucGrdBttn FWCoolButton3 
      Height          =   525
      Left            =   60
      TabIndex        =   3
      Top             =   2940
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   926
      Caption         =   "Sectores"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmTMPInicio.frx":75D6
   End
   Begin PCGestion.ucGrdBttn cbprov 
      Height          =   555
      Left            =   60
      TabIndex        =   4
      Top             =   90
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   979
      Caption         =   "Proveedores"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmTMPInicio.frx":75F2
   End
   Begin PCGestion.ucGrdBttn cbSec 
      Height          =   555
      Left            =   60
      TabIndex        =   5
      Top             =   660
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   979
      Caption         =   "Secciones"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmTMPInicio.frx":760E
   End
   Begin PCGestion.ucGrdBttn cbFam 
      Height          =   555
      Left            =   60
      TabIndex        =   6
      Top             =   1245
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   979
      Caption         =   "Familias"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmTMPInicio.frx":762A
   End
   Begin PCGestion.ucGrdBttn cbSfm 
      Height          =   555
      Left            =   60
      TabIndex        =   7
      Top             =   1830
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   979
      Caption         =   "SubFamilias"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmTMPInicio.frx":7646
   End
   Begin PCGestion.ucGrdBttn ucGrdBttn3 
      Height          =   555
      Left            =   1545
      TabIndex        =   8
      Top             =   1245
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   979
      Caption         =   "Articulos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmTMPInicio.frx":7662
   End
   Begin PCGestion.ucGrdBttn ucGrdBttn4 
      Height          =   555
      Left            =   1530
      TabIndex        =   9
      Top             =   1830
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   979
      Caption         =   "Formas Cobro"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmTMPInicio.frx":767E
   End
   Begin PCGestion.ucGrdBttn ucGrdBttn5 
      Height          =   555
      Left            =   1545
      TabIndex        =   10
      Top             =   90
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   979
      Caption         =   "IVA"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmTMPInicio.frx":769A
   End
   Begin PCGestion.ucGrdBttn ucGrdBttn6 
      Height          =   555
      Left            =   1545
      TabIndex        =   11
      Top             =   660
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   979
      Caption         =   "Temporadas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmTMPInicio.frx":76B6
   End
   Begin PCGestion.ucGrdBttn ucGrdBttn7 
      Height          =   510
      Left            =   1530
      TabIndex        =   12
      Top             =   2400
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   900
      Caption         =   "Empresas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmTMPInicio.frx":76D2
   End
   Begin PCGestion.ucGrdBttn kk 
      Height          =   525
      Left            =   1530
      TabIndex        =   13
      Top             =   2940
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   926
      Caption         =   "Representantes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmTMPInicio.frx":76EE
   End
   Begin PCGestion.ucGrdBttn ucGrdBttn8 
      Height          =   525
      Left            =   1530
      TabIndex        =   14
      Top             =   3480
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   926
      Caption         =   "Dependientes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmTMPInicio.frx":770A
   End
   Begin PCGestion.ucGrdBttn ucGrdBttn 
      Height          =   525
      Left            =   3390
      TabIndex        =   16
      Top             =   2940
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   926
      Caption         =   "Clientes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmTMPInicio.frx":7726
   End
   Begin PCGestion.ucGrdBttn ucGrdBttn10 
      Height          =   525
      Left            =   3405
      TabIndex        =   17
      Top             =   3495
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   926
      Caption         =   "Bancos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmTMPInicio.frx":7742
   End
   Begin PCGestion.ucGrdBttn ucGrdBttn9kl 
      Height          =   525
      Left            =   75
      TabIndex        =   18
      Top             =   4020
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   926
      Caption         =   "Tallas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmTMPInicio.frx":775E
   End
   Begin PCGestion.ucGrdBttn ucGrdBttn11 
      Height          =   525
      Left            =   1530
      TabIndex        =   19
      Top             =   4020
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   926
      Caption         =   "Colores"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmTMPInicio.frx":777A
   End
   Begin PCGestion.ucGrdBttn ucGrdBttn12 
      Height          =   510
      Left            =   3405
      TabIndex        =   20
      Top             =   4035
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   900
      Caption         =   "Almacenes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmTMPInicio.frx":7796
   End
   Begin PCGestion.ucGrdBttn ucGrdBttn13 
      Height          =   510
      Left            =   5190
      TabIndex        =   21
      Top             =   4050
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   900
      Caption         =   "Cajas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmTMPInicio.frx":77B2
   End
   Begin PCGestion.ucGrdBttn ucGrdBttn14 
      Height          =   510
      Left            =   5190
      TabIndex        =   22
      Top             =   3495
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   900
      Caption         =   "Tarifas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmTMPInicio.frx":77CE
   End
   Begin PCGestion.ucGrdBttn P 
      Height          =   510
      Left            =   5250
      TabIndex        =   23
      Top             =   2085
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   900
      Caption         =   "Transferencias"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmTMPInicio.frx":77EA
   End
   Begin PCGestion.ucGrdBttn cmdInicioSesion 
      Height          =   510
      Left            =   3165
      TabIndex        =   25
      Top             =   1710
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   900
      Caption         =   "Inicio de Sesión"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "FrmTMPInicio.frx":7806
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "INICIO TEMPORAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   3510
      TabIndex        =   24
      Top             =   150
      Width           =   2760
   End
End
Attribute VB_Name = "FrmTMPInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : FrmTMPInicio
' DateTime  : 04/11/2003 18:22
' Author    : Administrador
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Dim cambio_Sesion As Boolean

Private Sub cbProv_click()

FrmMntProv.Show

End Sub

Private Sub cbfam_click()

FrmMntFam.Show

End Sub

Private Sub cbsec_click()

FrmMntSec.Show

End Sub

Private Sub cbSfm_click()

FrmMntSfm.Show

End Sub

Private Sub cmdInicioSesion_Click()
cambio_Sesion = True
frmInicSesion.Show
Unload Me
End Sub

Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
Me.Caption = titulo
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


If cambio_Sesion Then
    cambio_Sesion = False
    Else
    Call cerrar_conexiones
    End
End If

End Sub

Private Sub FWCoolButton1_click()

FrmMntFpa.Show

End Sub

Private Sub FWCoolButton3_Click()
frmMntSct.Show
End Sub

Private Sub kk_Click()
frmMntRep.Show
End Sub

Private Sub P_Click()
frmPtrans.Show
End Sub

Private Sub ucGrdBttn_Click()
frmMntCli.Show
End Sub

Private Sub ucGrdBttn1_Click()
Plantilla.Show
End Sub

Private Sub ucGrdBttn10_Click()
frmMntBan.Show
End Sub

Private Sub ucGrdBttn11_Click()
frmMntCol.Show
End Sub

Private Sub ucGrdBttn12_Click()

frmMntAlm.Show
End Sub

Private Sub ucGrdBttn13_Click()
frmMntCaj.Show
End Sub

Private Sub ucGrdBttn14_Click()
frmMntTar.Show
End Sub

Private Sub ucGrdBttn2_Click()
frmPedProv.Show
End Sub

Private Sub ucGrdBttn3_Click()
FrmMntArt.Show
End Sub

Private Sub ucGrdBttn4_Click()
frmMntFco.Show
End Sub

Private Sub ucGrdBttn5_Click()
frmMntIva.Show
End Sub

Private Sub ucGrdBttn6_Click()
frmMntTem.Show
End Sub

Private Sub ucGrdBttn7_Click()
frmMntEmp.Show
End Sub

Private Sub ucGrdBttn8_Click()
frmMntPer.Show
End Sub


Private Sub ucGrdBttn9_Click()
frmMntCen.Show
End Sub

Private Sub ucGrdBttn9kl_Click()
frmMntTall.Show
End Sub
