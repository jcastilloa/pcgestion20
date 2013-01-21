VERSION 5.00
Begin VB.Form frmOptArreglo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seleccione una opción"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3330
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
   ScaleHeight     =   2670
   ScaleWidth      =   3330
   ShowInTaskbar   =   0   'False
   Begin PCGestion.chameleonButton cbNuevoArr 
      Height          =   795
      Left            =   600
      TabIndex        =   0
      Top             =   75
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "&Nuevo Arreglo"
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
      BCOLO           =   0
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmOptArreglo.frx":0000
      PICN            =   "frmOptArreglo.frx":001C
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
      Left            =   600
      TabIndex        =   2
      Top             =   1755
      Width           =   2130
      _ExtentX        =   3757
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
      BCOLO           =   0
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmOptArreglo.frx":0CF6
      PICN            =   "frmOptArreglo.frx":0D12
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin PCGestion.chameleonButton cbSeleccionarArr 
      Height          =   795
      Left            =   600
      TabIndex        =   1
      Top             =   915
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   1402
      BTYPE           =   9
      TX              =   "&Seleccionar Existente"
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
      BCOLO           =   0
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmOptArreglo.frx":15EC
      PICN            =   "frmOptArreglo.frx":1608
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmOptArreglo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Módulo     : frmOptArreglo
' Fecha/Hora : 23/01/2004 11:59
' Autor      : JCastillo
' Propósito  : Un simple formulario para seleccionar si desea añadir un nuevo arreglo
'                  o agregar a la venta un arreglo existente
'---------------------------------------------------------------------------------------
Option Explicit

'opcion = 0   (cancelar)
'opcion = 1   (nuevo arreglo)
'opcion = 2   (buscar arreglo)
Public Opcion As Byte

Private Sub cbNuevoArr_Click()

Opcion = 1
Unload Me

End Sub

Private Sub cbCancelar_Click()

Opcion = 0
Unload Me

End Sub

Private Sub cbSeleccionarArr_Click()

Opcion = 2
Unload Me

End Sub

Private Sub Form_Load()

Move (Screen.Width - Width) \ 2, Separacion_MDIForm

End Sub

