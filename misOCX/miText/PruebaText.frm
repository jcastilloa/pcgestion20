VERSION 5.00
Object = "{72025988-D2BC-40FC-A5D2-F76373FE1B56}#83.0#0"; "miText.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   75
      TabIndex        =   14
      Top             =   690
      Width           =   510
   End
   Begin PCGestion.miText miText1 
      Height          =   495
      Left            =   630
      TabIndex        =   1
      Top             =   1140
      Width           =   2805
      _extentx        =   4948
      _extenty        =   873
      font            =   "PruebaText.frx":0000
      dspformat       =   ""
      longmaxima      =   10
   End
   Begin PCGestion.miText miText2 
      Height          =   495
      Left            =   630
      TabIndex        =   2
      Top             =   1650
      Width           =   2790
      _extentx        =   4921
      _extenty        =   873
      font            =   "PruebaText.frx":0028
      dspformat       =   ""
      longmaxima      =   10
      alineacion      =   1
      permitirblanco  =   0   'False
   End
   Begin PCGestion.miText miText3 
      Height          =   495
      Left            =   630
      TabIndex        =   3
      Top             =   2190
      Width           =   2805
      _extentx        =   4948
      _extenty        =   873
      font            =   "PruebaText.frx":0054
      dspformat       =   ""
      longmaxima      =   10
      alineacion      =   2
      permitirblanco  =   0   'False
   End
   Begin PCGestion.miText miText4 
      Height          =   480
      Left            =   615
      TabIndex        =   0
      Top             =   645
      Width           =   2820
      _extentx        =   4974
      _extenty        =   847
      font            =   "PruebaText.frx":007C
      dspformat       =   ""
      longmaxima      =   10
   End
   Begin PCGestion.miText miText5 
      Height          =   495
      Left            =   630
      TabIndex        =   4
      Top             =   2685
      Width           =   2805
      _extentx        =   4948
      _extenty        =   873
      font            =   "PruebaText.frx":00A4
      dspformat       =   ""
      longmaxima      =   50
   End
   Begin PCGestion.miText miText6 
      Height          =   495
      Left            =   630
      TabIndex        =   11
      Top             =   3270
      Width           =   2805
      _extentx        =   4948
      _extenty        =   873
      font            =   "PruebaText.frx":00CC
      dspformat       =   ""
      longmaxima      =   50
      permitirblanco  =   0   'False
   End
   Begin VB.Line Line1 
      X1              =   3270
      X2              =   225
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label8 
      Caption         =   "Campo Requerido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4545
      TabIndex        =   13
      Top             =   135
      Width           =   2025
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H006D8639&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   4185
      Shape           =   3  'Circle
      Top             =   150
      Width           =   390
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H006D8639&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   165
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   390
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H006D8639&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   165
      Shape           =   3  'Circle
      Top             =   2295
      Width           =   390
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H006D8639&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   165
      Shape           =   3  'Circle
      Top             =   1740
      Width           =   390
   End
   Begin VB.Label Label7 
      Caption         =   "dd/mm/yyyy. Validación externa <> 01/01/2000"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   3525
      TabIndex        =   12
      Top             =   3345
      Width           =   2805
   End
   Begin VB.Label Label6 
      Caption         =   "Sin Formato"
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
      Left            =   3540
      TabIndex        =   10
      Top             =   2775
      Width           =   1665
   End
   Begin VB.Label Label5 
      Caption         =   "00000000"
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
      Left            =   3555
      TabIndex        =   9
      Top             =   1755
      Width           =   1665
   End
   Begin VB.Label Label4 
      Caption         =   "Currency"
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
      Left            =   3555
      TabIndex        =   8
      Top             =   2280
      Width           =   1665
   End
   Begin VB.Label Label3 
      Caption         =   "HH:MM:SS"
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
      Left            =   3555
      TabIndex        =   7
      Top             =   1260
      Width           =   1665
   End
   Begin VB.Label Label2 
      Caption         =   "dd/mm/yyyy"
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
      Left            =   3555
      TabIndex        =   6
      Top             =   735
      Width           =   1665
   End
   Begin VB.Label Label1 
      Caption         =   "Prueba miText"
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
      Left            =   240
      TabIndex        =   5
      Top             =   60
      Width           =   3690
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Prueba de miText
' DateTime  : 20/09/2003 21:07
' Author    : Administrador
' Purpose   :
'---------------------------------------------------------------------------------------
Option Explicit

Private Sub Command1_Click()
'miText4.Valor = "1/3/2000"

MsgBox " Texto 3: " & miText3.Text & "  Valor: " & miText3.Valor & Chr(13) & _
              " Texto 4: " & miText4.Text & "  Valor: " & miText4.Valor & Chr(13)

End Sub

'---------------------------------------------------------------------------------------
' Ejemplo de utilización de los parámetros:
'---------------------------------------------------------------------------------------

Private Sub Form_Load()

With miText3
   .Alineacion = 1  'igual que alingment
  .PermitirBlanco = False  'permitir pasar con el campo en blanco
  .dspFormat = "Currency" 'formato de visualización del a información
  .LongMaxima = 10  'igual que maxlenght
   .SoloNumeros = True  'bloquear a solo numérico y .,
End With

With miText4
   .Alineacion = 0  'igual que alingment
  .PermitirBlanco = False  'permitir pasar con el campo en blanco
  .dspFormat = "dd/mm/yyyy" 'formato de visualización del a información
  .intFormat = "yyyy/mm/dd" 'formato interno
   .LongMaxima = 10  'igual que maxlenght
   .SoloNumeros = False  'bloquear a solo numérico y .,
End With

End Sub

'---------------------------------------------------------------------------------------
'Ejemplo de como hacer una validación,
'llamar a la función: .CancelarValidacion
'para que el control cancele la validación interna.
'---------------------------------------------------------------------------------------
Private Sub miText6_Validate(Cancel As Boolean)

'No permitir fecha = 01/01/2000
With miText6

    If .Text = "01/01/2000" Then
    .CancelarValidacion
    Cancel = True
    End If
    
End With
 
End Sub


